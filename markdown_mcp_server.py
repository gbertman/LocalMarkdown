#!/usr/bin/env python3
"""
LocalMarkdown MCP Server
========================

A standalone MCP (Model Context Protocol) server that watches a local directory,
converts a wide range of file types into structured Markdown, and exposes that
Markdown to an MCP client (e.g. Claude Desktop / Claude Code) for reading and
searching.

Supported conversions
----------------------
    Documents : PDF / Word / Excel / PowerPoint / HTML  (Docling — layout &
                table aware, with OCR of scanned content)
    Media     : audio + video transcription  (faster-whisper)
    Images    : technical metadata (Pillow) + OCR (Docling) + natural-language
                description (Claude vision, or local BLIP caption offline)
    Plain     : .md / .txt passthrough, .csv -> Markdown table

Run modes (one script, three entry points)
-------------------------------------------
    serve    Run the MCP server over stdio *and* a background watcher.
                 python markdown_mcp_server.py serve --watch "D:\\Inbox"

    watch    Run only the background watcher (headless, e.g. systemd).
                 python markdown_mcp_server.py watch "/data/inbox"

    process  Convert a file/folder once and exit (drag-and-drop target).
                 python markdown_mcp_server.py process "D:\\Some Folder"

If the first argument is a path (not a sub-command) the script behaves as
``process <path>`` -- this lets a Windows shortcut / .bat pass a dropped folder
as ``%1`` and "just work".

Configuration (CLI flags override environment variables)
--------------------------------------------------------
    LM_WATCH_DIR      directory to monitor               (default: ./inbox)
    LM_OUTPUT_DIR     where .md files are written         (default: ./markdown_output)
    LM_WHISPER_MODEL  faster-whisper model size           (default: base)
    LM_OCR_LANGS      Docling OCR languages, comma list   (default: en)
    LM_IMAGE_DESCRIBE image describer: auto|ollama|openai|anthropic|blip|none
                                                          (default: auto)
    LM_OLLAMA_HOST    Ollama server URL    (default: http://localhost:11434)
    LM_OLLAMA_MODEL   Ollama vision model  (default: llama3.2-vision)
    LM_OPENAI_BASE_URL OpenAI-compatible local server (default: http://localhost:1234/v1)
    LM_OPENAI_MODEL   model name for the OpenAI-compatible server (default: local-model)
    LM_VLM_MODEL      Claude vision model for descriptions (default: claude-haiku-4-5)
    LM_VLM_TIMEOUT    seconds per VLM request             (default: 120)
    LM_LOG_LEVEL      logging level                       (default: INFO)

Heavy dependencies (Docling's layout/table models, faster-whisper, and the
optional image describer) are imported lazily, so the server starts immediately
and a clear, actionable error is raised only when an unsupported-because-
uninstalled file type is actually encountered. For image descriptions, "auto"
prefers a *local* Ollama server when one is reachable (zero API-token cost),
then Anthropic if ANTHROPIC_API_KEY is set, then a local BLIP caption; failing
all of those, descriptions are skipped. The Ollama and OpenAI-compatible
backends use only the Python standard library (no extra pip dependency).
"""

from __future__ import annotations

import argparse
import csv
import hashlib
import json
import logging
import os
import sys
import threading
import time
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Callable, Optional

# --------------------------------------------------------------------------- #
# Configuration
# --------------------------------------------------------------------------- #

def _env(name: str, default: str) -> str:
    return os.environ.get(name, default)


@dataclass
class Config:
    watch_dir: Path
    output_dir: Path
    whisper_model: str = _env("LM_WHISPER_MODEL", "base")
    ocr_langs: tuple[str, ...] = field(
        default_factory=lambda: tuple(
            s.strip() for s in _env("LM_OCR_LANGS", "en").split(",") if s.strip()
        )
    )
    # Skip files larger than this for OCR/transcription guard (bytes); 0 = no limit.
    max_bytes: int = int(_env("LM_MAX_BYTES", "0"))
    # Image-description backend: auto | ollama | openai | anthropic | blip | none.
    image_describe: str = _env("LM_IMAGE_DESCRIBE", "auto")
    # Claude vision model used when image_describe resolves to "anthropic".
    vlm_model: str = _env("LM_VLM_MODEL", "claude-haiku-4-5")
    # Local Ollama vision backend (zero API-token cost).
    ollama_host: str = _env("LM_OLLAMA_HOST", "http://localhost:11434")
    ollama_model: str = _env("LM_OLLAMA_MODEL", "llama3.2-vision")
    # Generic OpenAI-compatible local server (LM Studio / llama.cpp / vLLM).
    openai_base_url: str = _env("LM_OPENAI_BASE_URL", "http://localhost:1234/v1")
    openai_model: str = _env("LM_OPENAI_MODEL", "local-model")
    openai_api_key: str = _env("LM_OPENAI_API_KEY", os.environ.get("OPENAI_API_KEY", ""))
    # Timeout (seconds) for a single local/remote VLM request.
    vlm_timeout: int = int(_env("LM_VLM_TIMEOUT", "120"))
    # Output layout: "flat" = all .md in the output dir with mangled names;
    # "mirror" = recreate the source folder tree, each folder suffixed with
    # dir_suffix and each file written as <name>.md inside it.
    output_layout: str = _env("LM_OUTPUT_LAYOUT", "flat").lower()
    dir_suffix: str = _env("LM_DIR_SUFFIX", " md")
    # When true, move each original out of the watch dir into its output folder
    # after a successful conversion, so the inbox empties as work completes.
    consume_source: bool = _env("LM_CONSUME_INBOX", "0").lower() in ("1", "true", "yes", "on")
    # Separate archive for consumed originals. Moved files mirror the *source*
    # tree here (no dir_suffix) instead of sitting next to the .md. If unset,
    # __post_init__ defaults it to a "processed" folder beside the output dir
    # whenever consume_source is on, so raw originals never mix with the Markdown.
    archive_dir: Optional[Path] = field(default_factory=lambda: (
        Path(_env("LM_ARCHIVE_DIR", "")).expanduser().resolve()
        if _env("LM_ARCHIVE_DIR", "") else None
    ))

    def __post_init__(self) -> None:
        # Default consumed raw originals to "<output>/../processed" so they are
        # kept out of the output folder. An explicit LM_ARCHIVE_DIR still wins.
        if self.consume_source and self.archive_dir is None:
            self.archive_dir = (self.output_dir.parent / "processed").resolve()

    @classmethod
    def from_args(cls, watch: Optional[str], output: Optional[str]) -> "Config":
        watch_dir = Path(watch or _env("LM_WATCH_DIR", "./inbox")).expanduser().resolve()
        output_dir = Path(
            output or _env("LM_OUTPUT_DIR", "./markdown_output")
        ).expanduser().resolve()
        return cls(watch_dir=watch_dir, output_dir=output_dir)


# Extension groups ---------------------------------------------------------- #

PDF_EXTS = {".pdf"}
EXCEL_EXTS = {".xlsx", ".xlsm", ".xltx", ".xltm"}
WORD_EXTS = {".docx"}
PPTX_EXTS = {".pptx"}
HTML_EXTS = {".html", ".htm"}
AUDIO_EXTS = {".mp3", ".wav", ".m4a", ".flac", ".ogg", ".aac", ".wma"}
VIDEO_EXTS = {".mp4", ".mkv", ".mov", ".avi", ".webm", ".m4v", ".wmv", ".flv"}
IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".bmp", ".tiff", ".tif", ".gif", ".webp"}
TEXT_EXTS = {".txt", ".md", ".markdown", ".log", ".json", ".yaml", ".yml"}
CSV_EXTS = {".csv", ".tsv"}
# Email: .eml/.msg convert to one Markdown file each (attachments inlined);
# .pst (Outlook) and .mbox (Unix/Google Takeout) are *mailboxes* burst into many
# messages (see _process_pst / _process_mbox).
EMAIL_EXTS = {".eml", ".msg"}
PST_EXTS = {".pst"}
MBOX_EXTS = {".mbox"}

ALL_SUPPORTED = (
    PDF_EXTS | EXCEL_EXTS | WORD_EXTS | PPTX_EXTS | HTML_EXTS | AUDIO_EXTS
    | VIDEO_EXTS | IMAGE_EXTS | TEXT_EXTS | CSV_EXTS | EMAIL_EXTS | PST_EXTS
    | MBOX_EXTS
)

log = logging.getLogger("localmarkdown")


# --------------------------------------------------------------------------- #
# Lazy, cached loaders for the heavy ML stack
# --------------------------------------------------------------------------- #

class _Lazy:
    """Holds expensive singletons (Docling converter, faster-whisper model) thread-safely."""

    def __init__(self) -> None:
        self._lock = threading.Lock()
        self._docling = None
        self._docling_langs: Optional[tuple[str, ...]] = None
        self._fw_model = None
        self._fw_name: Optional[str] = None
        self._blip = None
        self._anthropic = None

    def docling(self, ocr_langs: tuple[str, ...]):
        """Return a cached Docling DocumentConverter (handles PDF/DOCX/XLSX/PPTX/images)."""
        with self._lock:
            if self._docling is None or self._docling_langs != ocr_langs:
                try:
                    from docling.document_converter import DocumentConverter
                except ImportError as exc:  # pragma: no cover - env dependent
                    raise RuntimeError(
                        "Docling is not installed. Install it with `pip install docling`. "
                        "It powers the PDF/Word/Excel/PowerPoint/image -> Markdown conversion."
                    ) from exc
                log.info("Initializing Docling (first use downloads layout + table models)...")
                # Try to pin the OCR language; fall back to defaults if the
                # pipeline-options API differs across Docling versions.
                try:
                    from docling.datamodel.base_models import InputFormat
                    from docling.datamodel.pipeline_options import (
                        EasyOcrOptions,
                        PdfPipelineOptions,
                    )
                    from docling.document_converter import PdfFormatOption

                    popts = PdfPipelineOptions()
                    popts.do_ocr = True
                    popts.do_table_structure = True
                    # Select GPU vs CPU via accelerator_options (the modern API);
                    # the old EasyOcrOptions(use_gpu=...) is deprecated and noisy.
                    _set_accelerator_device(popts, _cuda_available())
                    popts.ocr_options = EasyOcrOptions(lang=list(ocr_langs))
                    self._docling = DocumentConverter(
                        format_options={InputFormat.PDF: PdfFormatOption(pipeline_options=popts)}
                    )
                except Exception as exc:  # noqa: BLE001 - version-tolerant fallback
                    log.warning("Using default Docling configuration (%s)", exc)
                    self._docling = DocumentConverter()
                self._docling_langs = ocr_langs
            return self._docling

    def faster_whisper(self, model_name: str):
        """Return a cached faster-whisper model (CTranslate2 backend, no torch needed)."""
        with self._lock:
            if self._fw_model is None or self._fw_name != model_name:
                try:
                    from faster_whisper import WhisperModel
                except ImportError as exc:  # pragma: no cover - env dependent
                    raise RuntimeError(
                        "faster-whisper is not installed. Install it with "
                        "`pip install faster-whisper` to transcribe audio/video."
                    ) from exc
                device = "cuda" if _cuda_available() else "cpu"
                compute_type = "float16" if device == "cuda" else "int8"
                log.info(
                    "Loading faster-whisper '%s' on %s (%s); first use downloads weights...",
                    model_name, device, compute_type,
                )
                self._fw_model = WhisperModel(model_name, device=device, compute_type=compute_type)
                self._fw_name = model_name
            return self._fw_model

    def blip(self):
        """Return a cached local BLIP image-captioning pipeline (offline describer)."""
        with self._lock:
            if self._blip is None:
                try:
                    from transformers import pipeline
                except ImportError as exc:  # pragma: no cover - env dependent
                    raise RuntimeError(
                        "transformers is not installed. Install it with `pip install transformers` "
                        "for local image descriptions, or set LM_IMAGE_DESCRIBE=anthropic / none."
                    ) from exc
                device = 0 if _cuda_available() else -1
                log.info("Loading BLIP captioning model (first use downloads ~1 GB)...")
                self._blip = pipeline(
                    "image-to-text",
                    model="Salesforce/blip-image-captioning-base",
                    device=device,
                )
            return self._blip

    def anthropic_client(self):
        """Return a cached Anthropic client for high-quality Claude-vision descriptions."""
        with self._lock:
            if self._anthropic is None:
                try:
                    import anthropic
                except ImportError as exc:  # pragma: no cover - env dependent
                    raise RuntimeError(
                        "The `anthropic` SDK is not installed. Install it with `pip install anthropic` "
                        "and set ANTHROPIC_API_KEY, or set LM_IMAGE_DESCRIBE=blip for a local describer."
                    ) from exc
                self._anthropic = anthropic.Anthropic()  # reads ANTHROPIC_API_KEY from env
            return self._anthropic


def _cuda_available() -> bool:
    try:
        import torch  # type: ignore
        return bool(torch.cuda.is_available())
    except Exception:
        return False


def _set_accelerator_device(pipeline_options, use_cuda: bool) -> None:
    """Set Docling's accelerator device (replaces the deprecated EasyOcrOptions.use_gpu).

    Version-tolerant: AcceleratorOptions/AcceleratorDevice moved between modules
    across Docling releases, so try both import paths and degrade gracefully.
    To pick *which* physical GPU, set CUDA_VISIBLE_DEVICES in the environment.
    """
    try:
        try:
            from docling.datamodel.accelerator_options import (
                AcceleratorDevice,
                AcceleratorOptions,
            )
        except ImportError:
            from docling.datamodel.pipeline_options import (  # type: ignore
                AcceleratorDevice,
                AcceleratorOptions,
            )
        device = AcceleratorDevice.CUDA if use_cuda else AcceleratorDevice.CPU
        pipeline_options.accelerator_options = AcceleratorOptions(device=device)
    except Exception as exc:  # noqa: BLE001 - non-fatal; Docling auto-detects otherwise
        log.debug("Could not set Docling accelerator device explicitly: %s", exc)


LAZY = _Lazy()


# --------------------------------------------------------------------------- #
# Markdown helpers
# --------------------------------------------------------------------------- #

def _yaml_escape(value: str) -> str:
    return value.replace("\\", "\\\\").replace('"', '\\"')


def build_frontmatter(meta: dict) -> str:
    lines = ["---"]
    for key, value in meta.items():
        if isinstance(value, (list, tuple)):
            joined = ", ".join(str(v) for v in value)
            lines.append(f'{key}: "{_yaml_escape(joined)}"')
        else:
            lines.append(f'{key}: "{_yaml_escape(str(value))}"')
    lines.append("---")
    return "\n".join(lines)


def md_table(rows: list[list[str]], max_cols: int = 50) -> str:
    """Render a list of rows as a GitHub-flavoured Markdown table."""
    if not rows:
        return "_(empty)_"
    width = min(max(len(r) for r in rows), max_cols)
    norm = [[_cell(r[i] if i < len(r) else "") for i in range(width)] for r in rows]
    header = norm[0]
    body = norm[1:]
    out = ["| " + " | ".join(header) + " |", "| " + " | ".join(["---"] * width) + " |"]
    for r in body:
        out.append("| " + " | ".join(r) + " |")
    return "\n".join(out)


def _cell(value) -> str:
    if value is None:
        return ""
    text = str(value).replace("\n", " ").replace("|", "\\|").strip()
    return text


# --------------------------------------------------------------------------- #
# Converters  (each returns the Markdown *body*, frontmatter added later)
# --------------------------------------------------------------------------- #

def convert_with_docling(path: Path, cfg: Config) -> str:
    """Convert a document (PDF/DOCX/XLSX/PPTX/HTML/image) to Markdown via Docling.

    Docling provides layout-aware reading order, heading hierarchy, real table
    structure recognition, and OCR of scanned content in one pipeline.
    """
    converter = LAZY.docling(cfg.ocr_langs)
    log.info("Docling converting %s ...", path.name)
    result = converter.convert(str(path))
    md = result.document.export_to_markdown()
    return md.strip() if md and md.strip() else "_(Docling produced no content)_"


def convert_media(path: Path, cfg: Config) -> str:
    """Transcribe audio/video with faster-whisper (CTranslate2 backend)."""
    model = LAZY.faster_whisper(cfg.whisper_model)
    log.info("Transcribing %s ...", path.name)
    # vad_filter trims silence; iterating the generator is what runs the model.
    segments, info = model.transcribe(str(path), beam_size=5, vad_filter=True)
    lines = [seg.text.strip() for seg in segments if seg.text and seg.text.strip()]
    text = " ".join(lines).strip()
    prob = getattr(info, "language_probability", 0.0) or 0.0
    header = (
        f"**Detected language:** {getattr(info, 'language', 'unknown')} "
        f"(confidence {prob:.2f})\n\n**Transcript:**\n"
    )
    return header + (text if text else "_(no speech detected)_")


# -- Image description (the "vision" piece) --------------------------------- #

_VISION_PROMPT = (
    "Describe this image concisely for a searchable document index. "
    "Mention key objects, people, scene/context, any charts or diagrams, "
    "and summarize visible text. Keep it under ~120 words."
)


def describe_image(path: Path, cfg: Config) -> str:
    """Produce a natural-language description of an image.

    Backend selected by cfg.image_describe:
        auto      Local Ollama if a server is reachable (zero token cost),
                  else Anthropic if ANTHROPIC_API_KEY is set, else local BLIP,
                  else skip.
        ollama    Local Ollama vision model (e.g. llama3.2-vision) over HTTP.
        openai    Any OpenAI-compatible local server (LM Studio/llama.cpp/vLLM).
        anthropic Claude vision (cloud; uses API tokens; needs ANTHROPIC_API_KEY).
        blip      Local Salesforce/BLIP caption (offline, lightweight, no LLM).
        none      Skip description entirely.
    """
    mode = cfg.image_describe.lower()
    if mode == "auto":
        if _ollama_available(cfg):
            mode = "ollama"
        elif os.environ.get("ANTHROPIC_API_KEY"):
            mode = "anthropic"
        else:
            mode = "blip"
    if mode == "none":
        return ""
    backends = {
        "ollama": _describe_ollama,
        "openai": _describe_openai,
        "anthropic": _describe_anthropic,
        "blip": _describe_blip,
    }
    fn = backends.get(mode)
    if fn is None:
        return f"_(unknown image describer '{mode}')_"
    try:
        return fn(path, cfg)
    except Exception as exc:  # noqa: BLE001 - description is best-effort, never fatal
        log.warning("Image description failed for %s: %s", path.name, exc)
        return f"_(description unavailable: {exc})_"


def _ollama_available(cfg: Config) -> bool:
    """Quick probe so `auto` only picks Ollama when a server is actually up."""
    import urllib.request

    try:
        with urllib.request.urlopen(f"{cfg.ollama_host.rstrip('/')}/api/tags", timeout=0.75):
            return True
    except Exception:
        return False


def _describe_ollama(path: Path, cfg: Config) -> str:
    """Describe an image with a local Ollama vision model (no API tokens)."""
    import base64
    import json
    import urllib.request

    b64 = base64.standard_b64encode(path.read_bytes()).decode("ascii")
    payload = {
        "model": cfg.ollama_model,
        "messages": [{"role": "user", "content": _VISION_PROMPT, "images": [b64]}],
        "stream": False,
        "options": {"temperature": 0.2},
    }
    req = urllib.request.Request(
        f"{cfg.ollama_host.rstrip('/')}/api/chat",
        data=json.dumps(payload).encode("utf-8"),
        headers={"Content-Type": "application/json"},
    )
    log.info("Describing %s with Ollama (%s) ...", path.name, cfg.ollama_model)
    with urllib.request.urlopen(req, timeout=cfg.vlm_timeout) as resp:
        data = json.loads(resp.read().decode("utf-8"))
    return (data.get("message", {}).get("content") or "").strip() or "_(no description produced)_"


def _describe_openai(path: Path, cfg: Config) -> str:
    """Describe an image via any OpenAI-compatible /chat/completions endpoint."""
    import base64
    import json
    import urllib.request

    b64 = base64.standard_b64encode(path.read_bytes()).decode("ascii")
    data_uri = f"data:{_image_media_type(path)};base64,{b64}"
    payload = {
        "model": cfg.openai_model,
        "max_tokens": 400,
        "messages": [
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": _VISION_PROMPT},
                    {"type": "image_url", "image_url": {"url": data_uri}},
                ],
            }
        ],
    }
    headers = {"Content-Type": "application/json"}
    if cfg.openai_api_key:
        headers["Authorization"] = f"Bearer {cfg.openai_api_key}"
    req = urllib.request.Request(
        f"{cfg.openai_base_url.rstrip('/')}/chat/completions",
        data=json.dumps(payload).encode("utf-8"),
        headers=headers,
    )
    log.info("Describing %s via OpenAI-compatible endpoint (%s) ...", path.name, cfg.openai_model)
    with urllib.request.urlopen(req, timeout=cfg.vlm_timeout) as resp:
        data = json.loads(resp.read().decode("utf-8"))
    return (data["choices"][0]["message"]["content"] or "").strip() or "_(no description produced)_"


def _describe_blip(path: Path, cfg: Config) -> str:
    captioner = LAZY.blip()
    log.info("Captioning %s with BLIP ...", path.name)
    out = captioner(str(path))
    if isinstance(out, list) and out:
        return out[0].get("generated_text", "").strip() or "_(no caption produced)_"
    return "_(no caption produced)_"


def _describe_anthropic(path: Path, cfg: Config) -> str:
    import base64

    client = LAZY.anthropic_client()
    log.info("Describing %s with Claude vision (%s) ...", path.name, cfg.vlm_model)
    data = base64.standard_b64encode(path.read_bytes()).decode("ascii")
    message = client.messages.create(
        model=cfg.vlm_model,
        max_tokens=400,
        messages=[
            {
                "role": "user",
                "content": [
                    {
                        "type": "image",
                        "source": {
                            "type": "base64",
                            "media_type": _image_media_type(path),
                            "data": data,
                        },
                    },
                    {"type": "text", "text": _VISION_PROMPT},
                ],
            }
        ],
    )
    return "".join(
        block.text for block in message.content if getattr(block, "type", "") == "text"
    ).strip()


def _image_media_type(path: Path) -> str:
    return {
        ".png": "image/png",
        ".jpg": "image/jpeg",
        ".jpeg": "image/jpeg",
        ".gif": "image/gif",
        ".webp": "image/webp",
    }.get(path.suffix.lower(), "image/png")


def _describer_detail(cfg: Config) -> str:
    """Human-readable summary of which describer/model is effectively in use."""
    mode = cfg.image_describe.lower()
    if mode == "auto":
        if _ollama_available(cfg):
            mode = "ollama"
        elif os.environ.get("ANTHROPIC_API_KEY"):
            mode = "anthropic"
        else:
            mode = "blip"
    return {
        "ollama": f"ollama:{cfg.ollama_model} @ {cfg.ollama_host}",
        "openai": f"openai-compat:{cfg.openai_model} @ {cfg.openai_base_url}",
        "anthropic": f"anthropic:{cfg.vlm_model}",
        "blip": "local BLIP caption",
        "none": "disabled",
    }.get(mode, mode)


def convert_image(path: Path, cfg: Config) -> str:
    """Image -> Markdown: technical metadata + VLM description + OCR text."""
    parts: list[str] = []

    # 1. Technical properties via Pillow.
    try:
        from PIL import Image

        with Image.open(path) as img:
            parts.append("## Properties")
            parts.append(
                f"- Format: {img.format}\n"
                f"- Dimensions: {img.width} x {img.height} px\n"
                f"- Color mode: {img.mode}"
            )
    except Exception as exc:  # pragma: no cover - corrupt image etc.
        parts.append(f"_(could not read image metadata: {exc})_")

    # 2. Natural-language description (local Ollama/OpenAI VLM, Claude, or BLIP).
    description = describe_image(path, cfg)
    if description:
        parts.append("\n## Description")
        parts.append(description)

    # 3. OCR text via Docling (layout-aware, drives EasyOCR under the hood).
    parts.append("\n## Extracted text (OCR)")
    try:
        ocr = convert_with_docling(path, cfg)
        parts.append(ocr if ocr.strip() else "_(no text detected)_")
    except Exception as exc:  # noqa: BLE001
        parts.append(f"_(OCR failed: {exc})_")

    return "\n".join(parts)


def convert_text(path: Path, cfg: Config) -> str:
    raw = path.read_text(encoding="utf-8", errors="replace")
    if path.suffix.lower() in {".md", ".markdown"}:
        return raw  # already Markdown
    return f"```\n{raw}\n```"


def convert_csv(path: Path, cfg: Config) -> str:
    delimiter = "\t" if path.suffix.lower() == ".tsv" else ","
    rows: list[list[str]] = []
    with path.open("r", encoding="utf-8", errors="replace", newline="") as fh:
        for row in csv.reader(fh, delimiter=delimiter):
            rows.append(row)
            if len(rows) > 5000:  # guard against gigantic files
                rows.append(["...truncated..."])
                break
    return md_table(rows)


# -- Email (.eml / .msg) + Outlook PST -------------------------------------- #

from html.parser import HTMLParser as _HTMLParser


@dataclass
class ParsedEmail:
    """Format-neutral view of one email, shared by the .eml/.msg/.pst paths."""
    headers: dict
    body: str
    attachments: list  # list[tuple[str filename, bytes data]]


class _HTMLTextExtractor(_HTMLParser):
    """Minimal HTML -> text (stdlib only) for email bodies, no markup dependency."""

    def __init__(self) -> None:
        super().__init__()
        self._chunks: list[str] = []
        self._skip = 0

    def handle_starttag(self, tag, attrs):
        if tag in {"script", "style"}:
            self._skip += 1
        elif tag in {"br", "p", "div", "tr", "li", "h1", "h2", "h3"}:
            self._chunks.append("\n")

    def handle_endtag(self, tag):
        if tag in {"script", "style"} and self._skip:
            self._skip -= 1

    def handle_data(self, data):
        if not self._skip and data.strip():
            self._chunks.append(data)

    def text(self) -> str:
        import re
        return re.sub(r"\n{3,}", "\n\n", "".join(self._chunks)).strip()


def _html_to_text(html: str) -> str:
    parser = _HTMLTextExtractor()
    try:
        parser.feed(html)
    except Exception:  # noqa: BLE001 - never let a malformed body kill conversion
        return html
    return parser.text()


def _fmt_headers(headers: dict) -> str:
    if not headers:
        return "_(no headers)_"
    return "\n".join(f"- **{k}:** {v}" for k, v in headers.items())


def _email_body_text(msg) -> str:
    """Best body text from a parsed email.message.EmailMessage (plain preferred)."""
    plain = html = None
    parts = msg.walk() if msg.is_multipart() else [msg]
    for part in parts:
        if part.get_content_maintype() == "multipart":
            continue
        if part.get_content_disposition() == "attachment":
            continue
        try:
            content = part.get_content()
        except Exception:  # noqa: BLE001
            continue
        ctype = part.get_content_type()
        if ctype == "text/plain" and plain is None:
            plain = content
        elif ctype == "text/html" and html is None:
            html = content
    if isinstance(plain, str) and plain.strip():
        return plain.strip()
    if isinstance(html, str) and html.strip():
        return _html_to_text(html)
    return ""


def _parse_message(msg) -> ParsedEmail:
    """Build a ParsedEmail from a parsed email.message.EmailMessage (modern policy).

    Shared by the .eml, .pst (burst-to-eml) and .mbox paths so all email-like
    sources produce identical Markdown.
    """
    headers = {}
    for h in ("From", "To", "Cc", "Bcc", "Subject", "Date"):
        val = msg.get(h)
        if val:
            headers[h] = str(val).strip()
    attachments: list = []
    if hasattr(msg, "iter_attachments"):
        for part in msg.iter_attachments():
            filename = part.get_filename() or "attachment.bin"
            try:
                data = part.get_content()
            except Exception:  # noqa: BLE001
                data = part.get_payload(decode=True) or b""
            if isinstance(data, str):
                data = data.encode("utf-8", "replace")
            attachments.append((filename, bytes(data)))
    return ParsedEmail(headers=headers, body=_email_body_text(msg), attachments=attachments)


def _parse_eml(path: Path) -> ParsedEmail:
    import email
    from email import policy

    with path.open("rb") as fh:
        msg = email.message_from_binary_file(fh, policy=policy.default)
    return _parse_message(msg)


def _safe_component(text: str, maxlen: int = 60) -> str:
    """Turn an arbitrary string (e.g. a subject line) into a safe path component."""
    import re

    cleaned = re.sub(r'[\\/:*?"<>|\r\n\t]+', "_", text)
    cleaned = re.sub(r"\s+", " ", cleaned).strip().strip(".")
    return cleaned[:maxlen].strip() or "message"


def _iter_mbox(path: Path):
    """Yield (1-based index, ParsedEmail) for each message in an mbox file.

    Uses the stdlib `mailbox` module, which builds a lightweight table of contents
    and reads one message at a time — so multi-GB mailboxes (e.g. Google Takeout)
    don't have to be loaded into memory at once.
    """
    import mailbox
    from email import message_from_bytes, policy

    box = mailbox.mbox(str(path))
    try:
        for i, key in enumerate(box.keys(), start=1):
            try:
                raw = box.get_bytes(key)
                msg = message_from_bytes(raw, policy=policy.default)
                yield i, _parse_message(msg)
            except Exception:  # noqa: BLE001 - skip one bad message, keep going
                log.exception("Failed to parse message %d in %s", i, path.name)
    finally:
        box.close()


def _parse_msg(path: Path) -> ParsedEmail:
    try:
        import extract_msg
    except ImportError as exc:  # pragma: no cover - env dependent
        raise RuntimeError(
            "extract-msg is not installed. Install it with `pip install extract-msg` "
            "to convert Outlook .msg files."
        ) from exc

    m = extract_msg.Message(str(path))
    try:
        headers = {}
        for label, val in (
            ("From", m.sender), ("To", m.to), ("Cc", m.cc),
            ("Subject", m.subject), ("Date", m.date),
        ):
            if val:
                headers[label] = str(val).strip()
        body = (m.body or "").strip()
        if not body and getattr(m, "htmlBody", None):
            raw = m.htmlBody
            if isinstance(raw, (bytes, bytearray)):
                raw = raw.decode("utf-8", "replace")
            body = _html_to_text(raw)
        attachments: list = []
        for att in m.attachments:
            filename = att.longFilename or att.shortFilename or "attachment.bin"
            data = att.data
            if not isinstance(data, (bytes, bytearray, str)):
                continue  # embedded sub-message; skipped (rare)
            if isinstance(data, str):
                data = data.encode("utf-8", "replace")
            attachments.append((filename, bytes(data)))
        return ParsedEmail(headers=headers, body=body, attachments=attachments)
    finally:
        try:
            m.close()
        except Exception:  # noqa: BLE001
            pass


def _parse_email(path: Path) -> ParsedEmail:
    return _parse_msg(path) if path.suffix.lower() == ".msg" else _parse_eml(path)


def _render_email_inline(parsed: ParsedEmail) -> str:
    """Flatten a nested email (an .eml/.msg *attachment*) to text so it isn't lost."""
    out = [_fmt_headers(parsed.headers), "", parsed.body or "_(no body)_"]
    if parsed.attachments:
        names = ", ".join(n for n, _ in parsed.attachments)
        out.append(f"\n_(nested email carried {len(parsed.attachments)} attachment(s): {names})_")
    return "\n".join(out)


def convert_attachment_bytes(filename: str, data: bytes, cfg: Config) -> tuple[str, str]:
    """Convert one attachment's bytes to (kind, markdown_body) via the normal converters."""
    import tempfile

    suffix = Path(filename).suffix.lower()
    if suffix in EMAIL_EXTS:
        # A forwarded email: inline its text rather than recursing into more files.
        tmp = None
        try:
            with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as fh:
                fh.write(data)
                tmp = Path(fh.name)
            return ("email", _render_email_inline(_parse_email(tmp)))
        except Exception as exc:  # noqa: BLE001
            return ("email", f"_(nested email could not be parsed: {exc})_")
        finally:
            if tmp is not None:
                try:
                    tmp.unlink()
                except OSError:
                    pass

    mapping = converter_for(suffix)
    if mapping is None:
        return ("other", f"_(unsupported attachment type '{suffix or 'none'}', not converted)_")
    kind, convert = mapping
    tmp = None
    try:
        with tempfile.NamedTemporaryFile(suffix=suffix or ".bin", delete=False) as fh:
            fh.write(data)
            tmp = Path(fh.name)
        body = convert(tmp, cfg)
        return (kind, body if body and body.strip() else "_(no content extracted)_")
    except Exception as exc:  # noqa: BLE001 - one bad attachment must not kill the email
        log.warning("Attachment conversion failed for %s: %s", filename, exc)
        return (kind, f"_(attachment conversion failed: {exc})_")
    finally:
        if tmp is not None:
            try:
                tmp.unlink()
            except OSError:
                pass


# Dispatch table -----------------------------------------------------------> #

def converter_for(suffix: str) -> Optional[tuple[str, Callable[[Path, Config], str]]]:
    suffix = suffix.lower()
    if suffix in PDF_EXTS:
        return "pdf", convert_with_docling
    if suffix in EXCEL_EXTS:
        return "excel", convert_with_docling
    if suffix in WORD_EXTS:
        return "word", convert_with_docling
    if suffix in PPTX_EXTS:
        return "pptx", convert_with_docling
    if suffix in HTML_EXTS:
        return "html", convert_with_docling
    if suffix in AUDIO_EXTS:
        return "audio", convert_media
    if suffix in VIDEO_EXTS:
        return "video", convert_media
    if suffix in IMAGE_EXTS:
        return "image", convert_image
    if suffix in CSV_EXTS:
        return "csv", convert_csv
    if suffix in TEXT_EXTS:
        return "text", convert_text
    return None


# --------------------------------------------------------------------------- #
# Catalog / index persistence
# --------------------------------------------------------------------------- #

class Catalog:
    """Tracks which source files have been processed (by content hash)."""

    def __init__(self, output_dir: Path) -> None:
        self.output_dir = output_dir
        self.index_path = output_dir / "index.json"
        self._lock = threading.Lock()
        self._data: dict[str, dict] = {}
        self._load()

    def _load(self) -> None:
        if self.index_path.exists():
            try:
                self._data = json.loads(self.index_path.read_text(encoding="utf-8"))
            except Exception:
                log.warning("Could not parse %s, starting fresh.", self.index_path)
                self._data = {}

    def _save(self) -> None:
        tmp = self.index_path.with_suffix(".json.tmp")
        tmp.write_text(json.dumps(self._data, indent=2), encoding="utf-8")
        tmp.replace(self.index_path)

    def needs_processing(self, src: Path, digest: str) -> bool:
        with self._lock:
            entry = self._data.get(str(src))
            return not entry or entry.get("hash") != digest

    def record(self, src: Path, md_name: str, digest: str, kind: str, status: str) -> None:
        with self._lock:
            self._data[str(src)] = {
                "markdown": md_name,
                "hash": digest,
                "kind": kind,
                "status": status,
                "updated": datetime.now(timezone.utc).isoformat(timespec="seconds"),
                "source": str(src),
            }
            self._save()

    def entries(self) -> list[dict]:
        with self._lock:
            return list(self._data.values())


# --------------------------------------------------------------------------- #
# Processing pipeline
# --------------------------------------------------------------------------- #

def file_hash(path: Path, chunk: int = 1 << 20) -> str:
    h = hashlib.sha256()
    with path.open("rb") as fh:
        while True:
            block = fh.read(chunk)
            if not block:
                break
            h.update(block)
    return h.hexdigest()


def wait_until_stable(path: Path, checks: int = 3, interval: float = 0.6) -> bool:
    """Return True once the file size is stable across consecutive checks."""
    last = -1
    stable = 0
    for _ in range(checks * 5):
        try:
            size = path.stat().st_size
        except FileNotFoundError:
            return False
        if size == last:
            stable += 1
            if stable >= checks:
                return True
        else:
            stable = 0
            last = size
        time.sleep(interval)
    return True  # give up waiting, try anyway


def safe_md_name(src: Path, cfg: Config) -> str:
    """Deterministic, collision-resistant output filename."""
    rel = src
    try:
        rel = src.relative_to(cfg.watch_dir)
    except ValueError:
        rel = Path(src.name)
    slug = "_".join(rel.parts).replace(" ", "_")
    digest = hashlib.sha1(str(src).encode("utf-8")).hexdigest()[:8]
    return f"{slug}.{digest}.md"


def _source_rel(src: Path, cfg: Config) -> Path:
    """Path of a source file relative to the watch dir (or just its name if outside)."""
    try:
        return src.relative_to(cfg.watch_dir)
    except ValueError:
        return Path(src.name)


def output_md_path(src: Path, cfg: Config) -> Path:
    """Output .md path for a source file, honoring LM_OUTPUT_LAYOUT (flat|mirror).

    flat   -> <output>/<slug>.<hash>.md   (legacy; everything in one folder)
    mirror -> <output>/<dir> md/<dir> md/<name>.md   (source tree recreated)
    """
    rel = _source_rel(src, cfg)
    if cfg.output_layout == "mirror":
        dirs = [f"{p}{cfg.dir_suffix}" for p in rel.parts[:-1]]
        return cfg.output_dir.joinpath(*dirs, f"{rel.parts[-1]}.md")
    slug = "_".join(rel.parts).replace(" ", "_")
    digest = hashlib.sha1(str(src).encode("utf-8")).hexdigest()[:8]
    return cfg.output_dir / f"{slug}.{digest}.md"


def mirror_container_dir(rel: Path, cfg: Config) -> Path:
    """Folder that holds an email/PST's outputs (every rel part becomes a ' md' dir)."""
    if cfg.output_layout == "mirror":
        return cfg.output_dir.joinpath(*[f"{p}{cfg.dir_suffix}" for p in rel.parts])
    return cfg.output_dir


class Pipeline:
    def __init__(self, cfg: Config, catalog: Catalog) -> None:
        self.cfg = cfg
        self.catalog = catalog
        self.stats = {"processed": 0, "skipped": 0, "unsupported": 0, "errors": 0}
        # Persistent manifest of files we couldn't convert, so nothing is dropped
        # silently. Lives next to the Markdown so it's visible over the share.
        self._skip_path = cfg.output_dir / "_skipped.json"
        self._skip_lock = threading.Lock()
        self._skipped: dict[str, dict] = {}
        if self._skip_path.exists():
            try:
                self._skipped = json.loads(self._skip_path.read_text(encoding="utf-8"))
            except Exception:  # noqa: BLE001
                self._skipped = {}

    def _note_unsupported(self, path: Path, reason: str = "unsupported file type") -> None:
        """Record (and log, once) a file that was skipped instead of converted."""
        key = str(path)
        with self._skip_lock:
            is_new = key not in self._skipped
            self._skipped[key] = {
                "path": key,
                "ext": path.suffix.lower() or "(none)",
                "reason": reason,
                "updated": datetime.now(timezone.utc).isoformat(timespec="seconds"),
            }
            try:
                tmp = self._skip_path.with_suffix(".json.tmp")
                tmp.write_text(json.dumps(self._skipped, indent=2), encoding="utf-8")
                tmp.replace(self._skip_path)
            except OSError as exc:
                log.warning("Could not write skip manifest: %s", exc)
        if is_new:
            self.stats["unsupported"] += 1
            log.info("SKIPPED (%s): %s", reason, path)

    def skipped_entries(self) -> list[dict]:
        with self._skip_lock:
            return list(self._skipped.values())

    def _prune_empty_dirs(self, start: Path) -> None:
        """Remove now-empty source folders up to (but not including) the watch dir."""
        try:
            d = start.resolve()
        except OSError:
            return
        watch = self.cfg.watch_dir
        while d != watch and watch in d.parents:
            try:
                d.rmdir()  # only succeeds when empty
            except OSError:
                break
            d = d.parent

    def _consume_source(self, src: Path, md_dir: Path) -> None:
        """Move a successfully-converted original out of the inbox (clears it).

        Destination is LM_ARCHIVE_DIR (mirroring the source tree) when set, otherwise
        ``md_dir`` (the folder holding the generated Markdown).
        """
        if not self.cfg.consume_source:
            return
        import shutil
        if self.cfg.archive_dir:
            rel = _source_rel(src, self.cfg)
            dest_dir = self.cfg.archive_dir.joinpath(*rel.parts[:-1])
        else:
            dest_dir = md_dir
        try:
            dest_dir.mkdir(parents=True, exist_ok=True)
            dest = dest_dir / src.name
            if dest.resolve() == src.resolve():
                return
            if dest.exists():
                dest = dest_dir / f"{src.stem}.{file_hash(src)[:8]}{src.suffix}"
            shutil.move(str(src), str(dest))
            log.info("COMPLETED: moved original -> %s (inbox cleared)", dest)
            self._prune_empty_dirs(src.parent)
        except OSError as exc:
            log.warning("Could not move original %s: %s", src, exc)

    def process_file(self, src: Path, force: bool = False) -> Optional[Path]:
        src = src.resolve()
        if not src.is_file():
            return None
        # Emails and PST mailboxes don't fit the one-file->one-Markdown model:
        # an email also emits a separate, provenance-tagged file per attachment,
        # and a .pst expands into many messages. Route them to dedicated handlers.
        suffix = src.suffix.lower()
        if suffix in PST_EXTS:
            return self._process_pst(src, force=force)
        if suffix in MBOX_EXTS:
            return self._process_mbox(src, force=force)
        if suffix in EMAIL_EXTS:
            return self._process_email(src, force=force)
        mapping = converter_for(src.suffix)
        if mapping is None:
            self._note_unsupported(src)
            return None
        kind, convert = mapping

        if not wait_until_stable(src):
            log.warning("File vanished before processing: %s", src)
            return None

        try:
            digest = file_hash(src)
        except OSError as exc:
            log.warning("Cannot read %s: %s", src, exc)
            return None

        if not force and not self.catalog.needs_processing(src, digest):
            self.stats["skipped"] += 1
            log.debug("Unchanged, skipping: %s", src.name)
            md_name = self.catalog._data.get(str(src), {}).get("markdown")
            return self.cfg.output_dir / md_name if md_name else None

        if self.cfg.max_bytes and src.stat().st_size > self.cfg.max_bytes and kind in {"audio", "video", "image"}:
            log.warning("Skipping %s: exceeds LM_MAX_BYTES guard.", src.name)
            return None

        out_path = output_md_path(src, self.cfg)
        md_name = out_path.relative_to(self.cfg.output_dir).as_posix()
        log.info("Processing [%s] %s", kind, src)
        try:
            body = convert(src, self.cfg)
            status = "ok"
        except Exception as exc:
            self.stats["errors"] += 1
            log.exception("Failed to convert %s", src)
            body = f"> **Conversion failed:** `{exc}`"
            status = "error"

        frontmatter = build_frontmatter(
            {
                "source": str(src),
                "filename": src.name,
                "type": kind,
                "bytes": src.stat().st_size,
                "processed": datetime.now(timezone.utc).isoformat(timespec="seconds"),
                "hash": digest,
                "status": status,
            }
        )
        out_path.parent.mkdir(parents=True, exist_ok=True)
        out_path.write_text(f"{frontmatter}\n\n# {src.name}\n\n{body}\n", encoding="utf-8")
        self.catalog.record(src, md_name, digest, kind, status)
        if status == "ok":
            self.stats["processed"] += 1
            self._consume_source(src, out_path.parent)
        return out_path

    # -- Email / PST helpers ------------------------------------------------ #

    def _write_status_md(self, src: Path, digest: str, kind: str, status: str, body: str) -> Path:
        """Write a small placeholder Markdown so failures/notices are visible in the catalog."""
        out_path = output_md_path(src, self.cfg)
        md_name = out_path.relative_to(self.cfg.output_dir).as_posix()
        frontmatter = build_frontmatter({
            "source": str(src),
            "filename": src.name,
            "type": kind,
            "bytes": src.stat().st_size if src.exists() else 0,
            "processed": datetime.now(timezone.utc).isoformat(timespec="seconds"),
            "hash": digest,
            "status": status,
        })
        out_path.parent.mkdir(parents=True, exist_ok=True)
        out_path.write_text(f"{frontmatter}\n\n# {src.name}\n\n{body}\n", encoding="utf-8")
        self.catalog.record(src, md_name, digest, kind, status)
        return out_path

    def _emit_email_files(
        self, parsed: ParsedEmail, rel: Path, source_id: str, container: Optional[str]
    ) -> tuple[str, int, Path]:
        """Write the email body Markdown plus one provenance-tagged file per attachment.

        Returns (email_markdown_relpath, attachment_count, email_output_dir).
        In mirror layout the email gets its own ``<name> md`` folder holding the body
        (``<name>.md``) and each attachment as ``<filename>.md``; in flat layout
        everything lands in the output dir with a ``<slug>__attNN__`` naming scheme.
        Either way each attachment's frontmatter links back to the parent email.
        """
        out_dir = self.cfg.output_dir
        now = datetime.now(timezone.utc).isoformat(timespec="seconds")
        subject = parsed.headers.get("Subject", "")
        mirror = self.cfg.output_layout == "mirror"
        email_name = rel.parts[-1]

        if mirror:
            email_dir = mirror_container_dir(rel, self.cfg)
            email_md_path = email_dir / f"{email_name}.md"
        else:
            email_dir = out_dir
            base_slug = "_".join(rel.parts).replace(" ", "_")
            short_email = hashlib.sha1(source_id.encode("utf-8")).hexdigest()[:8]
            email_md_path = out_dir / f"{base_slug}.{short_email}.md"
        email_dir.mkdir(parents=True, exist_ok=True)

        used: set = set()
        att_links: list[tuple[str, str, str, int]] = []  # (filename, link_target, kind, size)
        for i, (filename, data) in enumerate(parsed.attachments, start=1):
            kind, body = convert_attachment_bytes(filename, data, self.cfg)
            if mirror:
                link_target = f"{filename}.md"
                n = 1
                while link_target in used:
                    link_target = f"{Path(filename).stem}_{n}{Path(filename).suffix}.md"
                    n += 1
                used.add(link_target)
                att_path = email_dir / link_target
            else:
                att_slug = f"{base_slug}__att{i:02d}__{filename}".replace(" ", "_")
                short = hashlib.sha1(f"{source_id}::{i}::{filename}".encode("utf-8")).hexdigest()[:8]
                link_target = f"{att_slug}.{short}.md"
                att_path = out_dir / link_target
            att_fm = build_frontmatter({
                "source": f"{source_id}#attachment:{filename}",
                "filename": filename,
                "type": kind,
                "origin": "email-attachment",
                "attachment_of": source_id,
                "email_subject": subject,
                "email_from": parsed.headers.get("From", ""),
                "email_date": parsed.headers.get("Date", ""),
                "container": container or "",
                "bytes": len(data),
                "processed": now,
                "status": "ok",
            })
            att_path.parent.mkdir(parents=True, exist_ok=True)
            att_path.write_text(
                f"{att_fm}\n\n# Attachment: {filename}\n\n"
                f"_From email: {subject or '(no subject)'}_\n\n{body}\n",
                encoding="utf-8",
            )
            self.catalog.record(
                Path(f"{source_id}#attachment:{filename}"),
                att_path.relative_to(out_dir).as_posix(), "", kind, "ok",
            )
            att_links.append((filename, link_target, kind, len(data)))

        parts = ["## Headers", _fmt_headers(parsed.headers),
                 "\n## Body", parsed.body or "_(no body text)_"]
        if att_links:
            parts.append("\n## Attachments")
            for filename, target, kind, size in att_links:
                parts.append(f"- [{filename}](<{target}>) — {kind}, {size:,} bytes")
        elif parsed.attachments:
            parts.append("\n## Attachments\n_(present but none could be converted)_")

        email_fm = build_frontmatter({
            "source": source_id,
            "filename": email_name,
            "type": "email",
            "origin": "email",
            "subject": subject,
            "from": parsed.headers.get("From", ""),
            "to": parsed.headers.get("To", ""),
            "date": parsed.headers.get("Date", ""),
            "attachments": [f for f, _, _, _ in att_links] or "",
            "container": container or "",
            "processed": now,
            "status": "ok",
        })
        email_md_path.write_text(
            f"{email_fm}\n\n# {subject or email_name}\n\n" + "\n".join(parts) + "\n",
            encoding="utf-8",
        )
        return email_md_path.relative_to(out_dir).as_posix(), len(att_links), email_dir

    def _process_email(self, src: Path, force: bool = False) -> Optional[Path]:
        """Convert a single .eml/.msg: email body + a separate file per attachment."""
        if not wait_until_stable(src):
            log.warning("File vanished before processing: %s", src)
            return None
        try:
            digest = file_hash(src)
        except OSError as exc:
            log.warning("Cannot read %s: %s", src, exc)
            return None
        if not force and not self.catalog.needs_processing(src, digest):
            self.stats["skipped"] += 1
            md_name = self.catalog._data.get(str(src), {}).get("markdown")
            return self.cfg.output_dir / md_name if md_name else None

        rel = _source_rel(src, self.cfg)

        try:
            parsed = _parse_email(src)
        except Exception as exc:  # noqa: BLE001
            self.stats["errors"] += 1
            log.exception("Failed to parse email %s", src)
            return self._write_status_md(src, digest, "email", "error",
                                         f"> **Email parse failed:** `{exc}`")

        email_md_name, n_att, email_dir = self._emit_email_files(parsed, rel, str(src), container=None)
        self.catalog.record(src, email_md_name, digest, "email", "ok")
        self.stats["processed"] += 1
        log.info("Email %s -> %s (+%d attachment file(s))", src.name, email_md_name, n_att)
        self._consume_source(src, email_dir)
        return self.cfg.output_dir / email_md_name

    def _process_pst(self, src: Path, force: bool = False) -> Optional[Path]:
        """Burst an Outlook .pst into messages with `readpst`, then convert each."""
        import shutil
        import subprocess
        import tempfile

        if not wait_until_stable(src):
            log.warning("File vanished before processing: %s", src)
            return None
        try:
            digest = file_hash(src)
        except OSError as exc:
            log.warning("Cannot read %s: %s", src, exc)
            return None
        if not force and not self.catalog.needs_processing(src, digest):
            self.stats["skipped"] += 1
            log.debug("Unchanged, skipping: %s", src.name)
            return None

        readpst = shutil.which("readpst")
        if readpst is None:
            note = ("readpst not found. Install it with `sudo apt install pst-utils` "
                    "to convert Outlook .pst mailboxes.")
            log.error(note)
            self.stats["errors"] += 1
            return self._write_status_md(src, digest, "pst", "error", f"> **{note}**")

        pst_rel = _source_rel(src, self.cfg)  # e.g. archive.pst (or sub/archive.pst)
        tmp = Path(tempfile.mkdtemp(prefix="lm_pst_"))
        messages = attachments = 0
        try:
            log.info("Bursting PST %s with readpst ...", src.name)
            subprocess.run(
                [readpst, "-e", "-D", "-o", str(tmp), str(src)],
                check=True, capture_output=True,
            )
            for eml in sorted(tmp.rglob("*.eml")):
                msg_rel = eml.relative_to(tmp)
                source_id = f"{src}!{msg_rel.as_posix()}"
                # Logical output path: <pst path>/<mailbox folders>/<message>
                rel = Path(*pst_rel.parts, *msg_rel.parts)
                try:
                    parsed = _parse_eml(eml)
                    _, n_att, _ = self._emit_email_files(parsed, rel, source_id, container=str(src))
                    messages += 1
                    attachments += n_att
                except Exception as exc:  # noqa: BLE001 - skip a bad message, keep going
                    log.exception("Failed to convert message %s from %s", msg_rel, src.name)
        except subprocess.CalledProcessError as exc:
            err = (exc.stderr or b"").decode("utf-8", "replace")[:500]
            log.error("readpst failed on %s: %s", src.name, err)
            self.stats["errors"] += 1
            return self._write_status_md(src, digest, "pst", "error",
                                         f"> **readpst failed:** `{err}`")
        finally:
            shutil.rmtree(tmp, ignore_errors=True)

        self.catalog.record(
            src, f"{messages} message(s), {attachments} attachment(s)", digest, "pst", "ok"
        )
        self.stats["processed"] += 1
        log.info("PST %s -> %d message(s), %d attachment(s).", src.name, messages, attachments)
        self._consume_source(src, mirror_container_dir(pst_rel, self.cfg))
        return None

    def _process_mbox(self, src: Path, force: bool = False) -> Optional[Path]:
        """Burst a Unix/Google-Takeout mbox mailbox into messages (stdlib `mailbox`)."""
        if not wait_until_stable(src):
            log.warning("File vanished before processing: %s", src)
            return None
        try:
            digest = file_hash(src)
        except OSError as exc:
            log.warning("Cannot read %s: %s", src, exc)
            return None
        if not force and not self.catalog.needs_processing(src, digest):
            self.stats["skipped"] += 1
            log.debug("Unchanged, skipping: %s", src.name)
            return None

        mbox_rel = _source_rel(src, self.cfg)  # e.g. archive.mbox (or sub/archive.mbox)
        messages = attachments = 0
        try:
            for i, parsed in _iter_mbox(src):
                subject = parsed.headers.get("Subject", "")
                # Each message gets a stable, sortable name under the mbox container:
                # <NNNNN>-<sanitized subject>.eml -> .../<message> md/<message>.eml.md
                msg_name = f"{i:05d}-{_safe_component(subject)}.eml"
                rel = Path(*mbox_rel.parts, msg_name)
                source_id = f"{src}!{i:05d}"
                try:
                    _, n_att, _ = self._emit_email_files(parsed, rel, source_id, container=str(src))
                    messages += 1
                    attachments += n_att
                except Exception:  # noqa: BLE001 - skip a bad message, keep going
                    log.exception("Failed to convert message %d from %s", i, src.name)
        except Exception as exc:  # noqa: BLE001 - a corrupt/locked mbox shouldn't crash the watcher
            self.stats["errors"] += 1
            log.exception("Failed to read mbox %s", src.name)
            return self._write_status_md(src, digest, "mbox", "error",
                                         f"> **mbox read failed:** `{exc}`")

        self.catalog.record(
            src, f"{messages} message(s), {attachments} attachment(s)", digest, "mbox", "ok"
        )
        self.stats["processed"] += 1
        log.info("Mbox %s -> %d message(s), %d attachment(s).", src.name, messages, attachments)
        self._consume_source(src, mirror_container_dir(mbox_rel, self.cfg))
        return None

    def process_path(self, path: Path, force: bool = False) -> int:
        """Process a single file or recurse into a folder. Returns count attempted."""
        path = path.expanduser().resolve()
        count = 0
        if path.is_file():
            self.process_file(path, force=force)
            count += 1
        elif path.is_dir():
            for child in sorted(path.rglob("*")):
                if not child.is_file():
                    continue
                # Avoid re-ingesting our own output / archive directories.
                if self.cfg.output_dir in child.parents:
                    continue
                if self.cfg.archive_dir and self.cfg.archive_dir in child.parents:
                    continue
                if child.suffix.lower() in ALL_SUPPORTED:
                    self.process_file(child, force=force)
                    count += 1
                else:
                    self._note_unsupported(child)
        else:
            log.warning("Path does not exist: %s", path)
        return count

    def initial_scan(self) -> None:
        log.info("Initial scan of %s ...", self.cfg.watch_dir)
        if self.cfg.watch_dir.exists():
            self.process_path(self.cfg.watch_dir)
        log.info("Initial scan done: %s", self.stats)


# --------------------------------------------------------------------------- #
# Watchdog integration
# --------------------------------------------------------------------------- #

def start_watcher(pipeline: Pipeline) -> "Observer":  # type: ignore[name-defined]
    from watchdog.events import FileSystemEventHandler
    from watchdog.observers import Observer

    cfg = pipeline.cfg

    class Handler(FileSystemEventHandler):
        def _handle(self, path_str: str) -> None:
            path = Path(path_str)
            resolved = path.resolve()
            if cfg.output_dir in resolved.parents or resolved == cfg.output_dir:
                return
            if cfg.archive_dir and (cfg.archive_dir in resolved.parents or resolved == cfg.archive_dir):
                return
            if path.is_dir():
                pipeline.process_path(path)
            elif path.suffix.lower() in ALL_SUPPORTED:
                pipeline.process_file(path)
            elif path.is_file():
                pipeline._note_unsupported(path)

        def on_created(self, event):
            self._handle(event.src_path)

        def on_moved(self, event):
            self._handle(event.dest_path)

        def on_modified(self, event):
            if not event.is_directory:
                self._handle(event.src_path)

    cfg.watch_dir.mkdir(parents=True, exist_ok=True)
    observer = Observer()
    observer.schedule(Handler(), str(cfg.watch_dir), recursive=True)
    observer.start()
    log.info("Watching %s (recursive) -> %s", cfg.watch_dir, cfg.output_dir)
    return observer


# --------------------------------------------------------------------------- #
# MCP server
# --------------------------------------------------------------------------- #

def run_mcp_server(pipeline: Pipeline, with_watcher: bool = True) -> None:
    try:
        from mcp.server.fastmcp import FastMCP
    except ImportError:
        log.error(
            "The MCP SDK is not installed. Run `pip install \"mcp[cli]\"` to enable "
            "the `serve` command."
        )
        sys.exit(1)

    cfg = pipeline.cfg
    catalog = pipeline.catalog
    mcp = FastMCP("LocalMarkdown")

    @mcp.tool()
    def list_documents() -> str:
        """List every processed document with its type, status and source path."""
        entries = catalog.entries()
        if not entries:
            return "No documents have been processed yet."
        lines = [f"{len(entries)} document(s) processed:\n"]
        for e in sorted(entries, key=lambda x: x.get("updated", ""), reverse=True):
            lines.append(
                f"- **{Path(e['source']).name}** "
                f"[{e['kind']}/{e['status']}] -> `{e['markdown']}`"
            )
        return "\n".join(lines)

    @mcp.tool()
    def search_documents(query: str, max_results: int = 20) -> str:
        """Full-text search across all generated Markdown. Returns matching snippets."""
        if not query.strip():
            return "Please provide a non-empty query."
        needle = query.lower()
        hits: list[str] = []
        for md in sorted(cfg.output_dir.rglob("*.md")):
            try:
                text = md.read_text(encoding="utf-8", errors="replace")
            except OSError:
                continue
            for i, line in enumerate(text.splitlines()):
                if needle in line.lower():
                    snippet = line.strip()
                    hits.append(f"- `{md.name}` (line {i + 1}): {snippet[:300]}")
                    if len(hits) >= max_results:
                        break
            if len(hits) >= max_results:
                break
        if not hits:
            return f"No matches for '{query}'."
        return f"Matches for '{query}':\n" + "\n".join(hits)

    @mcp.tool()
    def read_document(name: str) -> str:
        """Read a generated Markdown file by its name (or source filename)."""
        # Accept either the md filename or the original source name.
        candidate = cfg.output_dir / name
        if candidate.exists():
            return candidate.read_text(encoding="utf-8", errors="replace")
        for e in catalog.entries():
            if Path(e["source"]).name == name or e["markdown"] == name:
                p = cfg.output_dir / e["markdown"]
                if p.exists():
                    return p.read_text(encoding="utf-8", errors="replace")
        return f"No document found matching '{name}'. Use list_documents to see options."

    @mcp.tool()
    def process_path(path: str, force: bool = False) -> str:
        """Process a file or folder on demand. Set force=true to re-convert unchanged files."""
        count = pipeline.process_path(Path(path), force=force)
        return (
            f"Processed {count} candidate file(s) from '{path}'. "
            f"Session stats: {pipeline.stats}."
        )

    @mcp.tool()
    def list_skipped() -> str:
        """List files that were skipped (unsupported type or unconvertible), with reasons."""
        items = pipeline.skipped_entries()
        if not items:
            return "No files have been skipped."
        lines = [f"{len(items)} file(s) skipped (also in {pipeline._skip_path.name}):\n"]
        for e in sorted(items, key=lambda x: x.get("updated", ""), reverse=True):
            lines.append(f"- `{e['path']}` ({e['ext']}) — {e['reason']}")
        return "\n".join(lines)

    @mcp.tool()
    def server_status() -> str:
        """Report watcher configuration and processing statistics."""
        return (
            f"Watch dir   : {cfg.watch_dir}\n"
            f"Output dir  : {cfg.output_dir}\n"
            f"Layout      : {cfg.output_layout} (consume_source={cfg.consume_source}"
            f"{', archive=' + str(cfg.archive_dir) if cfg.archive_dir else ''})\n"
            f"Docs engine : Docling (PDF/Word/Excel/PPTX/HTML/image OCR)\n"
            f"Transcriber : faster-whisper '{cfg.whisper_model}'\n"
            f"OCR langs   : {', '.join(cfg.ocr_langs)}\n"
            f"Image desc  : {cfg.image_describe} ({_describer_detail(cfg)})\n"
            f"Documents   : {len(catalog.entries())}\n"
            f"Skipped     : {len(pipeline.skipped_entries())} (unsupported/unconvertible; see _skipped.json)\n"
            f"Stats       : {pipeline.stats}"
        )

    observer = None
    if with_watcher:
        # Run the initial scan in a thread so the MCP handshake is not blocked.
        threading.Thread(target=pipeline.initial_scan, daemon=True).start()
        observer = start_watcher(pipeline)

    try:
        log.info("Starting MCP server (stdio)...")
        mcp.run()
    finally:
        if observer is not None:
            observer.stop()
            observer.join(timeout=5)


# --------------------------------------------------------------------------- #
# CLI
# --------------------------------------------------------------------------- #

def configure_logging() -> None:
    level = getattr(logging, _env("LM_LOG_LEVEL", "INFO").upper(), logging.INFO)
    # IMPORTANT: log to stderr. stdout is reserved for the MCP stdio protocol.
    logging.basicConfig(
        level=level,
        stream=sys.stderr,
        format="%(asctime)s %(levelname)-7s %(name)s: %(message)s",
        datefmt="%H:%M:%S",
    )


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="markdown_mcp_server.py",
        description="Watch a folder, convert files to Markdown, serve them over MCP.",
    )
    sub = parser.add_subparsers(dest="command")

    p_serve = sub.add_parser("serve", help="Run MCP server + background watcher.")
    p_serve.add_argument("--watch", help="Directory to monitor.")
    p_serve.add_argument("--output", help="Directory for generated Markdown.")
    p_serve.add_argument("--no-watch", action="store_true", help="MCP only, no watcher.")

    p_watch = sub.add_parser("watch", help="Headless watcher only (e.g. for systemd).")
    p_watch.add_argument("dir", nargs="?", help="Directory to monitor.")
    p_watch.add_argument("--output", help="Directory for generated Markdown.")

    p_proc = sub.add_parser("process", help="Convert a file/folder once and exit.")
    p_proc.add_argument("paths", nargs="+", help="Files/folders to process (drag-and-drop).")
    p_proc.add_argument("--output", help="Directory for generated Markdown.")
    p_proc.add_argument("--watch-dir", help="Base watch dir (affects output naming).")
    p_proc.add_argument("--force", action="store_true", help="Re-convert unchanged files.")

    return parser


def main(argv: Optional[list[str]] = None) -> int:
    configure_logging()
    argv = list(sys.argv[1:] if argv is None else argv)

    # Drag-and-drop convenience: if the first arg is an existing path (not a known
    # sub-command), treat the whole invocation as `process <paths...>`.
    known = {"serve", "watch", "process"}
    if argv and argv[0] not in known and Path(argv[0]).exists():
        argv = ["process", *argv]

    args = build_parser().parse_args(argv)

    if args.command == "serve":
        cfg = Config.from_args(args.watch, args.output)
        cfg.output_dir.mkdir(parents=True, exist_ok=True)
        pipeline = Pipeline(cfg, Catalog(cfg.output_dir))
        run_mcp_server(pipeline, with_watcher=not args.no_watch)
        return 0

    if args.command == "watch":
        cfg = Config.from_args(args.dir, args.output)
        cfg.output_dir.mkdir(parents=True, exist_ok=True)
        pipeline = Pipeline(cfg, Catalog(cfg.output_dir))
        pipeline.initial_scan()
        observer = start_watcher(pipeline)
        try:
            while True:
                time.sleep(1)
        except KeyboardInterrupt:
            log.info("Stopping watcher...")
            observer.stop()
            observer.join(timeout=5)
        return 0

    if args.command == "process":
        watch_base = args.watch_dir
        # Default the "watch dir" to the common parent so output names stay tidy.
        cfg = Config.from_args(watch_base or str(Path(args.paths[0]).resolve().parent), args.output)
        cfg.output_dir.mkdir(parents=True, exist_ok=True)
        pipeline = Pipeline(cfg, Catalog(cfg.output_dir))
        total = 0
        for raw in args.paths:
            total += pipeline.process_path(Path(raw), force=args.force)
        log.info("Done. Attempted %d file(s). Stats: %s", total, pipeline.stats)
        log.info("Markdown written to: %s", cfg.output_dir)
        return 0

    build_parser().print_help()
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
