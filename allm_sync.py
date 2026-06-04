#!/usr/bin/env python3
"""
allm_sync.py — sync a folder of Markdown into a local AnythingLLM instance.

This is a STANDALONE, laptop-side tool. It has no dependency on LocalMarkdown
(the converter) and imports nothing from it — copy this single file to the
Windows laptop that runs AnythingLLM and run it there. Stdlib only.

What it does
------------
Watches a folder of `.md` files (e.g. a Syncthing-synced copy of the server's
`/DocConvert/output`) and pushes each one into AnythingLLM using the reliable
two-step that actually makes documents appear in the workspace UI:

    1. POST /api/v1/document/raw-text   -> creates the doc, returns its `location`
    2. POST /api/v1/workspace/<slug>/update-embeddings  {"adds": [location]}
       -> chunks + embeds it AND lists it in the workspace (this is the step
          people forget, which is why programmatic embeds "don't show up")

Using raw-text (not file upload) bypasses AnythingLLM's document collector,
removing a whole class of parse failures — your text is already clean Markdown.

It keeps a local state file mapping each source file to its content hash and the
AnythingLLM `location`, so re-runs are idempotent:
    * new file          -> embed
    * changed file      -> remove old embedding, then embed the new text
    * deleted file      -> remove its embedding (with --handle-deletes)
On startup it reconciles the whole folder against the state file, so anything
that arrived while the laptop was off/asleep gets caught up.

Quick start
-----------
    # one-shot catch-up:
    python allm_sync.py --once \
        --url http://localhost:3001 --key <ANYTHINGLLM_API_KEY> \
        --dir "C:\\Sync\\DocConvert-output" --workspace "Documents"

    # keep running, rescanning every 60s:
    python allm_sync.py --interval 60 \
        --url http://localhost:3001 --key <ANYTHINGLLM_API_KEY> \
        --dir "C:\\Sync\\DocConvert-output" --per-top-folder

    # drag-and-drop mode: embed specific files/folders, pick a workspace from a
    # numbered list (used by embed_drop.bat). --slug/--workspace skip the prompt.
    python allm_sync.py file1.md "C:\\some\\folder"      # then choose a workspace
    python allm_sync.py --list-workspaces               # just print names + slugs

Get the API key from AnythingLLM: Settings -> Tools -> Developer API -> generate.
Endpoint/field names can drift between AnythingLLM versions; confirm yours at
http://localhost:3001/api/docs and adjust the *_PATH constants below if needed.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import logging
import os
import sys
import time
import urllib.error
import urllib.request
from pathlib import Path
from typing import Optional

log = logging.getLogger("allm_sync")

# --- AnythingLLM API paths (relative to <url>/api/v1) ----------------------- #
AUTH_PATH = "/api/v1/auth"
WORKSPACES_PATH = "/api/v1/workspaces"
WORKSPACE_NEW_PATH = "/api/v1/workspace/new"
RAW_TEXT_PATH = "/api/v1/document/raw-text"
UPDATE_EMBED_PATH = "/api/v1/workspace/{slug}/update-embeddings"

# Files in the watched tree that are not documents.
IGNORE_NAMES = {"index.json", "_skipped.json"}
# Text-like files accepted in drag-and-drop mode.
DROP_EXTS = {".md", ".markdown", ".txt"}
STATE_VERSION = 1


# --------------------------------------------------------------------------- #
# HTTP helper (stdlib, with retry/backoff)
# --------------------------------------------------------------------------- #

class AnythingLLM:
    def __init__(self, base_url: str, api_key: str, timeout: int = 120, retries: int = 3):
        self.base = base_url.rstrip("/")
        self.key = api_key
        self.timeout = timeout
        self.retries = retries

    def _request(self, method: str, path: str, payload: Optional[dict] = None) -> dict:
        url = f"{self.base}{path}"
        data = json.dumps(payload).encode("utf-8") if payload is not None else None
        headers = {
            "Authorization": f"Bearer {self.key}",
            "Accept": "application/json",
        }
        if data is not None:
            headers["Content-Type"] = "application/json"

        last_err: Optional[Exception] = None
        for attempt in range(1, self.retries + 1):
            req = urllib.request.Request(url, data=data, headers=headers, method=method)
            try:
                with urllib.request.urlopen(req, timeout=self.timeout) as resp:
                    body = resp.read().decode("utf-8", "replace")
                    return json.loads(body) if body.strip() else {}
            except urllib.error.HTTPError as exc:
                detail = exc.read().decode("utf-8", "replace")[:300]
                # 4xx is a real client error — don't retry, surface it.
                if 400 <= exc.code < 500:
                    raise RuntimeError(f"{method} {path} -> HTTP {exc.code}: {detail}") from exc
                last_err = RuntimeError(f"HTTP {exc.code}: {detail}")
            except (urllib.error.URLError, TimeoutError, ConnectionError) as exc:
                last_err = exc
            if attempt < self.retries:
                backoff = 2 ** (attempt - 1)
                log.warning("%s %s failed (%s); retrying in %ss", method, path, last_err, backoff)
                time.sleep(backoff)
        raise RuntimeError(f"{method} {path} failed after {self.retries} attempts: {last_err}")

    # -- typed wrappers ----------------------------------------------------- #

    def verify(self) -> bool:
        try:
            self._request("GET", AUTH_PATH)
            return True
        except Exception as exc:  # noqa: BLE001
            log.error("Auth check failed: %s", exc)
            return False

    def list_workspaces(self) -> list[dict]:
        out = self._request("GET", WORKSPACES_PATH)
        return out.get("workspaces", []) if isinstance(out, dict) else []

    def create_workspace(self, name: str) -> str:
        out = self._request("POST", WORKSPACE_NEW_PATH, {"name": name})
        ws = out.get("workspace") or {}
        slug = ws.get("slug")
        if not slug:
            raise RuntimeError(f"create_workspace returned no slug: {out}")
        return slug

    def push_raw_text(self, text: str, title: str, source: str) -> str:
        """Create a document from raw text. Returns its store `location`."""
        payload = {
            "textContent": text,
            "metadata": {"title": title, "docSource": source},
        }
        out = self._request("POST", RAW_TEXT_PATH, payload)
        docs = out.get("documents") or []
        if not docs or not docs[0].get("location"):
            raise RuntimeError(f"raw-text returned no document location: {str(out)[:300]}")
        return docs[0]["location"]

    def update_embeddings(self, slug: str, adds=None, deletes=None) -> dict:
        payload = {"adds": adds or [], "deletes": deletes or []}
        return self._request("POST", UPDATE_EMBED_PATH.format(slug=slug), payload)


# --------------------------------------------------------------------------- #
# Workspace resolution (create-or-get, cached)
# --------------------------------------------------------------------------- #

class Workspaces:
    def __init__(self, client: AnythingLLM):
        self.client = client
        self._by_name: dict[str, str] = {}
        for ws in client.list_workspaces():
            if ws.get("name") and ws.get("slug"):
                self._by_name[ws["name"].lower()] = ws["slug"]

    def slug_for(self, name: str) -> str:
        key = name.lower()
        if key not in self._by_name:
            log.info("Creating workspace '%s'", name)
            self._by_name[key] = self.client.create_workspace(name)
        return self._by_name[key]


# --------------------------------------------------------------------------- #
# Local state (idempotency catalog)
# --------------------------------------------------------------------------- #

class State:
    def __init__(self, path: Path):
        self.path = path
        self.data: dict[str, dict] = {}
        if path.exists():
            try:
                loaded = json.loads(path.read_text(encoding="utf-8"))
                self.data = loaded.get("files", {})
            except Exception:  # noqa: BLE001
                log.warning("Could not read state %s; starting fresh", path)

    def save(self) -> None:
        tmp = self.path.with_suffix(".tmp")
        tmp.write_text(
            json.dumps({"version": STATE_VERSION, "files": self.data}, indent=2),
            encoding="utf-8",
        )
        tmp.replace(self.path)

    def get(self, rel: str) -> Optional[dict]:
        return self.data.get(rel)

    def put(self, rel: str, digest: str, location: str, slug: str) -> None:
        self.data[rel] = {"hash": digest, "location": location, "workspace": slug}

    def drop(self, rel: str) -> None:
        self.data.pop(rel, None)


# --------------------------------------------------------------------------- #
# Sync logic
# --------------------------------------------------------------------------- #

def file_hash(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as fh:
        for block in iter(lambda: fh.read(1 << 20), b""):
            h.update(block)
    return h.hexdigest()


def workspace_name_for(rel: Path, args) -> str:
    """Single workspace, or one per top-level folder (with optional suffix strip)."""
    default = args.workspace or "Documents"
    if not args.per_top_folder:
        return default
    top = rel.parts[0] if len(rel.parts) > 1 else default
    if args.strip_suffix and top.endswith(args.strip_suffix):
        top = top[: -len(args.strip_suffix)].strip()
    return top or default


def iter_markdown(root: Path):
    for p in sorted(root.rglob("*.md")):
        if p.is_file() and p.name not in IGNORE_NAMES:
            yield p


def _embed_one(client: AnythingLLM, slug: str, path: Path, state: State, stats: dict,
               *, state_key: str, title: str, source: str) -> None:
    """Embed (or re-embed) a single file into one workspace, idempotently."""
    try:
        digest = file_hash(path)
    except OSError as exc:
        log.warning("Cannot read %s: %s", path, exc)
        stats["errors"] += 1
        return

    prev = state.get(state_key)
    if prev and prev.get("hash") == digest and prev.get("workspace") == slug:
        stats["unchanged"] += 1
        return
    try:
        # If a previous version was embedded, remove its stale chunks first.
        if prev and prev.get("location") and prev.get("workspace"):
            client.update_embeddings(prev["workspace"], deletes=[prev["location"]])

        text = path.read_text(encoding="utf-8", errors="replace")
        location = client.push_raw_text(text, title=title, source=source)
        client.update_embeddings(slug, adds=[location])
        state.put(state_key, digest, location, slug)
        state.save()  # save incrementally so a crash mid-run doesn't redo everything
        if prev:
            stats["updated"] += 1
            log.info("Updated  %s -> %s", title, slug)
        else:
            stats["added"] += 1
            log.info("Embedded %s -> %s", title, slug)
    except Exception as exc:  # noqa: BLE001
        stats["errors"] += 1
        log.error("Failed on %s: %s", title, exc)


def collect_drop_files(paths) -> list:
    """Expand dropped files/folders into a de-duplicated list of text files."""
    out: list = []
    seen: set = set()
    for raw in paths:
        p = Path(raw).expanduser()
        candidates = []
        if p.is_dir():
            candidates = [c for c in sorted(p.rglob("*"))
                          if c.is_file() and c.suffix.lower() in DROP_EXTS
                          and c.name not in IGNORE_NAMES]
        elif p.is_file() and p.suffix.lower() in DROP_EXTS:
            candidates = [p]
        else:
            log.warning("Skipping (not a .md/.markdown/.txt file or folder): %s", raw)
        for c in candidates:
            rp = c.resolve()
            if rp not in seen:
                seen.add(rp)
                out.append(c)
    return out


def choose_workspace_interactive(client: AnythingLLM, workspaces: Workspaces) -> Optional[str]:
    """Print a numbered list of workspace slugs and return the chosen (or new) slug."""
    wss = client.list_workspaces()
    print("\nAvailable AnythingLLM workspaces:", flush=True)
    if wss:
        for i, ws in enumerate(wss, 1):
            print(f"  {i}) {ws.get('name', '?'):<32} [{ws.get('slug', '?')}]", flush=True)
    else:
        print("  (none yet)", flush=True)
    print("  Or type a NEW workspace name to create it.", flush=True)
    try:
        choice = input("\nPick a number, or enter a workspace name: ").strip()
    except EOFError:
        return None
    if not choice:
        return None
    if choice.isdigit():
        idx = int(choice) - 1
        if 0 <= idx < len(wss):
            return wss[idx].get("slug")
        print("Invalid selection.", flush=True)
        return None
    return workspaces.slug_for(choice)  # treat as a (possibly new) workspace name


def embed_paths(client: AnythingLLM, workspaces: Workspaces, state: State, args) -> dict:
    """Drag-and-drop mode: embed the given files/folders into one chosen workspace."""
    files = collect_drop_files(args.paths)
    stats = {"added": 0, "updated": 0, "unchanged": 0, "errors": 0}
    if not files:
        log.error("No .md/.markdown/.txt files found in the dropped items.")
        stats["errors"] += 1
        return stats

    if args.slug:
        slug = args.slug
    elif args.workspace:
        slug = workspaces.slug_for(args.workspace)
    else:
        slug = choose_workspace_interactive(client, workspaces)
    if not slug:
        log.error("No workspace selected; nothing embedded.")
        stats["errors"] += 1
        return stats

    print(f"\nEmbedding {len(files)} file(s) into workspace '{slug}' ...\n", flush=True)
    for path in files:
        abspath = str(path.resolve())
        # Key by (workspace, file) so the same file can live in multiple workspaces.
        _embed_one(client, slug, path, state, stats,
                   state_key=f"{slug}::{abspath}", title=path.name, source=abspath)
    return stats


def sync_once(client: AnythingLLM, workspaces: Workspaces, state: State, args) -> dict:
    root = Path(args.dir).expanduser().resolve()
    stats = {"added": 0, "updated": 0, "deleted": 0, "unchanged": 0, "errors": 0}
    seen: set[str] = set()

    for path in iter_markdown(root):
        rel = path.relative_to(root)
        rel_key = rel.as_posix()
        seen.add(rel_key)
        slug = workspaces.slug_for(workspace_name_for(rel, args))
        _embed_one(client, slug, path, state, stats,
                   state_key=rel_key, title=rel_key, source=rel_key)

    # Handle deletions: files in state that no longer exist on disk. Skip
    # drag-and-drop entries (keyed "slug::abspath") — only watch-mode rel keys.
    if args.handle_deletes:
        for rel_key in list(state.data.keys()):
            if "::" in rel_key or rel_key in seen:
                continue
            entry = state.data[rel_key]
            try:
                client.update_embeddings(entry["workspace"], deletes=[entry["location"]])
                state.drop(rel_key)
                state.save()
                stats["deleted"] += 1
                log.info("Removed  %s (source gone)", rel_key)
            except Exception as exc:  # noqa: BLE001
                stats["errors"] += 1
                log.error("Failed to remove %s: %s", rel_key, exc)

    return stats


# --------------------------------------------------------------------------- #
# CLI
# --------------------------------------------------------------------------- #

def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        description="Embed Markdown into a local AnythingLLM — watch a folder, or "
                    "drag-and-drop files and pick a workspace.")
    p.add_argument("paths", nargs="*",
                   help="Drag-and-drop mode: files/folders to embed (pick a workspace "
                        "interactively unless --slug/--workspace is given). If omitted, "
                        "runs watch mode over --dir.")
    p.add_argument("--url", default=os.environ.get("ALLM_URL", "http://localhost:3001"),
                   help="AnythingLLM base URL (default: http://localhost:3001 or $ALLM_URL).")
    p.add_argument("--key", default=os.environ.get("ALLM_KEY", ""),
                   help="AnythingLLM Developer API key (or $ALLM_KEY).")
    p.add_argument("--list-workspaces", action="store_true",
                   help="Print the available workspace names + slugs and exit.")
    p.add_argument("--dir", default=os.environ.get("ALLM_DIR", "."),
                   help="Watch mode: folder of .md files to sync (the synced copy of output/).")
    p.add_argument("--workspace", default=os.environ.get("ALLM_WORKSPACE") or None,
                   help="Target workspace NAME (created if missing). Drop mode: skips the menu.")
    p.add_argument("--slug",
                   help="Target workspace SLUG directly (drop mode; skips the menu).")
    p.add_argument("--per-top-folder", action="store_true",
                   help="Watch mode: route each file to a workspace named after its top-level folder.")
    p.add_argument("--strip-suffix", default=" md",
                   help="Suffix to strip from top-folder names for workspace naming (default: ' md').")
    p.add_argument("--state", default=os.environ.get("ALLM_STATE", "allm_sync_state.json"),
                   help="Path to the idempotency state file (keep it OUTSIDE the synced folder).")
    p.add_argument("--handle-deletes", action="store_true",
                   help="Watch mode: remove embeddings when their source .md disappears.")
    p.add_argument("--interval", type=int, default=0,
                   help="Watch mode: rescan every N seconds (0 = run once and exit).")
    p.add_argument("--once", action="store_true", help="Watch mode: run a single pass and exit.")
    p.add_argument("--timeout", type=int, default=120, help="Per-request timeout (seconds).")
    return p


def main(argv=None) -> int:
    logging.basicConfig(
        level=logging.INFO, stream=sys.stderr,
        format="%(asctime)s %(levelname)-7s %(message)s", datefmt="%H:%M:%S",
    )
    args = build_parser().parse_args(argv)
    if not args.key:
        log.error("No API key. Pass --key or set ALLM_KEY. (AnythingLLM: Settings -> Developer API.)")
        return 2

    client = AnythingLLM(args.url, args.key, timeout=args.timeout)
    if not client.verify():
        log.error("Could not authenticate to AnythingLLM at %s. Check --url and --key.", args.url)
        return 2

    workspaces = Workspaces(client)
    state = State(Path(args.state).expanduser().resolve())

    # --- List workspaces and exit ----------------------------------------- #
    if args.list_workspaces:
        wss = client.list_workspaces()
        if not wss:
            print("No workspaces yet.")
        for ws in wss:
            print(f"  {ws.get('name', '?'):<32} [{ws.get('slug', '?')}]")
        return 0

    # --- Drag-and-drop mode ----------------------------------------------- #
    if args.paths:
        stats = embed_paths(client, workspaces, state, args)
        log.info("Done: %s", stats)
        print(f"\nDone: {stats}", flush=True)
        return 0 if stats["errors"] == 0 else 1

    # --- Watch mode ------------------------------------------------------- #
    root = Path(args.dir).expanduser().resolve()
    if not root.is_dir():
        log.error("Watch dir does not exist: %s (pass --dir, or drop files to embed)", root)
        return 2
    log.info("Syncing %s -> %s (state: %s)", root, args.url, state.path)

    while True:
        try:
            stats = sync_once(client, workspaces, state, args)
            log.info("Pass complete: %s", stats)
        except KeyboardInterrupt:
            log.info("Stopped.")
            return 0
        except Exception as exc:  # noqa: BLE001 - keep the loop alive
            log.exception("Sync pass failed: %s", exc)
        if args.interval <= 0 or args.once:
            return 0
        time.sleep(args.interval)


if __name__ == "__main__":
    raise SystemExit(main())
