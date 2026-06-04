# LocalMarkdown MCP Server

A standalone Python **MCP (Model Context Protocol) server** that watches a local
directory, converts a wide range of file types into structured **Markdown**, and
exposes that Markdown to Claude (Desktop / Code) so it can read and search your
documents, media and images.

| Category   | Handled by                         | Extensions                                  |
|------------|------------------------------------|---------------------------------------------|
| Documents  | **Docling** (layout + tables + OCR)| `.pdf .docx .xlsx .pptx .html .htm`         |
| Audio      | **faster-whisper**                 | `.mp3 .wav .m4a .flac .ogg .aac .wma`       |
| Video      | **faster-whisper**                 | `.mp4 .mkv .mov .avi .webm .m4v .wmv .flv`  |
| Images     | Pillow + Docling OCR + VLM caption  | `.png .jpg .jpeg .bmp .tiff .gif .webp`     |
| Plain/data | built-in                           | `.txt .md .log .json .yaml`, `.csv .tsv`    |
| Email      | stdlib / extract-msg / readpst     | `.eml .msg .pst` (attachments converted to their own files) |

Image files get three layers: technical **properties** (Pillow), a natural-language
**description** from a vision model (Claude vision, or a local BLIP caption when
offline), and **OCR text** (Docling). See [§2.1 Image descriptions](#21-image-descriptions).

Each converted file becomes a Markdown file with YAML frontmatter (source path,
type, hash, timestamp) in the output directory, catalogued in `index.json`.

---

## 1. How it works

```
watched folder ──► watchdog / drag-drop ──► converter (by extension)
                                              │
                                              ▼
                       markdown_output/*.md  +  index.json
                                              │
                          MCP tools ◄──────────┘
        list_documents · search_documents · read_document
                  · process_path · server_status
```

Three run modes from one script:

| Command   | Purpose                                                  |
|-----------|----------------------------------------------------------|
| `serve`   | MCP server over stdio **plus** a live background watcher. |
| `watch`   | Headless watcher only (for a Linux systemd service).      |
| `process` | Convert a file/folder once and exit (drag-and-drop).      |

---

## 2. Quick start (any OS)

```bash
# 1. create an isolated environment
python -m venv .venv
# Windows:  .venv\Scripts\activate
# Linux/Mac: source .venv/bin/activate

# 2. install dependencies
pip install -r requirements.txt

# 3. install the ffmpeg binary (required by Whisper for audio/video)
#    Windows:  winget install Gyan.FFmpeg
#    Ubuntu :  sudo apt install ffmpeg
#    macOS  :  brew install ffmpeg

# 4. try a one-shot conversion
python markdown_mcp_server.py process "./some_folder"

# 5. run the MCP server + watcher
python markdown_mcp_server.py serve --watch "./inbox" --output "./markdown_output"
```

> The first document/image/media you process will **download model weights**
> (Docling layout + table models; faster-whisper ~140 MB for `base`; EasyOCR
> ~64 MB per OCR language; BLIP ~1 GB only if you use the local describer). This
> is a one-time cost — subsequent runs are fast.

### Configuration

CLI flags override environment variables:

| Env var             | Flag        | Default              | Meaning                         |
|---------------------|-------------|----------------------|---------------------------------|
| `LM_WATCH_DIR`      | `--watch`   | `./inbox`            | Folder to monitor               |
| `LM_OUTPUT_DIR`     | `--output`  | `./markdown_output`  | Where `.md` files are written   |
| `LM_OUTPUT_LAYOUT`  | —           | `flat`               | `flat` = all `.md` in one folder; `mirror` = recreate the source tree, each folder suffixed (see `LM_DIR_SUFFIX`), files written as `<name>.md` |
| `LM_DIR_SUFFIX`     | —           | ` md`                | In `mirror` layout, appended to every recreated folder name (e.g. `Smith` → `Smith md`) |
| `LM_CONSUME_INBOX`  | —           | `0`                  | `1` = **move** each original into its output folder after a successful convert, so the inbox empties (empty source dirs are pruned) |
| `LM_WHISPER_MODEL`  | —           | `base`               | faster-whisper size: `tiny`/`base`/`small`/`medium`/`large-v3` |
| `LM_OCR_LANGS`      | —           | `en`                 | Docling OCR languages, e.g. `en,fr` |
| `LM_IMAGE_DESCRIBE` | —           | `auto`               | `auto`/`ollama`/`openai`/`anthropic`/`blip`/`none` (see §2.1) |
| `LM_OLLAMA_HOST`    | —           | `http://localhost:11434` | Local Ollama server URL    |
| `LM_OLLAMA_MODEL`   | —           | `llama3.2-vision`    | Ollama vision model         |
| `LM_OPENAI_BASE_URL`| —           | `http://localhost:1234/v1` | OpenAI-compatible local server |
| `LM_OPENAI_MODEL`   | —           | `local-model`        | Model name for that server  |
| `LM_VLM_MODEL`      | —           | `claude-haiku-4-5`   | Claude vision model (cloud)  |
| `LM_VLM_TIMEOUT`    | —           | `120`                | Seconds per VLM request      |
| `LM_MAX_BYTES`      | —           | `0` (no limit)       | Skip OCR/transcribe above N bytes |
| `LM_LOG_LEVEL`      | —           | `INFO`               | `DEBUG`/`INFO`/`WARNING`        |

### 2.1 Image descriptions

OCR only extracts *text*; a real *description* of what an image shows needs a
vision model. **Use a local LLM to pay zero API tokens** — the `ollama` and
`openai` backends talk to a local server over HTTP and need no extra Python
package. The backend is chosen by `LM_IMAGE_DESCRIBE`:

| Value       | Cost | Behaviour                                                          |
|-------------|------|-------------------------------------------------------------------|
| `auto` (default) | — | Local **Ollama** if a server is reachable, else **Claude** if `ANTHROPIC_API_KEY` is set, else local **BLIP**, else skip. |
| `ollama`    | **free** | Local **Ollama** vision model (e.g. `llama3.2-vision`). No tokens, no pip dep. |
| `openai`    | **free** | Any **OpenAI-compatible** local server: LM Studio, `llama.cpp --server`, vLLM. |
| `anthropic` | tokens | Cloud **Claude vision**. Best quality. `pip install anthropic` + `ANTHROPIC_API_KEY`. |
| `blip`      | free | Local **BLIP** caption (no LLM, basic). `pip install transformers` (~1 GB). |
| `none`      | free | Skip descriptions; images still get properties + OCR text.        |

```powershell
# Recommended: local LLM via Ollama (no API tokens)
#   1) install Ollama from https://ollama.com
#   2) ollama pull llama3.2-vision
$env:LM_IMAGE_DESCRIBE = "ollama"
$env:LM_OLLAMA_MODEL   = "llama3.2-vision"

# Any OpenAI-compatible local server (e.g. LM Studio on :1234)
$env:LM_IMAGE_DESCRIBE  = "openai"
$env:LM_OPENAI_BASE_URL = "http://localhost:1234/v1"
$env:LM_OPENAI_MODEL    = "qwen2.5-vl-7b"
```

> **Privacy:** `ollama`, `openai` (local), `blip`, and `none` keep every image on
> your machine. Only `anthropic` uploads the image to the cloud.
> Description failures are non-fatal — the image still gets its OCR + properties.

#### Which local vision model? (GPU sizing)

The describer runs **one image at a time**, so VRAM, not throughput, is the
constraint. Pull any of these with `ollama pull <name>`:

| GPU VRAM | Recommended Ollama model | Notes |
|----------|--------------------------|-------|
| 6–8 GB   | `moondream` (1.8B) or `llava:7b` (Q4) | Light, fast, basic descriptions. |
| 10–12 GB | `llama3.2-vision:11b` (Q4) / `qwen2.5-vl:7b` | Good general captioning + reads text well. |
| **24 GB** | **`qwen2.5-vl:7b` (full/`q8`) for speed, or `llama3.2-vision:11b` (q8); for max quality `qwen2.5-vl:32b` (Q4) or `gemma3:27b`** | 24 GB comfortably runs an 11B at high precision, or a 27–32B quantized. See §2.2. |
| 48 GB+   | `llama3.2-vision:90b` (Q4) / `qwen2.5-vl:72b` (Q4) | Near-cloud quality, fully local. |

### 2.2 Best model for a 24 GB card (e.g. RTX 3090/4090, A5000)

24 GB is a sweet spot. Two good strategies for *this* task (concise image
descriptions + reading on-image text):

- **Best quality:** `qwen2.5-vl:32b` at Q4 (~20 GB) — excellent at charts,
  documents and dense text. Or `gemma3:27b` if you prefer its style. Both fit
  with headroom for the image tokens.
- **Best speed/quality balance (recommended):** `qwen2.5-vl:7b` at `q8` or
  `llama3.2-vision:11b` at `q8` (~12–13 GB). Fast, leaves VRAM free, and is more
  than enough for index-quality captions. Qwen2.5-VL is particularly strong at
  OCR-style text reading, which pairs well with this pipeline.

```powershell
ollama pull qwen2.5-vl:7b          # fast, ~8–9 GB, great text reading
# or, for maximum local quality on 24 GB:
ollama pull qwen2.5-vl:32b         # Q4 ~20 GB

$env:LM_IMAGE_DESCRIBE = "ollama"
$env:LM_OLLAMA_MODEL   = "qwen2.5-vl:7b"
```

> Tip: if you already run the **faster-whisper** transcription on the same GPU,
> stick to an 11B-class describer so both fit; Ollama loads the vision model
> on demand and unloads it after idle, so they rarely contend in practice.

---

## 3. Connecting it to Claude

The server speaks MCP over **stdio**. Add it to your client config.

**Claude Desktop** — edit `claude_desktop_config.json`
(`%APPDATA%\Claude\` on Windows, `~/Library/Application Support/Claude/` on macOS):

```json
{
  "mcpServers": {
    "localmarkdown": {
      "command": "C:\\Users\\you\\github\\LocalMarkdown\\.venv\\Scripts\\python.exe",
      "args": [
        "C:\\Users\\you\\github\\LocalMarkdown\\markdown_mcp_server.py",
        "serve",
        "--watch",  "C:\\Users\\you\\Inbox",
        "--output", "C:\\Users\\you\\github\\LocalMarkdown\\markdown_output"
      ]
    }
  }
}
```

**Claude Code (CLI):**

```bash
claude mcp add localmarkdown -- \
  /path/to/.venv/bin/python /path/to/markdown_mcp_server.py serve --watch /path/to/inbox
```

Restart the client, then ask Claude things like *“search my documents for the Q3
budget”* or *“read the transcript of meeting.mp4”*. Claude calls the
`search_documents` / `read_document` tools automatically.

> **Protocol note:** the server writes **all logs to stderr** and keeps stdout
> clean for the MCP JSON-RPC stream — don’t add `print()` calls to stdout.

---

## 4. Windows drag-and-drop shortcut

1. Right-click `drag_and_drop.bat` → **Create shortcut**.
2. Move the shortcut to your Desktop (rename it e.g. *“➜ Markdown”*).
3. **Drag any file or folder onto the shortcut.** It runs
   `markdown_mcp_server.py process <dropped paths>` and writes Markdown into
   `markdown_output\`. Multiple items can be dropped at once.

The `.bat` automatically uses `.venv\Scripts\python.exe` if that virtual
environment exists, otherwise the system `python`.

---

## 5. Installing on Windows with **WSL2** (recommended for the ML stack)

PyTorch (used by Docling) is happier on Linux, and WSL2 gives you optional
NVIDIA GPU acceleration. The plan: run the **converter/watcher inside WSL2**, and
point it at a folder that’s also visible from Windows.

### 5.1 Enable WSL2 and install Ubuntu

Open **PowerShell as Administrator**:

```powershell
wsl --install -d Ubuntu
wsl --set-default-version 2
```

Reboot if prompted, then launch **Ubuntu** from the Start menu and create your
Linux username/password. Verify you’re on version 2:

```powershell
wsl -l -v        # STATE Running, VERSION 2
```

### 5.2 Install system packages inside Ubuntu

```bash
sudo apt update && sudo apt upgrade -y
sudo apt install -y python3 python3-venv python3-pip ffmpeg git \
                    libgl1 libglib2.0-0          # libGL/glib needed by EasyOCR/OpenCV
```

### 5.3 Get the project and create the environment

```bash
# Option A: copy the project from your Windows drive (auto-mounted at /mnt/c)
cp -r /mnt/c/Users/you/github/LocalMarkdown ~/LocalMarkdown
cd ~/LocalMarkdown

python3 -m venv .venv
source .venv/bin/activate
pip install --upgrade pip
pip install -r requirements.txt
```

> **CPU-only / smaller install:** install the CPU PyTorch wheel *before* the
> requirements to avoid the multi-GB CUDA download:
> ```bash
> pip install torch torchvision --index-url https://download.pytorch.org/whl/cpu
> pip install -r requirements.txt
> ```

### 5.4 Choosing the watch folder (Windows ⇄ WSL2)

- **Watch a Windows folder from WSL2:** point `--watch` at the mounted path, e.g.
  `--watch /mnt/c/Users/you/Inbox`. Anything you drop into `C:\Users\you\Inbox`
  in Explorer is picked up.
  *Caveat:* `inotify` events across the `/mnt/c` boundary can be unreliable; if
  the watcher misses files, run the converter from inside Linux-native storage
  (next bullet) or trigger conversions with the `process` command.
- **Best reliability:** keep the inbox on the Linux filesystem
  (`~/Inbox`), which you can still open from Windows Explorer via
  `\\wsl$\Ubuntu\home\you\Inbox`.

### 5.5 Run it

```bash
# headless watcher
python markdown_mcp_server.py watch /mnt/c/Users/you/Inbox \
       --output ~/LocalMarkdown/markdown_output

# or the full MCP server (if your MCP client launches the WSL python)
python markdown_mcp_server.py serve --watch ~/Inbox
```

To launch a **WSL** Python from a **Windows** Claude Desktop config:

```json
{
  "mcpServers": {
    "localmarkdown": {
      "command": "wsl.exe",
      "args": ["-d", "Ubuntu", "--",
               "/home/you/LocalMarkdown/.venv/bin/python",
               "/home/you/LocalMarkdown/markdown_mcp_server.py",
               "serve", "--watch", "/home/you/Inbox"]
    }
  }
}
```

### 5.6 (Optional) GPU acceleration in WSL2

Install the latest **NVIDIA Windows driver** (it ships the WSL CUDA stack — do
**not** install a Linux GPU driver inside WSL). Then install a CUDA PyTorch wheel:

```bash
pip install torch torchvision --index-url https://download.pytorch.org/whl/cu121
python -c "import torch; print(torch.cuda.is_available())"   # -> True
```

The script auto-detects CUDA and enables GPU for Docling/EasyOCR and
faster-whisper (which switches to `float16` on GPU).

---

## 6. Running in the background on a **Linux server**

For an always-on converter on a headless box, run the **`watch`** mode under
**systemd**.

### 6.1 Install

```bash
sudo useradd -r -m -d /opt/localmarkdown -s /usr/sbin/nologin localmd  # service user
sudo -u localmd git clone <your-repo> /opt/localmarkdown/app   # or copy the files
cd /opt/localmarkdown/app

sudo apt install -y python3-venv ffmpeg libgl1 libglib2.0-0 pst-utils  # pst-utils only needed for .pst mailboxes
sudo -u localmd python3 -m venv /opt/localmarkdown/venv
sudo -u localmd /opt/localmarkdown/venv/bin/pip install -r requirements.txt

sudo -u localmd mkdir -p /opt/localmarkdown/inbox /opt/localmarkdown/out
```

### 6.2 Create the service unit

`/etc/systemd/system/localmarkdown.service`:

```ini
[Unit]
Description=LocalMarkdown — folder-to-Markdown watcher
After=network-online.target
Wants=network-online.target

[Service]
Type=simple
User=localmd
Group=localmd
WorkingDirectory=/opt/localmarkdown/app

# Configuration via environment
Environment=LM_WHISPER_MODEL=base
Environment=LM_OCR_LANGS=en
Environment=LM_LOG_LEVEL=INFO
# Image descriptions (pick one):
#   "ollama" = local LLM, no API tokens (run `ollama serve` + pull a vision model);
#   "blip"   = fully offline caption (pip install transformers);
#   "anthropic" = cloud Claude vision (set ANTHROPIC_API_KEY);
#   "none"   = skip descriptions.
Environment=LM_IMAGE_DESCRIBE=ollama
Environment=LM_OLLAMA_HOST=http://localhost:11434
Environment=LM_OLLAMA_MODEL=llama3.2-vision
# Environment=ANTHROPIC_API_KEY=sk-ant-...
# Cache model weights in the service user's home so they persist:
Environment=HOME=/opt/localmarkdown
Environment=XDG_CACHE_HOME=/opt/localmarkdown/.cache

ExecStart=/opt/localmarkdown/venv/bin/python \
          /opt/localmarkdown/app/markdown_mcp_server.py \
          watch /opt/localmarkdown/inbox \
          --output /opt/localmarkdown/out

Restart=on-failure
RestartSec=5

# Hardening (optional but recommended)
NoNewPrivileges=true
ProtectSystem=strict
ProtectHome=true
ReadWritePaths=/opt/localmarkdown
PrivateTmp=true

[Install]
WantedBy=multi-user.target
```

### 6.3 Enable and watch the logs

```bash
sudo systemctl daemon-reload
sudo systemctl enable --now localmarkdown.service

systemctl status localmarkdown.service        # health
journalctl -u localmarkdown.service -f         # live logs
```

Drop files into `/opt/localmarkdown/inbox` (via Samba/SFTP/syncthing) and watch
`*.md` appear in `/opt/localmarkdown/out`.

### 6.4 Serving MCP from the Linux server to a remote Claude

`serve` mode talks stdio, so a remote client connects by launching the command
**over SSH**. In the client config:

```json
{
  "mcpServers": {
    "localmarkdown": {
      "command": "ssh",
      "args": ["user@your-server",
               "/opt/localmarkdown/venv/bin/python",
               "/opt/localmarkdown/app/markdown_mcp_server.py", "serve",
               "--watch", "/opt/localmarkdown/inbox",
               "--output", "/opt/localmarkdown/out"]
    }
  }
}
```

Use SSH key auth so no password prompt blocks the stdio handshake. (Run the
systemd `watch` service for continuous conversion, and this on-demand `serve`
command for the MCP query interface — they share the same `--output` dir.)

### 6.5 No-systemd alternatives

```bash
# tmux: survives logout, easy to reattach
tmux new -s localmd
python markdown_mcp_server.py watch ./inbox --output ./out
# detach: Ctrl-b then d ;  reattach: tmux attach -t localmd

# or nohup
nohup python markdown_mcp_server.py watch ./inbox --output ./out > localmd.log 2>&1 &
```

---

## 7. Troubleshooting

| Symptom | Fix |
|---------|-----|
| `ffmpeg not found` during transcription | Install ffmpeg and ensure it’s on PATH. |
| `Docling is not installed` | `pip install docling` (covers PDF/Word/Excel/PPTX/HTML/image OCR; pulls torch). |
| `faster-whisper is not installed` | `pip install faster-whisper`. |
| `transformers is not installed` (image describe) | `pip install transformers` for local BLIP, or set `LM_IMAGE_DESCRIBE=ollama`/`none`. |
| `description unavailable: <urlopen error ...>` | Ollama/local server not running. `ollama serve` + `ollama pull llama3.2-vision`, or check `LM_OLLAMA_HOST`/`LM_OPENAI_BASE_URL`. |
| Ollama description times out on big images | Raise `LM_VLM_TIMEOUT`, or use a smaller vision model. |
| `anthropic` describer errors / 401 | `pip install anthropic` and set `ANTHROPIC_API_KEY`, or use `LM_IMAGE_DESCRIBE=ollama`. |
| `libGL.so.1: cannot open shared object` (Linux/WSL) | `sudo apt install libgl1 libglib2.0-0` (Docling's OCR uses OpenCV). |
| Server appears to “hang” on start | First model download is in progress — watch stderr/journalctl. |
| Watcher misses files on `/mnt/c` (WSL) | Use a Linux-native inbox, or trigger with `process`. |
| MCP client shows garbled handshake | Don’t print to stdout; logs already go to stderr. |
| Out of memory on large media | Use a smaller `LM_WHISPER_MODEL` (e.g. `tiny`/`base`) or set `LM_MAX_BYTES`. |
| Re-process an unchanged file | `process <path> --force`, or the `process_path(..., force=true)` tool. |
| `.msg` files skipped / `extract-msg is not installed` | `pip install extract-msg` (already in requirements.txt). |
| `.pst` shows `readpst not found` | `sudo apt install pst-utils`. The `.pst` is exploded into one Markdown per message. |
| Which email did an attachment come from? | Open the attachment `.md`: its frontmatter has `attachment_of`, `email_subject`, and `container`. Files also share the email's name prefix and carry an `__attNN__` marker. |

---

## 8. Project files

```
LocalMarkdown/
├── markdown_mcp_server.py   # the standalone server (serve / watch / process)
├── requirements.txt         # Python dependencies
├── drag_and_drop.bat        # Windows drag-and-drop helper
├── README.md                # this guide
├── inbox/                   # (created) default watched folder
└── markdown_output/         # (created) generated *.md + index.json
```

## 9. Security notes

- By default the watcher only **reads** source files and **writes** Markdown to
  the output directory — it never modifies or deletes your originals. The one
  exception is opt-in `LM_CONSUME_INBOX=1`, which **moves** each original into its
  output folder after a successful conversion (to keep the inbox clear); it still
  never deletes data — the file is relocated, not removed.
- OCR/transcription run fully **locally**; no file content leaves the machine
  except when your MCP client requests it.
- Run the Linux service as an unprivileged user (section 6) and keep the inbox
  separate from system paths.
