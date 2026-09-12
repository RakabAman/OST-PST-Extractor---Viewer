# OST / PST Extractor & Viewer

A cross-platform desktop application for **browsing, searching, and extracting** Outlook `.ost` / `.pst` mailboxes — with or without the `libpff` command-line tools. Built with **Python 3.11** and **PyQt5**.

Three workflows, one UI:

| Mode | What it does | Needs `pffexport.exe`? |
|------|--------------|:----------------------:|
| **Direct** | Open the OST/PST in place via `pypff`. Metadata-only index; bodies and attachments read on demand. | No |
| **Extract (pffexport)** | Drive `pffexport` / `pffinfo` as a subprocess; produces a full on-disk tree. | Yes |
| **Extract (pypff)** | Extract in-process via `pypff`. No external binaries. | No |
| **Browse** | Open an existing extraction folder (from either extractor). | No |

---

## Table of Contents

- [Features](#features)
- [Screenshots](#screenshots)
- [Requirements](#requirements)
- [Installation](#installation)
- [Usage](#usage)
  - [Direct mode](#direct-mode-recommended)
  - [Extract mode](#extract-mode)
  - [Browse an extracted folder](#browse-an-extracted-folder)
  - [Working with attachments](#working-with-attachments)
- [How It Works](#how-it-works)
  - [Architecture](#architecture)
  - [The lock problem](#the-lock-problem)
  - [Cache layout](#cache-layout)
  - [Database schema](#database-schema)
- [pffexport / pffinfo Options Reference](#pffexport--pffinfo-options-reference)
- [Troubleshooting](#troubleshooting)
- [License](#license)
- [Acknowledgements](#acknowledgements)

---

## Features

- 📂 **Three extraction backends** — direct (`pypff`), CLI (`pffexport`), in-process (`pypff extract`). Auto-picks the best one.
- 🌳 **Whole-tree browsing** — every root (`Root - Mailbox`, `Root - Public`, `NON_IPM_SUBTREE`, …) is shown, matching Outlook itself. Not limited to `IPM_SUBTREE`.
- ⚡ **Metadata-only direct scan** — first index of a 6 GB OST takes seconds; the cache is a few MB.
- 🖱️ **Lazy body / attachment loading** — bodies read from the OST in memory on message click; attachments extracted to `%TEMP%` only when opened.
- 🔒 **Friendly lock handling** — if Outlook holds a byte-range lock, the app either transparently copies the file to `%TEMP%` once or shows a clear "close Outlook" dialog. No tracebacks.
- 🔍 **Search** across subject, sender name, and sender email.
- 📎 **Attachment manager** — left panel lists every attachment per message; `View → Show All Attachments` opens a cross-mailbox dock.
- 💾 **Portable SQLite cache** — relative paths; move or rename the folder freely.
- 🔠 **Zoom controls** — `View → Zoom In / Out / Reset`.
- 🖥️ **Cross-platform** — Windows, Linux, macOS (PyQt5 + QtWebEngine).

---

## Screenshots

> Add screenshots here.

```
docs/screenshot-direct-mode.png
docs/screenshot-extract-dialog.png
docs/screenshot-attachments.png
docs/screenshot-full-tree.png
```

---

## Requirements

- **Python 3.11** (required — the `pypff` wheel is built for the CPython 3.11 ABI)
- **PyQt5** with **PyQtWebEngine**
- **`libpff-python`** (the `pypff` module) — pre-built wheel from [`Sygmei/wheels`](https://github.com/Sygmei/wheels)
- **Optional:** `pffexport` and `pffinfo` binaries for the CLI extract backend — [libyal/libpff releases](https://github.com/libyal/libpff/releases)

---

## Installation

### 1. Clone the repository

```bash
git clone https://github.com/RakabAman/OST-PST-Extractor---Viewer.git
cd OST-PST-Extractor---Viewer
```

### 2. Install Python dependencies (on Python 3.11)

```bash
py -3.11 -m pip install PyQt5 PyQtWebEngine
py -3.11 -m pip install https://github.com/Sygmei/wheels/raw/main/libpff_python-20211114-cp311-cp311-win_amd64.whl
```

Linux / macOS users: install `libpff` from source, then `pip install libpff-python` in the 3.11 environment.

### 3. (Optional) Place the `pffexport` / `pffinfo` binaries

Only needed for the CLI extract backend. The app searches:

1. Next to `ost_viewer_v6.py`
2. Inside a `tools/` subfolder
3. Anywhere on `PATH`

Example layout:

```
OST-PST-Extractor---Viewer/
├── ost_viewer_v6.py
├── pffexport.exe        ← optional
├── pffinfo.exe          ← optional
└── tools/               ← or here
    ├── pffexport
    └── pffinfo
```

### 4. Run

```bash
py -3.11 ost_viewer_v6.py
```

---

## Usage

### Direct mode (recommended)

Reads the OST/PST **without extracting anything to disk**.

1. Click **Open OST/PST (direct)…** in the toolbar (or `File → Open OST/PST (direct)…`).
2. Pick the `.ost` or `.pst` file.
3. First run builds a metadata-only index — a few MB — inside `<ost_dir>/<stem>.viewer/mail_cache.db`.
   - If Outlook has the file open, the app transparently copies it to `%TEMP%` once and reads the copy.
   - If the copy also fails, you get a clear dialog asking you to close Outlook.
4. Click a message → the body is read from the OST in memory.
5. Click an attachment in the left panel → it is extracted to `%TEMP%\OSTMailViewer\<hash>\` and opened with the OS default app.

Subsequent opens reuse the cache instantly (invalidated automatically if the source's mtime/size change).

### Extract mode

Full extraction to disk using either backend.

1. Click **Extract OST/PST…** in the toolbar or `File → Extract OST/PST…`.
2. In the dialog:
   - **Backend** — `Auto`, `pffexport CLI`, or `pypff (in-process)`.
   - **pffexport / pffinfo paths** — auto-detected; browse to override.
   - **Input file** — the `.ost` / `.pst` to extract.
   - **Output folder** — where the tree will be written.
   - **Target basename (`-t`)** — top-level folder name; auto-filled from the source filename.
   - **pffexport options** — codepage, format, mode, log file, `-d`, `-q`, `-v`.
   - **pffinfo options** — optional metadata preview before extraction.
3. Click **Extract**. A progress dialog appears for the pypff backend; the CLI backend runs `pffexport` in a subprocess.
4. When done, the app offers to open the extracted folder — this scans it and builds its `mail_cache.db`.

The extractor writes the **whole tree** (every root), not only `IPM_SUBTREE`:

```
<output>/<target>/
├── Root - Mailbox/
│   ├── IPM_SUBTREE/
│   │   ├── Archive/
│   │   ├── Calendar/
│   │   └── …
│   └── Common Views/
└── Root - Public/
    └── …
```

### Browse an extracted folder

`File → Open Extracted Folder…` — select the folder that contains the extraction output (the same one you chose as *Output folder*). If `mail_cache.db` exists it loads instantly; otherwise the folder is scanned and the cache is built.

### Working with attachments

- **Left panel** (below the folder tree) — for the selected message, `Outlook Data` shows metadata and `Attachments` lists every attachment.
  - Double-click → open with the OS default app.
  - Right-click → *Open* / *Open file location* / *Copy file path*.
- **`View → Show All Attachments`** — opens a dock listing every attachment across the mailbox with columns *Filename, Message Subject, Sender, Date, Folder, Size* and a search box.
  - Double-click to open; right-click for the same three actions.

---

## How It Works

### Architecture

```
┌──────────────────────────────────────────────────────────────────┐
│                         MainWindow (PyQt5)                       │
│  ┌────────────┐  ┌──────────────────┐  ┌──────────────────────┐  │
│  │ FolderTree │  │  Message Table   │  │   QWebEngineView     │  │
│  │  (QTree)   │  │   (QTableView)   │  │   (HTML preview)     │  │
│  └────────────┘  └──────────────────┘  └──────────────────────┘  │
│         │                │                        │              │
│         └────────────────┴────────────────────────┘              │
│                          │                                       │
│                    MailDB (SQLite)                               │
│                          │                                       │
│         ┌────────────────┴──────────────────┐                    │
│         │                                   │                    │
│   mail_cache.db                  PffBodyLoader (direct)          │
│   (folders, messages,            ─ on-demand HTML body           │
│    attachments, meta)            ─ on-demand attachment bytes    │
└──────────────────────────────────────────────────────────────────┘
```

Two entry points feed the DB:

| Mode | Producer |
|------|----------|
| Direct | `DirectScannerThread` — metadata only |
| Extracted | `ScannerThread` — walks the on-disk extraction folder |

Two extractors:

| Backend | Class |
|---------|-------|
| CLI | `PFFExportDialog._run_pffexport` → `subprocess` |
| In-process | `PFFExportDialog._run_pypff_extract` → `PypffExtractThread` |

### The lock problem

`libpff` opens files with a share mode that Outlook does not honour when it holds the OST/PST. The app therefore:

1. Tries a direct `pypff.open()`.
2. On a `locked` / `cannot access` failure, copies the file to `%TEMP%\OSTMailViewer\locked\<hash>.ost` (reused while the source mtime is unchanged) and opens the copy.
3. If the copy also fails, raises `OstLockedError` with a clear message.

Same helper is used by the body loader, so clicking a message works even if Outlook has the file locked.

### Cache layout

**Direct mode:**

```
<ost_dir>/
├── mailbox.ost                        ← source, untouched
└── mailbox.viewer/                    ← a few MB
    └── mail_cache.db
```

**Extracted mode:**

```
<output>/<target>/
├── Root - Mailbox/
│   └── …
├── Root - Public/
│   └── …
└── mail_cache.db
```

**Temp (created on demand):**

```
%TEMP%/OSTMailViewer/
├── locked/<hash>.ost                  ← only if the source is locked
└── <hash>/<msg_id>/<attachment>       ← lazily extracted attachments
```

### Database schema

```sql
CREATE TABLE folders (
    id INTEGER PRIMARY KEY,
    name TEXT,
    parent_id INTEGER,
    full_path TEXT UNIQUE          -- relative to mail_cache.db
);

CREATE TABLE messages (
    id INTEGER PRIMARY KEY,
    folder_id INTEGER,
    subject TEXT,
    sender_name TEXT,
    sender_email TEXT,
    date_utc DATETIME,
    size INTEGER,
    has_attachments BOOLEAN,
    message_html_path TEXT,        -- relative (extracted mode)
    outlook_headers_path TEXT,     -- relative
    internet_headers_path TEXT,    -- relative
    conversation_topic TEXT,
    to_emails TEXT,
    cc_emails TEXT,
    pff_folder_path TEXT,          -- direct mode: dot path, e.g. "0.3.1"
    pff_msg_idx INTEGER,           -- direct mode: message index in that folder
    FOREIGN KEY(folder_id) REFERENCES folders(id)
);

CREATE TABLE attachments (
    id INTEGER PRIMARY KEY,
    message_id INTEGER,
    filename TEXT,
    filepath TEXT,                 -- relative, empty in direct mode
    size INTEGER,
    FOREIGN KEY(message_id) REFERENCES messages(id)
);

CREATE TABLE meta (
    key TEXT PRIMARY KEY,
    value TEXT
);
```

Indexes are created on `folder_id`, `subject`, `sender_email`, `date_utc`, and `message_id`.

`meta` stores `mode`, `source_path`, `source_mtime`, `source_size` for direct-mode cache validation.

**Relative paths** keep the cache fully portable — move or rename the containing folder without a rescan.

---

## pffexport / pffinfo Options Reference

### `pffexport`

| Flag | Description | Values |
|------|-------------|--------|
| `-c codepage` | ASCII codepage | `ascii`, `windows-874`, `932`, `936`, `949`, `950`, `1250`–`1258` (default `windows-1252`) |
| `-f format` | Output format | `all`, `html`, `rtf`, `text` (default `text`) |
| `-l logfile` | Log file path | any path |
| `-m mode` | Export mode | `all`, `debug`, `items` (default), `recovered` |
| `-t target` | Target directory basename | default: source filename |
| `-d` | Dump item values to separate files | flag |
| `-q` | Quiet mode | flag |
| `-v` | Verbose output | flag |

### `pffinfo`

| Flag | Description |
|------|-------------|
| `-a` | Show allocation information |
| `-c codepage` | ASCII codepage |
| `-v` | Verbose output |

All options are exposed in the **Extract OST/PST** dialog.

---

## Troubleshooting

| Symptom | Fix |
|--------|-----|
| *"The file is currently in use by another process…"* | Close Outlook or any other app that has the OST/PST open, then retry. The app may also have fallen back to a `%TEMP%` copy automatically — check the console. |
| `pypff is not installed` | You are running a Python that isn't 3.11, or the wheel isn't installed. Run with `py -3.11` and install the wheel listed above. |
| `pffexport / pffinfo not found` | Only matters for the CLI backend. Place the binaries next to the script, in `tools/`, or on `PATH`, or switch the backend to `pypff (in-process)`. |
| Message preview is blank | For extracted mode, re-extract with Format = `html` or `all`. For direct mode, this usually means the body was empty in the source. |
| Cache seems stale | `File → Rescan` deletes `mail_cache.db` and rebuilds. For direct mode it also reopens the source. |
| Attachment opens the wrong file (direct mode) | This was a v6.0 bug; v6.1 orders attachments by `id` (pypff order) instead of filename. Delete the `.viewer` folder and rescan. |
| Chinese / Japanese / Arabic text garbled | Change the codepage (`-c`) in the extract dialog and re-extract. |
| Direct-mode cache too big | It shouldn't be — metadata only. If it is, delete `<stem>.viewer/` and rescan. |

---

## License

MIT License. See [`LICENSE`](LICENSE).

`libpff` is a separate project by Joachim Metz, distributed under the GNU LGPL v3. It is **not** bundled with this application.

---

## Acknowledgements

- [Joachim Metz](https://github.com/libyal) and contributors for **libpff**.
- [Sygmei/wheels](https://github.com/Sygmei/wheels) for the pre-built `libpff-python` Windows wheel.
- The **PyQt5 / Qt** teams.
- Everyone who filed issues and provided sample exports.
