OST Mail Viewer
A cross-platform desktop application for browsing and searching OST/PST mailbox exports produced by libpff's pffexport tool. Built with Python 3 and PyQt5, it provides a fast, cache-backed interface to explore mail folders, read messages, and manage attachments — without needing Outlook or any proprietary software.

The app also includes a built-in extraction dialog that drives pffexport and pffinfo for you, so you can go from a raw .ost/.pst file to a browsable mailbox in a single workflow.

Table of Contents
Features

Screenshots

Requirements

Installation

Usage

1. Extracting an OST/PST with libpff

2. Browsing the Extracted Mailbox

3. Working with Attachments

How It Works

Architecture

Database Schema

Relative Path Caching

pffexport / pffinfo Options Reference

Troubleshooting

Release Page Summary

License

Acknowledgements

Features
📂 Native OST/PST extraction — run pffexport (and optionally pffinfo) directly from the app.

🔍 Full-text search across subject, sender name, and sender email.

📎 Attachment manager — browse every attachment across the mailbox in a dedicated dock, with open / open-location / copy-path actions.

🌳 Folder tree navigation mirroring the original Outlook IPM_SUBTREE hierarchy.

💾 SQLite cache — first scan builds mail_cache.db inside the extracted folder; subsequent launches load instantly.

🔗 Portable cache with relative paths — move or rename the extracted folder freely; the cache keeps working.

🔠 Zoom controls (View → Zoom In / Out / Reset) with adjustable font and web view scaling.

🖥️ Cross-platform — Windows, Linux, macOS (PyQt5 + QtWebEngine).

Screenshots
Add screenshots here (folder tree + message list + preview).

text
docs/screenshot-main.png
docs/screenshot-extract-dialog.png
docs/screenshot-attachments.png
Requirements
Python 3.8+

PyQt5 (including PyQtWebEngine)

libpff binaries (pffexport, pffinfo) — see libyal/libpff

Windows: download the pre-built binaries from the libpff releases page.

Linux: sudo apt install libpff-utils (or build from source).

macOS: brew install libpff (or build from source).

Installation
1. Clone the repository
bash
git clone https://github.com/<your-user>/ost-mail-viewer.git
cd ost-mail-viewer
2. Install Python dependencies
bash
pip install PyQt5 PyQtWebEngine
3. Get the libpff tools
Place pffexport (and pffinfo) in one of the following locations — the app searches all of them automatically:

Next to ost_viewer_v2.py

Inside a tools/ subfolder next to the script

Anywhere on your system PATH

Example layout:

text
ost-mail-viewer/
├── ost_viewer_v2.py
├── pffexport.exe        ← Windows binary
├── pffinfo.exe
└── tools/               ← or here
    ├── pffexport
    └── pffinfo
4. Run the app
bash
python ost_viewer_v2.py
Usage
1. Extracting an OST/PST with libpff
The app can drive pffexport for you so you don't have to remember the command-line flags.

Launch the app and click Extract OST/PST… in the toolbar (or File → Extract OST/PST…).

In the dialog:

pffexport / pffinfo paths are auto-detected. If detection fails, click Browse… to select the executables manually.

Input file — select the .ost or .pst file to extract.

Output folder — pick an empty folder; this is where pffexport will write Root - Mailbox/IPM_SUBTREE/… and where mail_cache.db will be created.

Target basename (-t) — auto-filled from the input filename; change it if you want a different top-level folder name.

pffexport options — codepage, format (text, html, rtf, all), mode (items, all, debug, recovered), log file, and the -d, -q, -v flags.

pffinfo options — optionally run pffinfo first to display mailbox metadata and allocation info before extracting.

Click Extract.

When extraction finishes, the app offers to open the extracted folder immediately — this triggers an automatic first-time scan and populates the cache.

Tip: The extraction runs with its working directory set to the chosen output folder, so the -t basename behaves exactly as if you ran pffexport from a terminal there.

2. Browsing the Extracted Mailbox
File → Open Folder — select the folder that contains Root - Mailbox/. The app looks for mail_cache.db:

If it exists, it loads instantly from the cache.

If not, it scans IPM_SUBTREE and builds the cache.

Folder tree (left) — mirrors the Outlook folder hierarchy.

Message table (top-right) — sortable columns: Subject, Sender, Date, Size, Attachments.

Search bar — filter by subject / sender name / sender email.

Only with attachments — checkbox to show only messages that carry attachments.

Preview pane (bottom-right) — renders the exported Message.html via QtWebEngine.

View → Zoom In / Out / Reset — scales fonts and the web preview.

3. Working with Attachments
Inline strip below the preview shows attachment icons for the currently selected message. Click to open with the default OS handler.

View → Show All Attachments opens a dock listing every attachment in the mailbox with columns: Filename, Message Subject, Sender, Date, Folder, Size.

Double-click to open.

Right-click for Open, Open file location, Copy file path.

How It Works
Architecture
text
┌──────────────────────────────────────────────────────────────┐
│                        MainWindow (PyQt5)                    │
│  ┌────────────┐ ┌──────────────────┐ ┌────────────────────┐  │
│  │ FolderTree │ │  Message Table   │ │   QWebEngineView   │  │
│  │  (QTree)   │ │    (QTableView)  │ │  (HTML preview)    │  │
│  └────────────┘ └──────────────────┘ └────────────────────┘  │
│           │              │                     │             │
│           └──────────────┴─────────────────────┘             │
│                          │                                   │
│                    MailDB (SQLite)                           │
│                          │                                   │
│              ┌───────────┴────────────┐                      │
│              │                        │                      │
│      mail_cache.db             Extracted files               │
│      (folders, messages,       (Message.html,                │
│       attachments)              OutlookHeaders.txt,          │
│                                 InternetHeaders.txt,         │
│                                 Attachments/)                │
└──────────────────────────────────────────────────────────────┘
                          │
                          │ driven by
                          ▼
              ┌───────────────────────┐
              │ PFFExportDialog       │
              │  ├── pffinfo (opt.)   │
              │  └── pffexport        │
              └───────────────────────┘
Scan pipeline (ScannerThread):

Walk the selected root folder to find Root - Mailbox/IPM_SUBTREE.

Identify message folders by the presence of OutlookHeaders.txt.

Rebuild the folder hierarchy as folders rows (parent_id links).

For each message folder, parse:

OutlookHeaders.txt → subject, sender name, submit time, size, attachment flag, conversation topic.

InternetHeaders.txt → From, To, Cc email addresses.

Message.html → stored as the preview source.

Attachments/ → inserted as rows in the attachments table.

Commit and hand the DB back to the UI.

Database Schema
sql
CREATE TABLE folders (
    id INTEGER PRIMARY KEY,
    name TEXT,
    parent_id INTEGER,
    full_path TEXT UNIQUE           -- stored relative to mail_cache.db
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
    message_html_path TEXT,         -- relative
    outlook_headers_path TEXT,      -- relative
    internet_headers_path TEXT,     -- relative
    conversation_topic TEXT,
    to_emails TEXT,
    cc_emails TEXT,
    FOREIGN KEY(folder_id) REFERENCES folders(id)
);

CREATE TABLE attachments (
    id INTEGER PRIMARY KEY,
    message_id INTEGER,
    filename TEXT,
    filepath TEXT,                  -- relative
    size INTEGER,
    FOREIGN KEY(message_id) REFERENCES messages(id)
);
Indexes are created on folder_id, subject, sender_email, date_utc, and message_id for fast filtering.

Relative Path Caching
All paths stored in mail_cache.db are relative to the folder that contains the database (i.e. the extracted root). Two helpers on MailDB handle translation:

python
def _to_relative(self, abs_path):   # used on INSERT
    return os.path.relpath(abs_path, self.base_dir)

def _to_absolute(self, rel_path):   # used on SELECT
    return os.path.normpath(os.path.join(self.base_dir, rel_path))
Why this matters: the cache is written inside the extracted folder. If you move or rename that folder, absolute paths would break. Relative paths keep the cache fully portable — no rescan needed.

pffexport / pffinfo Options Reference
pffexport
Flag	Description	Values
-c codepage	ASCII codepage	ascii, windows-874, windows-932, windows-936, windows-949, windows-950, windows-1250–1258 (default: windows-1252)
-f format	Output format	all, html, rtf, text (default: text)
-l logfile	Log file path	any path
-m mode	Export mode	all, debug, items (default), recovered
-t target	Target directory basename	default: source filename
-d	Dump item values to separate files	flag
-q	Quiet mode	flag
-v	Verbose output	flag
pffinfo
Flag	Description
-a	Show allocation information
-c codepage	ASCII codepage
-v	Verbose output
All of these are exposed as widgets in the Extract OST/PST dialog.

Troubleshooting
Symptom	Fix
"pffexport / pffinfo not found"	Place the binaries next to ost_viewer_v2.py, in a tools/ subfolder, or on PATH. Or select them manually in the dialog.
"IPM_SUBTREE not found!"	The output folder doesn't contain Root - Mailbox/IPM_SUBTREE. Re-run pffexport with the correct -t and output path.
Message preview is blank	pffexport may have been run without -f html. Re-extract with Format = html or all.
Cache seems stale after re-extracting	Use File → Rescan — it deletes mail_cache.db and rebuilds.
Attachments fail to open	The extracted folder was moved and files are missing, or the DB still holds absolute paths from an old version. Delete mail_cache.db and rescan.
Chinese / Japanese / Arabic text garbled	Change the codepage (-c) in the extract dialog and re-extract.
