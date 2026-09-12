#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
OST Mail Viewer – Read and browse OST/PST mailboxes
Copyright (c) 2025

Requires Python 3.11 with:
    pip install PyQt5 PyQtWebEngine
    pip install https://github.com/Sygmei/wheels/raw/main/libpff_python-20211114-cp311-cp311-win_amd64.whl

Three modes:
  * Extract   – run pffexport CLI or pypff, then scan the resulting folder
  * Browse    – open an already-extracted folder
  * Direct    – read an OST/PST in-place via pypff (metadata-only index,
                bodies and attachments read on demand)

v6.1 changes
------------
* Friendly lock message when the OST/PST is held by Outlook (temp-copy fallback)
* ScannerThread now walks the entire extracted root, not just IPM_SUBTREE
* PypffExtractThread now extracts the entire pypff tree, not just IPM_SUBTREE
* MailDB.insert_folder lookup fixes duplicated full_path rows
"""

import sys
import os
import re
import shutil
import hashlib
import sqlite3
import datetime
import subprocess

from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
from PyQt5.QtWebEngineWidgets import QWebEngineView

# ----------------------------------------------------------------------
# pypff (libpff-python) – only available on Python 3.11 with the wheel
try:
    import pypff
    HAS_PYPFF = True
except ImportError:
    pypff = None
    HAS_PYPFF = False

# ----------------------------------------------------------------------
# Global font scaling factor (1.0 = normal)
SCALE_FACTOR = 1.0

# ----------------------------------------------------------------------
# Direct-mode cache root.
#
# Preferred location: next to the OST/PST file, in <stem>.viewer/
# Fallback: %LOCALAPPDATA%\OSTMailViewer\direct\ (used only if the source
# folder is not writable, e.g. read-only network share).
APP_CACHE_ROOT = os.path.join(
    os.environ.get('LOCALAPPDATA') or os.path.expanduser('~/.cache'),
    'OSTMailViewer', 'direct')


class OstLockedError(Exception):
    """Raised when the OST/PST cannot be opened because another process
    (typically Outlook) holds a byte-range lock on it."""


# ----------------------------------------------------------------------
# Utility: find pffexport / pffinfo executables
def find_pff_tool(tool_name):
    exe_name = tool_name + '.exe' if sys.platform == 'win32' else tool_name
    app_dir = os.path.dirname(os.path.abspath(__file__))
    candidate = os.path.join(app_dir, exe_name)
    if os.path.isfile(candidate):
        return candidate
    candidate = os.path.join(app_dir, 'tools', exe_name)
    if os.path.isfile(candidate):
        return candidate
    from shutil import which
    found = which(exe_name)
    if found:
        return found
    return None


def direct_cache_dir(ost_path):
    """
    Return the cache directory for an OST/PST.

    Tries <ost_dir>/<stem>.viewer/ first (portable, same drive as source).
    Falls back to %LOCALAPPDATA%\\OSTMailViewer\\direct\\<stem>_<hash>\\ if
    the source folder isn't writable.
    """
    ost_path = os.path.abspath(ost_path)
    stem = re.sub(r'[<>:"/\\|?*\x00-\x1f]', '_',
                  os.path.splitext(os.path.basename(ost_path))[0]) or 'mailbox'

    candidate = os.path.join(os.path.dirname(ost_path), f'{stem}.viewer')
    try:
        os.makedirs(candidate, exist_ok=True)
        probe = os.path.join(candidate, '.writable')
        with open(probe, 'w') as fh:
            fh.write('1')
        os.remove(probe)
        return candidate
    except OSError:
        pass

    os.makedirs(APP_CACHE_ROOT, exist_ok=True)
    h = hashlib.md5(ost_path.encode('utf-8')).hexdigest()[:12]
    return os.path.join(APP_CACHE_ROOT, f'{stem}_{h}')


def open_ost_locked_safe(ost_path, log=print):
    """
    Open an OST/PST with pypff.

    If another process (typically Outlook) holds a byte-range lock, copy
    the file once to %TEMP%\\OSTMailViewer\\locked\\<hash>.ost and open the
    copy. The copy is reused while the source mtime hasn't advanced.

    Returns (pypff_file, effective_path).
    Raises OstLockedError if even the copy fails.
    """
    ost_path = os.path.normpath(os.path.abspath(ost_path))

    # Attempt 1 – direct open
    try:
        f = pypff.file()
        f.open(ost_path)
        return f, ost_path
    except OSError as e:
        msg = str(e).lower()
        if 'locked' not in msg and 'cannot access' not in msg:
            raise   # some other error – surface it as-is
        log('  ! OST/PST is locked by another process; copying to TEMP.')

    # Attempt 2 – copy to TEMP and open the copy
    temp_root = os.path.join(
        os.environ.get('TEMP') or os.path.expanduser('~/.cache'),
        'OSTMailViewer', 'locked')
    os.makedirs(temp_root, exist_ok=True)
    h = hashlib.md5(ost_path.encode('utf-8')).hexdigest()[:12]
    ext = os.path.splitext(ost_path)[1].lower() or '.ost'
    temp_copy = os.path.join(temp_root, f'{h}{ext}')

    need_copy = True
    if os.path.exists(temp_copy):
        try:
            if os.path.getmtime(temp_copy) >= os.path.getmtime(ost_path):
                need_copy = False
        except OSError:
            pass

    if need_copy:
        log(f'  Copying {ost_path}')
        log(f'      -> {temp_copy}')
        try:
            shutil.copy2(ost_path, temp_copy)
        except OSError as e:
            raise OstLockedError(
                "The file is currently in use by another process "
                "(most likely Microsoft Outlook).\n\n"
                "Please close Outlook (or any other application that has "
                "this OST/PST open) and try again.\n\n"
                f"File: {ost_path}\n\nDetails: {e}")

    try:
        f = pypff.file()
        f.open(temp_copy)
    except OSError as e:
        raise OstLockedError(
            "Could not open the OST/PST even after copying it to TEMP.\n\n"
            f"File: {ost_path}\n\nDetails: {e}")
    log(f'  Opened a locked-OST copy: {temp_copy}')
    return f, temp_copy


# ----------------------------------------------------------------------
# Database helper – stores relative paths
class MailDB:
    def __init__(self, db_path):
        self.db_path = os.path.abspath(db_path)
        self.base_dir = os.path.dirname(self.db_path)
        self.conn = None
        self.init_db()

    def _to_relative(self, abs_path):
        if not abs_path:
            return abs_path
        try:
            return os.path.relpath(abs_path, self.base_dir)
        except ValueError:
            return abs_path

    def _to_absolute(self, rel_path):
        if not rel_path:
            return rel_path
        if os.path.isabs(rel_path):
            return rel_path
        return os.path.normpath(os.path.join(self.base_dir, rel_path))

    def init_db(self):
        self.conn = sqlite3.connect(self.db_path)
        self.conn.row_factory = sqlite3.Row
        self.conn.execute('PRAGMA foreign_keys = ON;')
        self.conn.execute('''
            CREATE TABLE IF NOT EXISTS folders (
                id INTEGER PRIMARY KEY,
                name TEXT,
                parent_id INTEGER,
                full_path TEXT UNIQUE
            )
        ''')
        self.conn.execute('''
            CREATE TABLE IF NOT EXISTS messages (
                id INTEGER PRIMARY KEY,
                folder_id INTEGER,
                subject TEXT,
                sender_name TEXT,
                sender_email TEXT,
                date_utc DATETIME,
                size INTEGER,
                has_attachments BOOLEAN,
                message_html_path TEXT,
                outlook_headers_path TEXT,
                internet_headers_path TEXT,
                conversation_topic TEXT,
                to_emails TEXT,
                cc_emails TEXT,
                FOREIGN KEY(folder_id) REFERENCES folders(id)
            )
        ''')
        self.conn.execute('''
            CREATE TABLE IF NOT EXISTS attachments (
                id INTEGER PRIMARY KEY,
                message_id INTEGER,
                filename TEXT,
                filepath TEXT,
                size INTEGER,
                FOREIGN KEY(message_id) REFERENCES messages(id)
            )
        ''')
        self.conn.execute('''
            CREATE TABLE IF NOT EXISTS meta (
                key TEXT PRIMARY KEY,
                value TEXT
            )
        ''')
        cur = self.conn.execute('PRAGMA table_info(messages)')
        cols = [r[1] for r in cur.fetchall()]
        if 'pff_folder_path' not in cols:
            self.conn.execute(
                'ALTER TABLE messages ADD COLUMN pff_folder_path TEXT')
        if 'pff_msg_idx' not in cols:
            self.conn.execute(
                'ALTER TABLE messages ADD COLUMN pff_msg_idx INTEGER')

        self.conn.execute('CREATE INDEX IF NOT EXISTS idx_messages_folder ON messages(folder_id)')
        self.conn.execute('CREATE INDEX IF NOT EXISTS idx_messages_subject ON messages(subject)')
        self.conn.execute('CREATE INDEX IF NOT EXISTS idx_messages_sender ON messages(sender_email)')
        self.conn.execute('CREATE INDEX IF NOT EXISTS idx_messages_date ON messages(date_utc)')
        self.conn.execute('CREATE INDEX IF NOT EXISTS idx_attachments_message ON attachments(message_id)')
        self.conn.commit()

    def set_meta(self, key, value):
        self.conn.execute('INSERT OR REPLACE INTO meta (key, value) VALUES (?, ?)',
                          (key, str(value) if value is not None else ''))

    def get_meta(self, key, default=None):
        cur = self.conn.execute('SELECT value FROM meta WHERE key = ?', (key,))
        row = cur.fetchone()
        return row['value'] if row else default

    def close(self):
        if self.conn:
            self.conn.close()

    def clear_all(self):
        self.conn.execute('DELETE FROM attachments')
        self.conn.execute('DELETE FROM messages')
        self.conn.execute('DELETE FROM folders')
        self.conn.commit()

    def insert_folder(self, name, parent_id, full_path):
        rel_path = self._to_relative(full_path)
        cur = self.conn.cursor()
        cur.execute('INSERT OR IGNORE INTO folders (name, parent_id, full_path) VALUES (?, ?, ?)',
                    (name, parent_id, rel_path))
        if cur.lastrowid:
            return cur.lastrowid
        # INSERT OR IGNORE skipped — look up the existing row.
        cur.execute('SELECT id FROM folders WHERE full_path = ?', (rel_path,))
        row = cur.fetchone()
        return row['id'] if row else None

    def insert_message(self, folder_id, subject, sender_name, sender_email, date_utc,
                       size, has_attachments, html_path, outlook_path, internet_path,
                       conversation_topic, to_emails, cc_emails):
        rel_html = self._to_relative(html_path)
        rel_outlook = self._to_relative(outlook_path)
        rel_internet = self._to_relative(internet_path)
        cur = self.conn.cursor()
        cur.execute('''
            INSERT INTO messages
            (folder_id, subject, sender_name, sender_email, date_utc, size,
             has_attachments, message_html_path, outlook_headers_path,
             internet_headers_path, conversation_topic, to_emails, cc_emails)
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)
        ''', (folder_id, subject, sender_name, sender_email, date_utc,
              size, has_attachments, rel_html, rel_outlook, rel_internet,
              conversation_topic, to_emails, cc_emails))
        return cur.lastrowid

    def insert_attachment(self, message_id, filename, filepath, size):
        rel_path = self._to_relative(filepath)
        self.conn.execute('''
            INSERT INTO attachments (message_id, filename, filepath, size)
            VALUES (?,?,?,?)
        ''', (message_id, filename, rel_path, size))

    def get_folders(self, parent_id=None):
        cur = self.conn.cursor()
        if parent_id is None:
            cur.execute('SELECT * FROM folders WHERE parent_id IS NULL ORDER BY name')
        else:
            cur.execute('SELECT * FROM folders WHERE parent_id = ? ORDER BY name', (parent_id,))
        return cur.fetchall()

    def get_messages_for_folder(self, folder_id, search_text='', only_attachments=False):
        query = '''
            SELECT * FROM messages
            WHERE folder_id = ?
        '''
        params = [folder_id]
        if search_text:
            query += ' AND (subject LIKE ? OR sender_email LIKE ? OR sender_name LIKE ?)'
            like = f'%{search_text}%'
            params.extend([like, like, like])
        if only_attachments:
            query += ' AND has_attachments = 1'
        query += ' ORDER BY date_utc DESC'
        cur = self.conn.cursor()
        cur.execute(query, params)
        return cur.fetchall()

    def get_attachment_by_id(self, att_id):
        cur = self.conn.cursor()
        cur.execute('SELECT * FROM attachments WHERE id = ?', (att_id,))
        row = cur.fetchone()
        if row:
            d = dict(row)
            d['filepath'] = self._to_absolute(d['filepath'])
            return d
        return None

    def get_attachments_for_message(self, message_id):
        cur = self.conn.cursor()
        cur.execute('SELECT * FROM attachments WHERE message_id = ? ORDER BY id',
                    (message_id,))
        rows = cur.fetchall()
        result = []
        for row in rows:
            d = dict(row)
            d['filepath'] = self._to_absolute(d['filepath'])
            result.append(d)
        return result

    def get_all_attachments(self, search_text=''):
        query = '''
            SELECT a.id, a.filename, a.filepath, a.size,
                   m.subject AS msg_subject, m.sender_name, m.date_utc, f.name AS folder_name
            FROM attachments a
            JOIN messages m ON a.message_id = m.id
            JOIN folders f ON m.folder_id = f.id
        '''
        params = []
        if search_text:
            query += ' WHERE a.filename LIKE ? OR m.subject LIKE ?'
            like = f'%{search_text}%'
            params = [like, like]
        query += ' ORDER BY a.filename'
        cur = self.conn.cursor()
        cur.execute(query, params)
        rows = cur.fetchall()
        result = []
        for row in rows:
            d = dict(row)
            d['filepath'] = self._to_absolute(d['filepath'])
            result.append(d)
        return result

    def get_folder_by_id(self, folder_id):
        cur = self.conn.cursor()
        cur.execute('SELECT * FROM folders WHERE id = ?', (folder_id,))
        row = cur.fetchone()
        if row:
            d = dict(row)
            d['full_path'] = self._to_absolute(d['full_path'])
            return d
        return None

    def get_message_by_id(self, msg_id):
        cur = self.conn.cursor()
        cur.execute('SELECT * FROM messages WHERE id = ?', (msg_id,))
        row = cur.fetchone()
        if row:
            d = dict(row)
            for key in ('message_html_path', 'outlook_headers_path', 'internet_headers_path'):
                if d.get(key):
                    d[key] = self._to_absolute(d[key])
            return d
        return None

# ----------------------------------------------------------------------
# PFF Export Dialog
class PFFExportDialog(QDialog):
    """Dialog for extracting OST/PST files using pffexport / pffinfo."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle('Extract OST/PST using libpff')
        self.setMinimumWidth(650)
        self.result_root = None
        self._build_ui()
        self._auto_detect_tools()
        self._on_backend_changed()

    def _build_ui(self):
        layout = QVBoxLayout(self)

        backend_group = QGroupBox('Extraction Backend')
        backend_layout = QVBoxLayout(backend_group)

        self.backend_combo = QComboBox()
        self.backend_combo.addItem(
            'Auto — pffexport CLI if available, otherwise pypff', 0)
        self.backend_combo.addItem('pffexport CLI (external executable)', 1)
        self.backend_combo.addItem('pypff (in-process, no external tools)', 2)
        if not HAS_PYPFF:
            item = self.backend_combo.model().item(0)
            item.setEnabled(False)
            item = self.backend_combo.model().item(2)
            item.setEnabled(False)
        self.backend_combo.currentIndexChanged.connect(self._on_backend_changed)
        backend_layout.addWidget(self.backend_combo)

        self.backend_hint = QLabel('')
        self.backend_hint.setStyleSheet('color: #666; font-style: italic;')
        backend_layout.addWidget(self.backend_hint)
        layout.addWidget(backend_group)

        tool_group = QGroupBox('libpff Tools')
        tool_layout = QGridLayout(tool_group)

        tool_layout.addWidget(QLabel('pffexport:'), 0, 0)
        self.export_path_edit = QLineEdit()
        self.export_path_edit.setPlaceholderText('Path to pffexport executable...')
        tool_layout.addWidget(self.export_path_edit, 0, 1)
        export_browse = QPushButton('Browse...')
        export_browse.clicked.connect(self._browse_export)
        tool_layout.addWidget(export_browse, 0, 2)

        tool_layout.addWidget(QLabel('pffinfo:'), 1, 0)
        self.info_path_edit = QLineEdit()
        self.info_path_edit.setPlaceholderText('Path to pffinfo executable...')
        tool_layout.addWidget(self.info_path_edit, 1, 1)
        info_browse = QPushButton('Browse...')
        info_browse.clicked.connect(self._browse_info)
        tool_layout.addWidget(info_browse, 1, 2)

        layout.addWidget(tool_group)

        io_group = QGroupBox('Input / Output')
        io_layout = QGridLayout(io_group)

        io_layout.addWidget(QLabel('Input file (OST/PST):'), 0, 0)
        self.input_edit = QLineEdit()
        self.input_edit.setPlaceholderText('Select OST or PST file...')
        io_layout.addWidget(self.input_edit, 0, 1)
        input_browse = QPushButton('Browse...')
        input_browse.clicked.connect(self._browse_input)
        io_layout.addWidget(input_browse, 0, 2)

        io_layout.addWidget(QLabel('Output folder:'), 1, 0)
        self.output_edit = QLineEdit()
        self.output_edit.setPlaceholderText('Select output folder...')
        io_layout.addWidget(self.output_edit, 1, 1)
        output_browse = QPushButton('Browse...')
        output_browse.clicked.connect(self._browse_output)
        io_layout.addWidget(output_browse, 1, 2)

        io_layout.addWidget(QLabel('Target basename (-t):'), 2, 0)
        self.target_edit = QLineEdit()
        self.target_edit.setPlaceholderText('Optional – defaults to source filename')
        io_layout.addWidget(self.target_edit, 2, 1, 1, 2)

        layout.addWidget(io_group)

        exp_group = QGroupBox('pffexport Options')
        exp_layout = QGridLayout(exp_group)

        exp_layout.addWidget(QLabel('Codepage (-c):'), 0, 0)
        self.exp_codepage = QComboBox()
        self.exp_codepage.addItems([
            'windows-1252', 'ascii', 'windows-874', 'windows-932', 'windows-936',
            'windows-949', 'windows-950', 'windows-1250', 'windows-1251',
            'windows-1253', 'windows-1254', 'windows-1255', 'windows-1256',
            'windows-1257', 'windows-1258'
        ])
        exp_layout.addWidget(self.exp_codepage, 0, 1)

        exp_layout.addWidget(QLabel('Format (-f):'), 1, 0)
        self.exp_format = QComboBox()
        self.exp_format.addItems(['text', 'all', 'html', 'rtf'])
        exp_layout.addWidget(self.exp_format, 1, 1)

        exp_layout.addWidget(QLabel('Mode (-m):'), 2, 0)
        self.exp_mode = QComboBox()
        self.exp_mode.addItems(['items', 'all', 'debug', 'recovered'])
        exp_layout.addWidget(self.exp_mode, 2, 1)

        exp_layout.addWidget(QLabel('Log file (-l):'), 3, 0)
        self.exp_logfile = QLineEdit()
        self.exp_logfile.setPlaceholderText('Optional log file path...')
        exp_layout.addWidget(self.exp_logfile, 3, 1)

        self.chk_dump_values = QCheckBox('Dump item values (-d)')
        exp_layout.addWidget(self.chk_dump_values, 4, 0, 1, 2)

        self.chk_quiet = QCheckBox('Quiet mode (-q)')
        exp_layout.addWidget(self.chk_quiet, 5, 0, 1, 2)

        self.chk_verbose = QCheckBox('Verbose output (-v)')
        exp_layout.addWidget(self.chk_verbose, 6, 0, 1, 2)

        layout.addWidget(exp_group)

        inf_group = QGroupBox('pffinfo Options (run before export)')
        inf_layout = QVBoxLayout(inf_group)

        self.chk_run_info = QCheckBox('Run pffinfo first')
        self.chk_run_info.setChecked(True)
        inf_layout.addWidget(self.chk_run_info)

        self.chk_info_alloc = QCheckBox('Show allocation information (-a)')
        inf_layout.addWidget(self.chk_info_alloc)

        self.chk_info_verbose = QCheckBox('Verbose output (-v)')
        inf_layout.addWidget(self.chk_info_verbose)

        layout.addWidget(inf_group)

        btn_layout = QHBoxLayout()
        self.btn_extract = QPushButton('Extract')
        self.btn_extract.setDefault(True)
        self.btn_extract.clicked.connect(self._start_extract)
        btn_layout.addStretch()
        btn_layout.addWidget(self.btn_extract)

        cancel_btn = QPushButton('Cancel')
        cancel_btn.clicked.connect(self.reject)
        btn_layout.addWidget(cancel_btn)

        layout.addLayout(btn_layout)

    def _auto_detect_tools(self):
        exp = find_pff_tool('pffexport')
        if exp:
            self.export_path_edit.setText(exp)
        inf = find_pff_tool('pffinfo')
        if inf:
            self.info_path_edit.setText(inf)

    def _current_backend(self):
        choice = self.backend_combo.currentData()
        if choice == 1:
            return 'pffexport'
        if choice == 2:
            return 'pypff'
        path = self.export_path_edit.text().strip()
        if path and os.path.isfile(path):
            return 'pffexport'
        if HAS_PYPFF:
            return 'pypff'
        return 'pffexport'

    def _on_backend_changed(self):
        backend = self._current_backend()
        if backend == 'pffexport':
            self.backend_hint.setText(
                'pffexport.exe will be launched as a subprocess.')
        else:
            self.backend_hint.setText(
                'pypff will extract in-process (no external binaries).')

    def _browse_export(self):
        path, _ = QFileDialog.getOpenFileName(
            self, 'Select pffexport executable', '',
            'Executables (*.exe);;All files (*)')
        if path:
            self.export_path_edit.setText(path)

    def _browse_info(self):
        path, _ = QFileDialog.getOpenFileName(
            self, 'Select pffinfo executable', '',
            'Executables (*.exe);;All files (*)')
        if path:
            self.info_path_edit.setText(path)

    def _browse_input(self):
        path, _ = QFileDialog.getOpenFileName(
            self, 'Select OST/PST file', '',
            'Outlook files (*.ost *.pst);;All files (*)')
        if path:
            self.input_edit.setText(path)
            if not self.target_edit.text():
                base = os.path.splitext(os.path.basename(path))[0]
                self.target_edit.setText(base)

    def _browse_output(self):
        folder = QFileDialog.getExistingDirectory(self, 'Select output folder')
        if folder:
            self.output_edit.setText(folder)

    def _build_command(self, tool_path, is_export=True):
        args = [tool_path]
        if is_export:
            args.extend(['-c', self.exp_codepage.currentText()])
            args.extend(['-f', self.exp_format.currentText()])
            args.extend(['-m', self.exp_mode.currentText()])
            logfile = self.exp_logfile.text().strip()
            if logfile:
                args.extend(['-l', logfile])
            target = self.target_edit.text().strip()
            if target:
                args.extend(['-t', target])
            if self.chk_dump_values.isChecked():
                args.append('-d')
            if self.chk_quiet.isChecked():
                args.append('-q')
            if self.chk_verbose.isChecked():
                args.append('-v')
            args.append(self.input_edit.text().strip())
        else:
            if self.chk_info_alloc.isChecked():
                args.append('-a')
            if self.chk_info_verbose.isChecked():
                args.append('-v')
            args.append(self.input_edit.text().strip())
        return args

    def _start_extract(self):
        input_file = self.input_edit.text().strip()
        output_dir = self.output_edit.text().strip()

        if not input_file or not os.path.isfile(input_file):
            QMessageBox.warning(self, 'Error',
                                'Please select a valid input OST/PST file.')
            return
        if not output_dir or not os.path.isdir(output_dir):
            QMessageBox.warning(self, 'Error',
                                'Please select a valid output folder.')
            return

        backend = self._current_backend()

        if backend == 'pypff':
            if not HAS_PYPFF:
                QMessageBox.critical(
                    self, 'Error',
                    'pypff is not installed in this Python.\n\n'
                    'Install with:\n'
                    '    pip install https://github.com/Sygmei/wheels/raw/main/'
                    'libpff_python-20211114-cp311-cp311-win_amd64.whl')
                return
            self._run_pypff_extract(input_file, output_dir)
        else:
            export_path = self.export_path_edit.text().strip()
            if not export_path or not os.path.isfile(export_path):
                QMessageBox.warning(
                    self, 'Error',
                    'No valid pffexport executable selected.\n\n'
                    'Either pick one via Browse... or switch the backend '
                    'to "pypff (in-process)".')
                return
            self._run_pffexport(input_file, output_dir, export_path)

    def _run_pffexport(self, input_file, output_dir, export_path):
        if self.chk_run_info.isChecked():
            info_path = self.info_path_edit.text().strip()
            if info_path and os.path.isfile(info_path):
                info_args = self._build_command(info_path, is_export=False)
                try:
                    result = subprocess.run(
                        info_args, capture_output=True, text=True, timeout=120)
                    info_text = result.stdout + result.stderr
                    QMessageBox.information(
                        self, 'pffinfo Output',
                        f'Command: {" ".join(info_args)}\n\n'
                        f'{info_text[:3000]}')
                except Exception as e:
                    QMessageBox.warning(self, 'pffinfo Error', str(e))

        export_args = self._build_command(export_path, is_export=True)
        try:
            result = subprocess.run(
                export_args, capture_output=True, text=True,
                cwd=output_dir, timeout=600)
            if result.returncode == 0:
                QMessageBox.information(
                    self, 'Extraction Complete',
                    f'pffexport finished successfully.\n\n'
                    f'{result.stdout[:2000]}')
                self.result_root = self._effective_result_root(output_dir)
                self.accept()
            else:
                QMessageBox.critical(
                    self, 'Extraction Failed',
                    f'pffexport returned code {result.returncode}.\n\n'
                    f'STDOUT:\n{result.stdout[:2000]}\n\n'
                    f'STDERR:\n{result.stderr[:2000]}')
        except subprocess.TimeoutExpired:
            QMessageBox.critical(self, 'Timeout',
                                 'pffexport timed out after 600 seconds.')
        except Exception as e:
            QMessageBox.critical(self, 'Error', str(e))

    def _effective_result_root(self, output_dir):
        target = self.target_edit.text().strip()
        if not target:
            target = os.path.splitext(
                os.path.basename(self.input_edit.text().strip()))[0]
        candidate = os.path.join(output_dir, target)
        return candidate if os.path.isdir(candidate) else output_dir

    def _run_pypff_extract(self, input_file, output_dir):
        target = self.target_edit.text().strip() or None
        thread = PypffExtractThread(input_file, output_dir,
                                    target_basename=target)

        progress_dlg = QProgressDialog(
            'Extracting with pypff...', 'Cancel', 0, 100, self)
        progress_dlg.setWindowModality(Qt.WindowModal)
        progress_dlg.setMinimumDuration(0)
        progress_dlg.setAutoClose(False)
        progress_dlg.setAutoReset(False)
        progress_dlg.setValue(0)

        state = {'ok': False}

        def on_progress(done, total):
            if total > 0:
                progress_dlg.setValue(int(done * 100 / total))
                progress_dlg.setLabelText(f'Extracting {done}/{total}...')

        def on_status(msg):
            progress_dlg.setLabelText(msg)

        def on_done(ok):
            state['ok'] = ok
            progress_dlg.reset()
            progress_dlg.close()

        def on_error(msg):
            QMessageBox.critical(self, 'Extraction Error', msg)

        thread.progress.connect(on_progress)
        thread.status.connect(on_status)
        thread.finished.connect(on_done)
        thread.error_message.connect(on_error)
        progress_dlg.canceled.connect(thread.requestInterruption)

        thread.start()
        progress_dlg.exec_()
        thread.wait(5000)

        if state['ok']:
            QMessageBox.information(
                self, 'Extraction Complete',
                'pypff extraction finished successfully.')
            self.result_root = self._effective_result_root(output_dir)
            self.accept()
        else:
            # error_message dialog already shown if it fired
            pass

# ----------------------------------------------------------------------
# Parser helpers
def parse_outlook_headers(path):
    data = {
        'subject': '', 'sender_name': '', 'date_utc': None,
        'size': 0, 'has_attachments': False, 'conversation_topic': ''
    }
    try:
        with open(path, 'r', encoding='utf-8', errors='ignore') as f:
            for line in f:
                line = line.strip()
                if not line:
                    continue
                if line.startswith('Subject:'):
                    data['subject'] = line.split(':', 1)[1].strip()
                elif line.startswith('Sender name:'):
                    data['sender_name'] = line.split(':', 1)[1].strip()
                elif line.startswith('Client submit time:'):
                    time_str = line.split(':', 1)[1].strip()
                    try:
                        dt_str = time_str.replace(' UTC', '').split('.')[0]
                        dt = datetime.datetime.strptime(dt_str, '%b %d, %Y %H:%M:%S')
                        data['date_utc'] = dt.isoformat()
                    except:
                        pass
                elif line.startswith('Size:'):
                    size_str = line.split(':', 1)[1].strip()
                    try:
                        data['size'] = int(size_str)
                    except:
                        pass
                elif line.startswith('Flags:'):
                    if 'Has attachments' in line:
                        data['has_attachments'] = True
                elif line.startswith('Conversation topic:'):
                    data['conversation_topic'] = line.split(':', 1)[1].strip()
    except:
        pass
    return data

def parse_internet_headers(path):
    data = {'from_email': '', 'to_emails': '', 'cc_emails': ''}
    try:
        with open(path, 'r', encoding='utf-8', errors='ignore') as f:
            content = f.read()
        from_match = re.search(r'^From:\s*(.+?)(?:\r?\n\s+.*?)*?(?=\r?\n\S+:)', content, re.MULTILINE | re.DOTALL)
        if from_match:
            from_line = from_match.group(1).replace('\r\n', ' ').replace('\n', ' ').strip()
            email_match = re.search(r'<([^>]+)>', from_line)
            if email_match:
                data['from_email'] = email_match.group(1)
            else:
                email_match2 = re.search(r'[\w\.-]+@[\w\.-]+', from_line)
                if email_match2:
                    data['from_email'] = email_match2.group(0)
        to_match = re.search(r'^To:\s*(.+?)(?:\r?\n\s+.*?)*?(?=\r?\n\S+:)', content, re.MULTILINE | re.DOTALL)
        if to_match:
            to_line = to_match.group(1).replace('\r\n', ' ').replace('\n', ' ').strip()
            emails = re.findall(r'[\w\.-]+@[\w\.-]+', to_line)
            data['to_emails'] = '; '.join(emails)
        cc_match = re.search(r'^Cc:\s*(.+?)(?:\r?\n\s+.*?)*?(?=\r?\n\S+:)', content, re.MULTILINE | re.DOTALL)
        if cc_match:
            cc_line = cc_match.group(1).replace('\r\n', ' ').replace('\n', ' ').strip()
            emails = re.findall(r'[\w\.-]+@[\w\.-]+', cc_line)
            data['cc_emails'] = '; '.join(emails)
    except:
        pass
    return data

# ----------------------------------------------------------------------
# Scanner thread for extracted folders.
#
# Walks the entire selected root directory and rebuilds the folder tree
# from the directory layout. Any directory that contains OutlookHeaders.txt
# (or Message.html) is treated as a message folder.
class ScannerThread(QThread):
    progress = pyqtSignal(int, int)
    status = pyqtSignal(str)
    finished = pyqtSignal(bool)

    def __init__(self, root_path, db_path):
        super().__init__()
        self.root_path = os.path.normpath(os.path.abspath(root_path))
        self.db_path = db_path
        self.db = None

    def run(self):
        try:
            self.db = MailDB(self.db_path)
            self.db.clear_all()
            self.status.emit('Scanning folders...')

            # -------- 1. Find every message folder beneath the root --------
            message_folders = []
            for dirpath, dirnames, filenames in os.walk(self.root_path):
                # Skip the cache directory itself
                if os.path.normpath(dirpath) == os.path.normpath(
                        os.path.dirname(self.db_path)):
                    # Only skip if it's exactly the cache dir, not a
                    # legitimate folder that happens to contain the DB
                    if 'mail_cache.db' in filenames:
                        continue
                if 'OutlookHeaders.txt' in filenames or \
                        ('Message.html' in filenames and
                         'InternetHeaders.txt' in filenames):
                    message_folders.append(dirpath)

            if not message_folders:
                self.status.emit('No messages found in this folder.')
                self.finished.emit(False)
                return

            # -------- 2. Collect every directory on the path to a message --
            folder_paths = set()
            for msg_path in message_folders:
                p = os.path.dirname(msg_path)
                while p and len(p) >= len(self.root_path) and \
                        p.startswith(self.root_path):
                    folder_paths.add(p)
                    parent = os.path.dirname(p)
                    if parent == p:
                        break
                    p = parent

            self.status.emit('Building folder tree...')

            # -------- 3. Insert the root folder --------
            root_name = os.path.basename(self.root_path) or self.root_path
            root_id = self.db.insert_folder(root_name, None, self.root_path)
            folder_id_map = {self.root_path: root_id}

            # -------- 4. Insert intermediate folders from shallow to deep --
            sorted_dirs = sorted(folder_paths,
                                 key=lambda x: x.count(os.sep))
            for folder_path in sorted_dirs:
                parent_path = os.path.dirname(folder_path)
                parent_id = folder_id_map.get(parent_path)
                if parent_id is None:
                    # Parent wasn't recorded — attach to root
                    parent_id = root_id
                folder_name = os.path.basename(folder_path) or folder_path
                fid = self.db.insert_folder(folder_name, parent_id, folder_path)
                if fid is not None:
                    folder_id_map[folder_path] = fid

            # -------- 5. Insert messages --------
            total = len(message_folders)
            self.status.emit(f'Processing {total} messages...')
            for idx, msg_folder in enumerate(message_folders):
                self.progress.emit(idx + 1, total)
                parent_dir = os.path.dirname(msg_folder)
                folder_id = folder_id_map.get(parent_dir)
                if folder_id is None:
                    folder_id = root_id

                outlook_path = os.path.join(msg_folder, 'OutlookHeaders.txt')
                outlook_data = parse_outlook_headers(outlook_path)
                internet_path = os.path.join(msg_folder, 'InternetHeaders.txt')
                internet_data = parse_internet_headers(internet_path) \
                    if os.path.exists(internet_path) else {}
                html_path = os.path.join(msg_folder, 'Message.html')
                if not os.path.exists(html_path):
                    html_path = ''

                msg_id = self.db.insert_message(
                    folder_id,
                    outlook_data['subject'],
                    outlook_data['sender_name'],
                    internet_data.get('from_email', ''),
                    outlook_data['date_utc'],
                    outlook_data['size'],
                    outlook_data['has_attachments'],
                    html_path,
                    outlook_path,
                    internet_path if os.path.exists(internet_path) else '',
                    outlook_data['conversation_topic'],
                    internet_data.get('to_emails', ''),
                    internet_data.get('cc_emails', '')
                )
                att_dir = os.path.join(msg_folder, 'Attachments')
                if os.path.isdir(att_dir):
                    for fname in os.listdir(att_dir):
                        full = os.path.join(att_dir, fname)
                        if os.path.isfile(full):
                            size = os.path.getsize(full)
                            self.db.insert_attachment(msg_id, fname, full, size)

            self.db.set_meta('mode', 'extracted')
            self.db.conn.commit()
            self.status.emit('Scan complete.')
            self.finished.emit(True)
        except Exception as e:
            import traceback
            traceback.print_exc()
            self.status.emit(f'Error: {str(e)}')
            self.finished.emit(False)
        finally:
            if self.db:
                self.db.close()

# ----------------------------------------------------------------------
# Direct scanner – metadata-only index of an OST/PST via pypff
class DirectScannerThread(QThread):
    """Fast metadata-only scanner. Does NOT read bodies or attachments."""
    progress = pyqtSignal(int, int)
    status = pyqtSignal(str)
    finished = pyqtSignal(bool)
    error_message = pyqtSignal(str)

    def __init__(self, ost_path, cache_dir):
        super().__init__()
        self.ost_path = ost_path
        self.cache_dir = cache_dir
        self.db = None
        self.pff = None
        self._att_name_method = None
        self._att_size_method = None
        self._counter = {'done': 0, 'total': 1, 'failed': 0}

    # ---------- helpers ----------
    @staticmethod
    def _safe_name(name, fallback='unnamed'):
        if name is None:
            return fallback
        try:
            s = str(name)
        except Exception:
            return fallback
        s = re.sub(r'[<>:"/\\|?*\x00-\x1f]', '_', s).strip(' .')
        return s[:180] or fallback

    def _probe_att_name(self, att, idx):
        fallback = f'attachment_{idx}'
        if self._att_name_method is None:
            for name in ('get_name', 'name', 'get_filename', 'filename',
                         'get_long_filename', 'long_filename'):
                try:
                    v = getattr(att, name)
                    val = v() if callable(v) else v
                    if val:
                        self._att_name_method = name
                        return self._safe_name(val, fallback)
                except Exception:
                    continue
            self._att_name_method = ''
            return fallback
        if self._att_name_method == '':
            return fallback
        try:
            v = getattr(att, self._att_name_method)
            val = v() if callable(v) else v
            return self._safe_name(val, fallback)
        except Exception:
            return fallback

    def _probe_att_size(self, att):
        if self._att_size_method is None:
            for name in ('get_size', 'size'):
                try:
                    v = getattr(att, name)
                    val = v() if callable(v) else v
                    self._att_size_method = name
                    try:
                        return int(val or 0)
                    except (TypeError, ValueError):
                        return 0
                except Exception:
                    continue
            self._att_size_method = ''
            return 0
        if self._att_size_method == '':
            return 0
        try:
            v = getattr(att, self._att_size_method)
            val = v() if callable(v) else v
            return int(val or 0)
        except Exception:
            return 0

    @staticmethod
    def _email_first(raw, field):
        if not raw:
            return ''
        try:
            if isinstance(raw, str):
                raw = raw.encode('utf-8', errors='replace')
            m = re.search(rb'^' + field.encode() + rb':\s*(.+?)(?:\r?\n(?![ \t])|\Z)',
                          raw, re.MULTILINE | re.DOTALL | re.IGNORECASE)
            if not m:
                return ''
            line = m.group(1).decode('utf-8', errors='ignore')
            line = line.replace('\r', ' ').replace('\n', ' ')
            em = re.search(r'[\w\.\-\+]+@[\w\.\-]+', line)
            return em.group(0) if em else ''
        except Exception:
            return ''

    @staticmethod
    def _emails_all(raw, field):
        if not raw:
            return ''
        try:
            if isinstance(raw, str):
                raw = raw.encode('utf-8', errors='replace')
            m = re.search(rb'^' + field.encode() + rb':\s*(.+?)(?:\r?\n(?![ \t])|\Z)',
                          raw, re.MULTILINE | re.DOTALL | re.IGNORECASE)
            if not m:
                return ''
            line = m.group(1).decode('utf-8', errors='ignore')
            line = line.replace('\r', ' ').replace('\n', ' ')
            return '; '.join(re.findall(r'[\w\.\-\+]+@[\w\.\-]+', line))
        except Exception:
            return ''

    def _count_messages(self, folder):
        total = 0
        try:
            total += folder.get_number_of_sub_messages()
        except Exception:
            pass
        try:
            n = folder.get_number_of_sub_folders()
        except Exception:
            n = 0
        for i in range(n):
            try:
                total += self._count_messages(folder.get_sub_folder(i))
            except Exception:
                pass
        return total

    # ---------- run ----------
    def run(self):
        if not HAS_PYPFF:
            self.status.emit('pypff is not installed.')
            self.finished.emit(False)
            return
        try:
            os.makedirs(self.cache_dir, exist_ok=True)
            db_path = os.path.join(self.cache_dir, 'mail_cache.db')
            self.db = MailDB(db_path)
            self.db.clear_all()

            self.status.emit('Opening OST/PST...')
            try:
                self.pff, _effective = open_ost_locked_safe(self.ost_path, print)
            except OstLockedError as e:
                self.error_message.emit(str(e))
                self.finished.emit(False)
                return

            self.status.emit('Counting messages...')
            root = self.pff.get_root_folder()
            total = self._count_messages(root) or 1
            self._counter = {'done': 0, 'total': total, 'failed': 0}
            self.status.emit(f'Indexing {total} messages (metadata only)...')

            root_name = os.path.splitext(os.path.basename(self.ost_path))[0] or 'Mailbox'
            root_id = self.db.insert_folder(root_name, None, self.cache_dir)
            self._walk(root, root_id, '')

            self.db.set_meta('source_path', os.path.abspath(self.ost_path))
            try:
                st = os.stat(self.ost_path)
                self.db.set_meta('source_mtime', str(st.st_mtime))
                self.db.set_meta('source_size', str(st.st_size))
            except OSError:
                pass
            self.db.set_meta('mode', 'direct')
            self.db.conn.commit()
            self.status.emit('Index complete.')
            self.finished.emit(True)
        except Exception as e:
            import traceback
            traceback.print_exc()
            self.status.emit(f'Error: {e}')
            self.finished.emit(False)
        finally:
            if self.pff:
                try:
                    self.pff.close()
                except Exception:
                    pass
            if self.db:
                self.db.close()

    def _walk(self, folder, parent_id, pff_path):
        try:
            n_msg = folder.get_number_of_sub_messages()
        except Exception:
            n_msg = 0
        for i in range(n_msg):
            try:
                self._index_message(folder.get_sub_message(i), parent_id,
                                    pff_path, i)
            except Exception as e:
                self._counter['failed'] += 1
                print(f'  ! message {i}: {e}')
            self._counter['done'] += 1
            self.progress.emit(self._counter['done'], self._counter['total'])
            if self._counter['done'] % 500 == 0:
                try:
                    self.db.conn.commit()
                except Exception:
                    pass

        try:
            n_sub = folder.get_number_of_sub_folders()
        except Exception:
            n_sub = 0
        for i in range(n_sub):
            try:
                sub = folder.get_sub_folder(i)
                sub_name = sub.get_name() or f'Folder {i}'
            except Exception:
                continue
            child_pff = f'{pff_path}.{i}' if pff_path else str(i)
            vpath = os.path.join(self.cache_dir, 'virtual',
                                 child_pff.replace('.', '_'))
            sub_id = self.db.insert_folder(sub_name, parent_id, vpath)
            if sub_id is None:
                continue
            self._walk(sub, sub_id, child_pff)

    def _index_message(self, msg, folder_id, pff_path, msg_idx):
        def get(fn_name, default=''):
            try:
                fn = getattr(msg, fn_name)
                v = fn() if callable(fn) else fn
                return v if v is not None else default
            except Exception:
                return default

        subject = str(get('get_subject', '') or '').strip()
        sender_name = str(get('get_sender_name', '') or '').strip()
        conversation_topic = str(get('get_conversation_topic', '') or '')

        date_utc = None
        try:
            dt = msg.get_client_submit_time() or msg.get_delivery_time()
            if dt:
                date_utc = dt.isoformat()
        except Exception:
            pass

        try:
            msg_size = int(msg.get_size() or 0)
        except Exception:
            msg_size = 0

        sender_email = ''
        to_emails = ''
        cc_emails = ''
        try:
            headers_raw = msg.get_transport_headers()
            if headers_raw:
                sender_email = self._email_first(headers_raw, 'From')
                to_emails = self._emails_all(headers_raw, 'To')
                cc_emails = self._emails_all(headers_raw, 'Cc')
        except Exception:
            pass

        try:
            n_att = int(msg.get_number_of_attachments() or 0)
        except Exception:
            n_att = 0
        has_att = n_att > 0

        msg_id = self.db.insert_message(
            folder_id, subject, sender_name, sender_email, date_utc,
            msg_size, has_att, '', '', '', conversation_topic, to_emails, cc_emails)

        try:
            self.db.conn.execute(
                'UPDATE messages SET pff_folder_path = ?, pff_msg_idx = ? WHERE id = ?',
                (pff_path, msg_idx, msg_id))
        except Exception:
            pass

        for j in range(n_att):
            try:
                att = msg.get_attachment(j)
                name = self._probe_att_name(att, j)
                size = self._probe_att_size(att)
                self.db.conn.execute(
                    'INSERT INTO attachments (message_id, filename, filepath, size) '
                    'VALUES (?,?,?,?)',
                    (msg_id, name, '', size))
            except Exception as e:
                print(f'  ! attachment meta {j} of msg {msg_id}: {e}')

# ----------------------------------------------------------------------
# pypff-based extractor – writes a pffexport-compatible folder tree
class PypffExtractThread(QThread):
    """
    Walks the whole pypff tree and writes an on-disk layout that mirrors
    the source mailbox exactly (Root - Mailbox, Root - Public, …).

    Layout:
        <out_dir>/<target>/<pff structure>/
            <Folder>/
                Message NNNNNNNN/
                    Message.html
                    OutlookHeaders.txt
                    InternetHeaders.txt
                    Attachments/
    """
    progress = pyqtSignal(int, int)
    status = pyqtSignal(str)
    finished = pyqtSignal(bool)
    error_message = pyqtSignal(str)

    def __init__(self, ost_path, out_dir, target_basename=None):
        super().__init__()
        self.ost_path = ost_path
        self.out_dir = out_dir
        self.target = target_basename or \
            os.path.splitext(os.path.basename(ost_path))[0] or 'mailbox'
        self.result_root = None
        self._done = 0
        self._total = 1

    # ---------------- helpers ----------------
    @staticmethod
    def _safe(name, fallback='unnamed'):
        if name is None:
            return fallback
        try:
            s = re.sub(r'[<>:"/\\|?*\x00-\x1f]', '_', str(name)).strip(' .')
        except Exception:
            return fallback
        return s[:180] or fallback

    @staticmethod
    def _to_bytes(raw):
        if raw is None:
            return b''
        if isinstance(raw, bytes):
            return raw
        if isinstance(raw, str):
            return raw.encode('utf-8', errors='replace')
        try:
            return bytes(raw)
        except Exception:
            return b''

    @staticmethod
    def _decode(b):
        if b is None:
            return ''
        if isinstance(b, str):
            return b
        try:
            return b.decode('utf-8')
        except UnicodeDecodeError:
            return b.decode('latin-1', errors='replace')

    @staticmethod
    def _strip_mime_preamble(text):
        if not text:
            return ''
        m = re.match(r'^.*?\r?\n\r?\n', text, re.DOTALL)
        if m and ':' in text[:m.end()] and '<' not in text[:m.end()]:
            return text[m.end():]
        return text

    def _count(self, folder):
        n = 0
        try:
            n += folder.get_number_of_sub_messages()
        except Exception:
            pass
        try:
            ns = folder.get_number_of_sub_folders()
        except Exception:
            ns = 0
        for i in range(ns):
            try:
                n += self._count(folder.get_sub_folder(i))
            except Exception:
                pass
        return n

    # ---------------- run ----------------
    def run(self):
        if not HAS_PYPFF:
            self.status.emit('pypff is not installed.')
            self.finished.emit(False)
            return
        pff = None
        try:
            self.status.emit('Opening OST/PST (pypff)...')
            try:
                pff, _effective = open_ost_locked_safe(self.ost_path, print)
            except OstLockedError as e:
                self.error_message.emit(str(e))
                self.finished.emit(False)
                return

            root = pff.get_root_folder()

            self.status.emit('Counting messages...')
            self._total = self._count(root) or 1
            self._done = 0

            target_root = os.path.join(self.out_dir, self.target)
            os.makedirs(target_root, exist_ok=True)

            self.status.emit(
                f'Extracting {self._total} messages to {target_root}...')
            self._walk(root, target_root)

            if self.isInterruptionRequested():
                self.status.emit('Cancelled.')
                self.finished.emit(False)
                return

            self.result_root = target_root
            self.status.emit('Extraction complete.')
            self.finished.emit(True)
        except Exception as e:
            import traceback
            traceback.print_exc()
            self.status.emit(f'Error: {e}')
            self.finished.emit(False)
        finally:
            if pff:
                try:
                    pff.close()
                except Exception:
                    pass

    def _walk(self, folder, out_dir):
        try:
            n_msg = folder.get_number_of_sub_messages()
        except Exception:
            n_msg = 0
        for i in range(n_msg):
            if self.isInterruptionRequested():
                return
            try:
                self._extract_message(folder.get_sub_message(i), out_dir, i + 1)
            except Exception as e:
                print(f'  ! message {i}: {e}')
            self._done += 1
            self.progress.emit(self._done, self._total)

        try:
            n_sub = folder.get_number_of_sub_folders()
        except Exception:
            n_sub = 0
        for i in range(n_sub):
            if self.isInterruptionRequested():
                return
            try:
                sub = folder.get_sub_folder(i)
                sub_name = self._safe(sub.get_name(), f'Folder {i}')
            except Exception:
                continue
            sub_dir = os.path.join(out_dir, sub_name)
            base = sub_dir
            j = 1
            while os.path.exists(sub_dir):
                sub_dir = f'{base}_{j}'
                j += 1
            try:
                os.makedirs(sub_dir, exist_ok=True)
            except OSError as e:
                print(f'  ! could not create {sub_dir}: {e}')
                continue
            self._walk(sub, sub_dir)

    def _extract_message(self, msg, out_dir, idx):
        msg_dir = os.path.join(out_dir, f'Message {idx:08d}')
        os.makedirs(msg_dir, exist_ok=True)

        subject = ''
        try:
            subject = (msg.get_subject() or '').strip()
        except Exception:
            pass
        sender = ''
        try:
            sender = (msg.get_sender_name() or '').strip()
        except Exception:
            pass
        topic = ''
        try:
            topic = msg.get_conversation_topic() or ''
        except Exception:
            pass

        html_text = ''
        try:
            raw = msg.get_html_body()
            if raw:
                html_text = self._strip_mime_preamble(self._decode(raw))
        except Exception:
            pass
        if not html_text:
            try:
                raw = msg.get_plain_text_body()
                if raw:
                    plain = self._decode(raw)
                    esc = (plain.replace('&', '&amp;')
                                .replace('<', '&lt;')
                                .replace('>', '&gt;'))
                    html_text = f'<pre>{esc}</pre>'
            except Exception:
                pass
        if html_text:
            try:
                with open(os.path.join(msg_dir, 'Message.html'), 'w',
                          encoding='utf-8', errors='replace') as fh:
                    fh.write(html_text)
            except OSError:
                pass

        try:
            with open(os.path.join(msg_dir, 'OutlookHeaders.txt'), 'w',
                      encoding='utf-8', errors='replace') as fh:
                fh.write(f'Subject: {subject}\n')
                fh.write(f'Sender name: {sender}\n')
                if topic:
                    fh.write(f'Conversation topic: {topic}\n')
                try:
                    dt = msg.get_client_submit_time() or msg.get_delivery_time()
                    if dt:
                        fh.write(f'Client submit time: {dt.isoformat()}\n')
                except Exception:
                    pass
                try:
                    size = int(msg.get_size() or 0)
                    if size:
                        fh.write(f'Size: {size}\n')
                except Exception:
                    pass
                try:
                    if int(msg.get_number_of_attachments() or 0) > 0:
                        fh.write('Flags: Has attachments\n')
                except Exception:
                    pass
        except OSError:
            pass

        try:
            headers = msg.get_transport_headers()
            if headers:
                with open(os.path.join(msg_dir, 'InternetHeaders.txt'),
                          'wb') as fh:
                    fh.write(self._to_bytes(headers))
        except Exception:
            pass

        try:
            n_att = int(msg.get_number_of_attachments() or 0)
        except Exception:
            n_att = 0
        if n_att:
            att_dir = os.path.join(msg_dir, 'Attachments')
            os.makedirs(att_dir, exist_ok=True)
            for j in range(n_att):
                try:
                    att = msg.get_attachment(j)
                    fname = None
                    for name in ('get_name', 'name', 'get_filename',
                                 'filename', 'get_long_filename',
                                 'long_filename'):
                        try:
                            v = getattr(att, name)
                            val = v() if callable(v) else v
                            if val:
                                fname = self._safe(val, f'attachment_{j}')
                                break
                        except Exception:
                            continue
                    if not fname:
                        fname = f'attachment_{j}'
                    base, ext = os.path.splitext(fname)
                    target = os.path.join(att_dir, fname)
                    k = 1
                    while os.path.exists(target):
                        target = os.path.join(att_dir, f'{base}_{k}{ext}')
                        k += 1
                    data = b''
                    try:
                        size = int(att.get_size() or 0)
                        data = self._to_bytes(att.read_buffer(size))
                    except Exception:
                        for mname in ('read_buffer', 'get_data', 'read',
                                      'get_contents', 'get_buffer', 'read_data'):
                            try:
                                fn = getattr(att, mname)
                                data = self._to_bytes(
                                    fn() if callable(fn) else fn)
                                if data:
                                    break
                            except Exception:
                                continue
                    with open(target, 'wb') as fh:
                        fh.write(data)
                except Exception as e:
                    print(f'  ! attachment {j}: {e}')

# ----------------------------------------------------------------------
# On-demand body/attachment reader for direct mode
class PffBodyLoader:
    """Holds one open pypff handle; serves bodies and attachment bytes."""

    def __init__(self, ost_path):
        self.ost_path = ost_path
        self.pff = None

    def open(self):
        if self.pff is not None:
            return True
        try:
            self.pff, _effective = open_ost_locked_safe(self.ost_path, print)
            return True
        except OstLockedError as e:
            print(f'PffBodyLoader.open: {e}')
            self.pff = None
            return False
        except Exception as e:
            print(f'PffBodyLoader.open failed: {e}')
            self.pff = None
            return False

    def close(self):
        if self.pff is not None:
            try:
                self.pff.close()
            except Exception:
                pass
            self.pff = None

    @staticmethod
    def _decode(b):
        if b is None:
            return ''
        if isinstance(b, str):
            return b
        try:
            return b.decode('utf-8')
        except UnicodeDecodeError:
            return b.decode('latin-1', errors='replace')

    @staticmethod
    def _strip_preamble(text):
        if not text:
            return ''
        m = re.match(r'^.*?\r?\n\r?\n', text, re.DOTALL)
        if m and ':' in text[:m.end()] and '<' not in text[:m.end()]:
            return text[m.end():]
        return text

    def _navigate(self, folder_path, msg_idx):
        folder = self.pff.get_root_folder()
        if folder_path:
            for seg in folder_path.split('.'):
                if not seg:
                    continue
                folder = folder.get_sub_folder(int(seg))
        return folder.get_sub_message(int(msg_idx))

    def get_html(self, folder_path, msg_idx):
        if self.pff is None and not self.open():
            return '<p><i>Could not open OST/PST.</i></p>'
        try:
            msg = self._navigate(folder_path, msg_idx)
        except Exception as e:
            return f'<p><i>Cannot locate message: {e}</i></p>'
        try:
            raw = msg.get_html_body()
            if raw:
                return self._strip_preamble(self._decode(raw))
        except Exception:
            pass
        try:
            raw = msg.get_plain_text_body()
            if raw:
                plain = self._decode(raw)
                esc = (plain.replace('&', '&amp;')
                            .replace('<', '&lt;')
                            .replace('>', '&gt;'))
                return f'<pre>{esc}</pre>'
        except Exception:
            pass
        return '<p><i>No body content.</i></p>'

    def get_attachment_bytes(self, folder_path, msg_idx, att_idx):
        if self.pff is None and not self.open():
            raise RuntimeError('OST/PST not open')
        msg = self._navigate(folder_path, msg_idx)
        att = msg.get_attachment(int(att_idx))
        for takes_size in (True, False):
            try:
                data = att.read_buffer(att.get_size()) if takes_size else att.read_buffer()
                if data is not None:
                    if isinstance(data, str):
                        return data.encode('latin-1')
                    return bytes(data)
            except Exception:
                continue
        for name in ('get_data', 'read', 'get_contents', 'get_buffer', 'read_data'):
            try:
                fn = getattr(att, name)
                data = fn() if callable(fn) else fn
                if data is not None:
                    if isinstance(data, str):
                        return data.encode('latin-1')
                    return bytes(data)
            except Exception:
                continue
        return b''

# ----------------------------------------------------------------------
# Main Window
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.db = None
        self.current_folder_id = None
        self.current_message_id = None
        self.root_path = None
        self.source_ost = None
        self.mode = 'extracted'
        self.pff_loader = None
        self.scale = SCALE_FACTOR
        self.init_ui()
        self.load_settings()

    def init_ui(self):
        self.setWindowTitle('OST Mail Viewer')
        self.setGeometry(100, 100, 1400, 900)

        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QVBoxLayout(central)
        main_layout.setContentsMargins(0, 0, 0, 0)

        # ---------------- Top bar ----------------
        top_bar = QWidget()
        top_bar_layout = QHBoxLayout(top_bar)
        top_bar_layout.setContentsMargins(5, 5, 5, 5)

        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText('Search subject or sender...')
        self.search_input.textChanged.connect(self.on_search)
        top_bar_layout.addWidget(self.search_input)

        self.attach_check = QCheckBox('Only with attachments')
        self.attach_check.stateChanged.connect(self.on_search)
        top_bar_layout.addWidget(self.attach_check)

        clear_btn = QPushButton('Clear')
        clear_btn.clicked.connect(self.clear_search)
        top_bar_layout.addWidget(clear_btn)

        top_bar_layout.addStretch()

        open_extracted_btn = QPushButton('Open Extracted Folder...')
        open_extracted_btn.setToolTip(
            'Open an already-extracted pffexport output folder')
        open_extracted_btn.clicked.connect(self.select_root_folder)
        top_bar_layout.addWidget(open_extracted_btn)

        direct_btn = QPushButton('Open OST/PST (direct)...')
        direct_btn.setToolTip('Browse an OST/PST in place (metadata only)')
        direct_btn.clicked.connect(self.open_ost_direct)
        if not HAS_PYPFF:
            direct_btn.setEnabled(False)
            direct_btn.setToolTip('pypff not installed')
        top_bar_layout.addWidget(direct_btn)

        extract_btn = QPushButton('Extract OST/PST...')
        extract_btn.setToolTip('Extract to disk with pypff or pffexport')
        extract_btn.clicked.connect(self.open_extract_dialog)
        top_bar_layout.addWidget(extract_btn)

        main_layout.addWidget(top_bar)

        # ---------------- Main horizontal splitter ----------------
        h_splitter = QSplitter(Qt.Horizontal)

        left_v_splitter = QSplitter(Qt.Vertical)

        self.folder_tree = QTreeView()
        self.folder_tree.setHeaderHidden(True)
        self.folder_tree.clicked.connect(self.on_folder_selected)
        left_v_splitter.addWidget(self.folder_tree)

        self.msg_folder_view = QTreeWidget()
        self.msg_folder_view.setHeaderLabels(['Name', 'Size'])
        self.msg_folder_view.setIndentation(10)
        self.msg_folder_view.setContextMenuPolicy(Qt.CustomContextMenu)
        self.msg_folder_view.customContextMenuRequested.connect(self.msg_folder_context_menu)
        self.msg_folder_view.itemDoubleClicked.connect(self.open_msg_folder_item)
        left_v_splitter.addWidget(self.msg_folder_view)

        left_v_splitter.setSizes([400, 300])
        h_splitter.addWidget(left_v_splitter)

        right_v_splitter = QSplitter(Qt.Vertical)

        self.message_table = QTableView()
        self.message_table.setSortingEnabled(True)
        self.message_table.horizontalHeader().setStretchLastSection(True)
        self.message_table.clicked.connect(self.on_message_selected)
        right_v_splitter.addWidget(self.message_table)

        preview_widget = QWidget()
        preview_layout = QVBoxLayout(preview_widget)
        preview_layout.setContentsMargins(0, 0, 0, 0)

        self.webview = QWebEngineView()
        preview_layout.addWidget(self.webview, 1)

        right_v_splitter.addWidget(preview_widget)
        right_v_splitter.setSizes([400, 500])

        h_splitter.addWidget(right_v_splitter)
        h_splitter.setSizes([250, 1150])
        main_layout.addWidget(h_splitter, 1)

        # ---------------- Menus ----------------
        menubar = self.menuBar()
        file_menu = menubar.addMenu('File')
        open_action = QAction('Open Extracted Folder', self)
        open_action.triggered.connect(self.select_root_folder)
        file_menu.addAction(open_action)

        direct_action = QAction('Open OST/PST (direct)...', self)
        direct_action.triggered.connect(self.open_ost_direct)
        if not HAS_PYPFF:
            direct_action.setEnabled(False)
        file_menu.addAction(direct_action)

        extract_action = QAction('Extract OST/PST...', self)
        extract_action.triggered.connect(self.open_extract_dialog)
        file_menu.addAction(extract_action)

        rescan_action = QAction('Rescan', self)
        rescan_action.triggered.connect(self.rescan)
        file_menu.addAction(rescan_action)

        file_menu.addSeparator()
        exit_action = QAction('Exit', self)
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)

        view_menu = menubar.addMenu('View')
        zoom_in = QAction('Zoom In', self)
        zoom_in.triggered.connect(lambda: self.change_scale(1.1))
        view_menu.addAction(zoom_in)
        zoom_out = QAction('Zoom Out', self)
        zoom_out.triggered.connect(lambda: self.change_scale(0.9))
        view_menu.addAction(zoom_out)
        reset_zoom = QAction('Reset Zoom', self)
        reset_zoom.triggered.connect(lambda: self.change_scale(1.0))
        view_menu.addAction(reset_zoom)

        self.attachments_dock = QDockWidget('All Attachments', self)
        self.attachments_dock.setAllowedAreas(Qt.RightDockWidgetArea | Qt.BottomDockWidgetArea)
        dock_widget = QWidget()
        dock_layout = QVBoxLayout(dock_widget)
        self.att_search = QLineEdit()
        self.att_search.setPlaceholderText('Search attachments...')
        self.att_search.textChanged.connect(self.refresh_all_attachments)
        dock_layout.addWidget(self.att_search)
        self.att_table = QTableView()
        self.att_table.setSortingEnabled(True)
        self.att_table.doubleClicked.connect(self.open_attachment_from_table)
        self.att_table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.att_table.customContextMenuRequested.connect(self.att_context_menu)
        dock_layout.addWidget(self.att_table)
        self.attachments_dock.setWidget(dock_widget)
        self.addDockWidget(Qt.RightDockWidgetArea, self.attachments_dock)
        self.attachments_dock.hide()

        view_menu.addSeparator()
        show_att_action = QAction('Show All Attachments', self)
        show_att_action.setCheckable(True)
        show_att_action.toggled.connect(self.toggle_attachments_dock)
        view_menu.addAction(show_att_action)

        self.statusBar().showMessage('Ready')

        # ---------------- Models ----------------
        self.folder_model = QStandardItemModel()
        self.folder_tree.setModel(self.folder_model)

        self.message_model = QStandardItemModel()
        self.message_model.setHorizontalHeaderLabels(['Subject', 'Sender', 'Date', 'Size', 'Attachments'])
        self.message_table.setModel(self.message_model)

        self.att_model = QStandardItemModel()
        self.att_model.setHorizontalHeaderLabels(
            ['Filename', 'Message Subject', 'Sender', 'Date', 'Folder', 'Size'])
        self.att_table.setModel(self.att_model)

    # ------------------------------------------------------------------
    # Direct OST/PST mode
    def open_ost_direct(self):
        if not HAS_PYPFF:
            QMessageBox.critical(
                self, 'pypff missing',
                'pypff is not installed in this Python.\n\n'
                'Install with:\n'
                '    pip install https://github.com/Sygmei/wheels/raw/main/'
                'libpff_python-20211114-cp311-cp311-win_amd64.whl\n'
                'And run with Python 3.11.'
            )
            return

        ost_path, _ = QFileDialog.getOpenFileName(
            self, 'Open OST/PST directly (no extraction)',
            '', 'Outlook files (*.ost *.pst);;All files (*)')
        if not ost_path:
            return

        ost_path = os.path.normpath(os.path.abspath(ost_path))

        self.mode = 'direct'
        self.source_ost = ost_path
        self.root_path = direct_cache_dir(ost_path)
        self.save_settings()
        self.update_window_title()

        if self.pff_loader:
            self.pff_loader.close()
            self.pff_loader = None

        db_path = os.path.join(self.root_path, 'mail_cache.db')
        use_cache = False
        if os.path.exists(db_path):
            try:
                probe = MailDB(db_path)
                cached_mtime = probe.get_meta('source_mtime')
                cached_size = probe.get_meta('source_size')
                cached_mode = probe.get_meta('mode')
                probe.close()
                try:
                    st = os.stat(ost_path)
                    if (cached_mode == 'direct'
                            and cached_mtime == str(st.st_mtime)
                            and cached_size == str(st.st_size)):
                        use_cache = True
                except OSError:
                    pass
            except Exception:
                use_cache = False

        if use_cache:
            self.db = MailDB(db_path)
            self.populate_tree()
            self.pff_loader = PffBodyLoader(ost_path)
            self.statusBar().showMessage(f'Index loaded: {db_path}')
        else:
            if os.path.exists(db_path):
                try:
                    os.remove(db_path)
                except OSError:
                    pass
            self.start_direct_scan(ost_path, self.root_path)

    def start_direct_scan(self, ost_path, cache_dir):
        self.scanner = DirectScannerThread(ost_path, cache_dir)
        self.scanner.progress.connect(self.update_progress)
        self.scanner.status.connect(self.statusBar().showMessage)
        self.scanner.finished.connect(self.on_scan_finished)
        self.scanner.error_message.connect(self._show_error)
        self.scanner.start()
        self.statusBar().showMessage('Reading OST/PST directly...')

    def _show_error(self, message):
        QMessageBox.critical(self, 'Error', message)

    # ------------------------------------------------------------------
    # Extract dialog
    def open_extract_dialog(self):
        dlg = PFFExportDialog(self)
        if dlg.exec_() == QDialog.Accepted:
            reply = QMessageBox.question(
                self, 'Extraction Complete',
                'Extraction finished successfully.\n\n'
                'Would you like to open the extracted folder now?',
                QMessageBox.Yes | QMessageBox.No
            )
            if reply == QMessageBox.Yes:
                self.mode = 'extracted'
                self.source_ost = None
                if self.pff_loader:
                    self.pff_loader.close()
                    self.pff_loader = None
                self.root_path = dlg.result_root
                self.save_settings()
                self.load_db()

    # ------------------------------------------------------------------
    # Settings / DB loading
    def load_settings(self):
        settings = QSettings('OSTMailViewer', 'Settings')
        last_root = settings.value('last_root', '')
        self.mode = settings.value('mode', 'extracted')
        self.source_ost = settings.value('source_ost', '') or None
        if last_root and os.path.isdir(last_root):
            self.root_path = last_root
            self.load_db()
        else:
            self.statusBar().showMessage(
                'Ready — use the toolbar or File menu to open an OST/PST '
                'or an extracted folder.')

    def save_settings(self):
        settings = QSettings('OSTMailViewer', 'Settings')
        settings.setValue('last_root', self.root_path or '')
        settings.setValue('mode', self.mode)
        settings.setValue('source_ost', self.source_ost or '')

    def update_window_title(self):
        if self.mode == 'direct' and self.source_ost:
            self.setWindowTitle(
                f'OST Mail Viewer  –  [Direct] {os.path.basename(self.source_ost)}')
        elif self.root_path:
            self.setWindowTitle(
                f'OST Mail Viewer  –  [Extracted] {self.root_path}')
        else:
            self.setWindowTitle('OST Mail Viewer')

    def select_root_folder(self):
        folder = QFileDialog.getExistingDirectory(
            self, 'Select root folder (containing Root - Mailbox)')
        if folder:
            self.mode = 'extracted'
            self.source_ost = None
            if self.pff_loader:
                self.pff_loader.close()
                self.pff_loader = None
            self.root_path = folder
            self.save_settings()
            self.load_db()

    def load_db(self):
        db_path = os.path.join(self.root_path, 'mail_cache.db')
        if os.path.exists(db_path):
            self.db = MailDB(db_path)
            self.populate_tree()
            self.update_window_title()
            if self.mode == 'direct' and self.source_ost:
                self.pff_loader = PffBodyLoader(self.source_ost)
            self.statusBar().showMessage(f'Loaded from cache: {db_path}')
        else:
            self.scan_and_load()

    def scan_and_load(self):
        db_path = os.path.join(self.root_path, 'mail_cache.db')
        self.scanner = ScannerThread(self.root_path, db_path)
        self.scanner.progress.connect(self.update_progress)
        self.scanner.status.connect(self.statusBar().showMessage)
        self.scanner.finished.connect(self.on_scan_finished)
        self.scanner.start()
        self.statusBar().showMessage('Scanning...')

    def on_scan_finished(self, success):
        if success:
            self.db = MailDB(os.path.join(self.root_path, 'mail_cache.db'))
            self.populate_tree()
            if self.mode == 'direct' and self.source_ost:
                self.pff_loader = PffBodyLoader(self.source_ost)
            self.update_window_title()
            self.statusBar().showMessage('Scan complete.')
        else:
            self.statusBar().showMessage('Scan failed.')

    def update_progress(self, current, total):
        self.statusBar().showMessage(f'Processing {current}/{total}...')

    def rescan(self):
        if not self.root_path:
            return
        db_path = os.path.join(self.root_path, 'mail_cache.db')
        if os.path.exists(db_path):
            try:
                os.remove(db_path)
            except OSError:
                pass
        if self.mode == 'direct' and self.source_ost and os.path.isfile(self.source_ost):
            self.start_direct_scan(self.source_ost, self.root_path)
        else:
            self.scan_and_load()

    # ------------------------------------------------------------------
    # Tree / messages
    def populate_tree(self):
        self.folder_model.clear()
        root_item = self.folder_model.invisibleRootItem()
        folders = self.db.get_folders()
        for f in folders:
            item = QStandardItem(f['name'])
            item.setData(f['id'], Qt.UserRole)
            item.setData(f['full_path'], Qt.UserRole + 1)
            root_item.appendRow(item)
            self.populate_children(item, f['id'])
        self.folder_tree.expandToDepth(1)

    def populate_children(self, parent_item, parent_id):
        children = self.db.get_folders(parent_id)
        for f in children:
            item = QStandardItem(f['name'])
            item.setData(f['id'], Qt.UserRole)
            item.setData(f['full_path'], Qt.UserRole + 1)
            parent_item.appendRow(item)
            self.populate_children(item, f['id'])

    def on_folder_selected(self, index):
        item = self.folder_model.itemFromIndex(index)
        if item:
            folder_id = item.data(Qt.UserRole)
            if folder_id is not None:
                self.current_folder_id = folder_id
                self.load_messages(folder_id)

    def load_messages(self, folder_id, search='', only_attachments=False):
        rows = self.db.get_messages_for_folder(folder_id, search, only_attachments)
        self.message_model.removeRows(0, self.message_model.rowCount())
        for r in rows:
            subject = r['subject'] or ''
            sender = r['sender_name'] or r['sender_email'] or ''
            date_str = ''
            if r['date_utc']:
                try:
                    dt = datetime.datetime.fromisoformat(r['date_utc'])
                    date_str = dt.strftime('%Y-%m-%d %H:%M')
                except:
                    date_str = r['date_utc'][:16] if r['date_utc'] else ''
            size_str = f"{r['size']:,}" if r['size'] else ''
            attach_str = 'Yes' if r['has_attachments'] else ''
            item_subject = QStandardItem(subject)
            item_subject.setData(r['id'], Qt.UserRole)
            item_subject.setEditable(False)
            item_sender = QStandardItem(sender)
            item_sender.setEditable(False)
            item_date = QStandardItem(date_str)
            item_date.setEditable(False)
            item_size = QStandardItem(size_str)
            item_size.setEditable(False)
            item_attach = QStandardItem(attach_str)
            item_attach.setEditable(False)
            self.message_model.appendRow(
                [item_subject, item_sender, item_date, item_size, item_attach])
        self.message_table.resizeColumnsToContents()

    def on_search(self):
        if self.current_folder_id is None:
            return
        search_text = self.search_input.text()
        only_attachments = self.attach_check.isChecked()
        self.load_messages(self.current_folder_id, search_text, only_attachments)

    def clear_search(self):
        self.search_input.clear()
        self.attach_check.setChecked(False)

    def on_message_selected(self, index):
        item = self.message_model.itemFromIndex(index.sibling(index.row(), 0))
        if item:
            msg_id = item.data(Qt.UserRole)
            if msg_id is not None:
                self.current_message_id = msg_id
                self.show_message(msg_id)

    def show_message(self, msg_id):
        row = self.db.get_message_by_id(msg_id)
        if not row:
            return

        if self.mode == 'direct' and self.pff_loader:
            folder_path = row.get('pff_folder_path') or ''
            msg_idx = row.get('pff_msg_idx')
            if msg_idx is None:
                self.webview.setHtml('<p><i>No locator for this message.</i></p>')
            else:
                QApplication.setOverrideCursor(Qt.WaitCursor)
                try:
                    html = self.pff_loader.get_html(folder_path, int(msg_idx))
                finally:
                    QApplication.restoreOverrideCursor()
                self.webview.setHtml(html)
        else:
            html_path = row['message_html_path']
            if html_path and os.path.exists(html_path):
                with open(html_path, 'r', encoding='utf-8', errors='ignore') as f:
                    html = f.read()
                msg_dir = os.path.dirname(html_path)
                base_url = QUrl.fromLocalFile(msg_dir + os.sep)
                self.webview.setHtml(html, base_url)
            else:
                self.webview.setHtml('<p>No HTML content</p>')

        self.populate_message_folder(row)

    # ------------------------------------------------------------------
    # Bottom-left message folder contents
    def populate_message_folder(self, msg_row):
        self.msg_folder_view.clear()

        outlook_item = QTreeWidgetItem(self.msg_folder_view)
        outlook_item.setText(0, 'Outlook Data')
        outlook_item.setExpanded(True)
        attachments_item = QTreeWidgetItem(self.msg_folder_view)
        attachments_item.setText(0, 'Attachments')
        attachments_item.setExpanded(True)

        if self.mode == 'direct':
            for label, key in (('Subject', 'subject'),
                               ('Sender', 'sender_name'),
                               ('From', 'sender_email'),
                               ('To', 'to_emails'),
                               ('Cc', 'cc_emails'),
                               ('Date', 'date_utc')):
                val = msg_row.get(key)
                if val:
                    it = QTreeWidgetItem(outlook_item)
                    it.setText(0, label)
                    it.setText(1, str(val)[:120])
        else:
            outlook_path = msg_row['outlook_headers_path']
            if outlook_path and os.path.exists(outlook_path):
                msg_dir = os.path.dirname(outlook_path)
                try:
                    entries = os.listdir(msg_dir)
                except OSError:
                    entries = []
                for fname in entries:
                    full_path = os.path.join(msg_dir, fname)
                    if os.path.isfile(full_path) and fname.lower() in (
                            'internetheaders.txt', 'outlookheaders.txt',
                            'conversationindex.txt', 'recipients.txt',
                            'message.html'):
                        it = QTreeWidgetItem(outlook_item)
                        it.setText(0, fname)
                        it.setText(1, f'{os.path.getsize(full_path):,}')
                        it.setData(0, Qt.UserRole, full_path)

        for att in self.db.get_attachments_for_message(msg_row['id']):
            it = QTreeWidgetItem(attachments_item)
            it.setText(0, att['filename'])
            it.setText(1, f'{att["size"]:,}')
            it.setData(0, Qt.UserRole, att['filepath'])
            it.setData(0, Qt.UserRole + 2, att['id'])

        self.msg_folder_view.resizeColumnToContents(0)
        self.msg_folder_view.resizeColumnToContents(1)

    def open_msg_folder_item(self, item, column):
        att_id = item.data(0, Qt.UserRole + 2)
        if att_id is not None:
            self.open_attachment_by_id(att_id)
            return
        filepath = item.data(0, Qt.UserRole)
        if filepath and os.path.exists(filepath):
            QDesktopServices.openUrl(QUrl.fromLocalFile(filepath))

    def msg_folder_context_menu(self, pos):
        item = self.msg_folder_view.itemAt(pos)
        if not item:
            return
        att_id = item.data(0, Qt.UserRole + 2)
        filepath = item.data(0, Qt.UserRole)
        if att_id is None and not filepath:
            return
        menu = QMenu()
        open_act = menu.addAction('Open')
        open_loc_act = menu.addAction('Open file location')
        copy_path_act = menu.addAction('Copy file path')
        action = menu.exec_(self.msg_folder_view.mapToGlobal(pos))

        if att_id is not None:
            att = self.db.get_attachment_by_id(att_id)
            path = att['filepath'] if att else ''
        else:
            path = filepath

        if action == open_act:
            if att_id is not None:
                self.open_attachment_by_id(att_id)
            elif path and os.path.exists(path):
                QDesktopServices.openUrl(QUrl.fromLocalFile(path))
        elif action == open_loc_act:
            if path and os.path.exists(path):
                if sys.platform == 'win32':
                    subprocess.Popen(['explorer', '/select,', path])
                else:
                    QDesktopServices.openUrl(QUrl.fromLocalFile(os.path.dirname(path)))
        elif action == copy_path_act:
            QApplication.clipboard().setText(path or '')

    # ------------------------------------------------------------------
    # Attachment actions
    def open_attachment_by_id(self, att_id):
        att = self.db.get_attachment_by_id(att_id)
        if not att:
            return

        if self.mode != 'direct' and att.get('filepath') and os.path.exists(att['filepath']):
            QDesktopServices.openUrl(QUrl.fromLocalFile(att['filepath']))
            return

        if not self.pff_loader:
            QMessageBox.warning(self, 'Error', 'OST/PST is not open.')
            return

        row = self.db.get_message_by_id(att['message_id'])
        if not row or row.get('pff_msg_idx') is None:
            QMessageBox.warning(self, 'Error', 'No locator for this attachment.')
            return

        temp_dir = os.path.join(
            os.environ.get('TEMP') or os.path.expanduser('~/.cache'),
            'OSTMailViewer',
            hashlib.md5(os.path.abspath(self.source_ost or '').encode()).hexdigest()[:12],
            f'{row["id"]:08d}')
        os.makedirs(temp_dir, exist_ok=True)
        safe = re.sub(r'[<>:"/\\|?*\x00-\x1f]', '_', att['filename']) or 'attachment'
        target = os.path.join(temp_dir, safe)

        if not os.path.exists(target):
            QApplication.setOverrideCursor(Qt.WaitCursor)
            try:
                atts = self.db.get_attachments_for_message(row['id'])
                idx = next((i for i, a in enumerate(atts) if a['id'] == att_id), None)
                if idx is None:
                    raise RuntimeError('attachment index not found')
                data = self.pff_loader.get_attachment_bytes(
                    row.get('pff_folder_path') or '', int(row['pff_msg_idx']), idx)
                with open(target, 'wb') as fh:
                    fh.write(data)
            except Exception as e:
                QMessageBox.warning(self, 'Error', f'Could not extract attachment: {e}')
                return
            finally:
                QApplication.restoreOverrideCursor()

        QDesktopServices.openUrl(QUrl.fromLocalFile(target))

    def open_attachment_from_table(self, index):
        att_id = self.att_model.itemFromIndex(index.sibling(index.row(), 0)).data(Qt.UserRole)
        if att_id:
            self.open_attachment_by_id(att_id)

    def att_context_menu(self, pos):
        index = self.att_table.indexAt(pos)
        if not index.isValid():
            return
        att_id = self.att_model.itemFromIndex(index.sibling(index.row(), 0)).data(Qt.UserRole)
        if att_id is None:
            return
        att = self.db.get_attachment_by_id(att_id)
        if not att:
            return
        menu = QMenu()
        open_act = menu.addAction('Open')
        open_loc_act = menu.addAction('Open file location')
        copy_path_act = menu.addAction('Copy file path')
        action = menu.exec_(self.att_table.mapToGlobal(pos))
        if action == open_act:
            self.open_attachment_by_id(att_id)
        elif action == open_loc_act:
            path = att['filepath']
            if path and os.path.exists(path):
                if sys.platform == 'win32':
                    subprocess.Popen(['explorer', '/select,', path])
                else:
                    QDesktopServices.openUrl(QUrl.fromLocalFile(os.path.dirname(path)))
        elif action == copy_path_act:
            QApplication.clipboard().setText(att['filepath'] or '')

    def refresh_all_attachments(self):
        search = self.att_search.text()
        rows = self.db.get_all_attachments(search)
        self.att_model.removeRows(0, self.att_model.rowCount())
        for r in rows:
            subject = r['msg_subject'] or ''
            sender = r['sender_name'] or ''
            date_str = ''
            if r['date_utc']:
                try:
                    dt = datetime.datetime.fromisoformat(r['date_utc'])
                    date_str = dt.strftime('%Y-%m-%d %H:%M')
                except:
                    date_str = r['date_utc'][:16] if r['date_utc'] else ''
            size_str = f"{r['size']:,}" if r['size'] else ''
            item_filename = QStandardItem(r['filename'])
            item_filename.setData(r['id'], Qt.UserRole)
            item_filename.setEditable(False)
            item_subject = QStandardItem(subject)
            item_subject.setEditable(False)
            item_sender = QStandardItem(sender)
            item_sender.setEditable(False)
            item_date = QStandardItem(date_str)
            item_date.setEditable(False)
            item_folder = QStandardItem(r['folder_name'] or '')
            item_folder.setEditable(False)
            item_size = QStandardItem(size_str)
            item_size.setEditable(False)
            self.att_model.appendRow(
                [item_filename, item_subject, item_sender, item_date, item_folder, item_size])
        self.att_table.resizeColumnsToContents()

    def toggle_attachments_dock(self, checked):
        if checked:
            self.attachments_dock.show()
            self.refresh_all_attachments()
        else:
            self.attachments_dock.hide()

    # ------------------------------------------------------------------
    # Zoom
    def change_scale(self, factor):
        global SCALE_FACTOR
        SCALE_FACTOR *= factor
        SCALE_FACTOR = max(0.5, min(2.0, SCALE_FACTOR))
        self.apply_scale()
        self.statusBar().showMessage(f'Zoom: {SCALE_FACTOR:.1f}x')

    def apply_scale(self):
        base_size = int(9 * SCALE_FACTOR)
        font = QFont()
        font.setPointSize(base_size)
        QApplication.setFont(font)
        self.webview.setZoomFactor(SCALE_FACTOR)

    def closeEvent(self, event):
        if self.pff_loader:
            self.pff_loader.close()
            self.pff_loader = None
        if self.db:
            self.db.close()
        super().closeEvent(event)

# ----------------------------------------------------------------------
if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setOrganizationName('OSTMailViewer')
    app.setApplicationName('OSTMailViewer')
    win = MainWindow()
    win.show()
    sys.exit(app.exec_())