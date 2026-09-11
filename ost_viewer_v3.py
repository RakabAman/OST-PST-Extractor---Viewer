#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
OST Mail Viewer – Read and browse OST exports (IPM_SUBTREE)
Copyright (c) 2025
"""

import sys
import os
import re
import sqlite3
import datetime
import subprocess

from PyQt5.QtCore import (
    Qt, QThread, pyqtSignal, QUrl, QSettings, QFileInfo,
)
from PyQt5.QtGui import (
    QFont, QDesktopServices,
    QStandardItem, QStandardItemModel,
)
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QDialog,
    QVBoxLayout, QHBoxLayout, QGridLayout,
    QLineEdit, QPushButton, QCheckBox, QComboBox, QLabel,
    QGroupBox, QSplitter, QTreeView, QTreeWidget, QTreeWidgetItem,
    QTableView, QMenu, QAction, QFileDialog, QMessageBox,
    QDockWidget, QToolButton, QTextBrowser, QFileIconProvider,
)

# ----------------------------------------------------------------------
# Global font scaling factor (1.0 = normal)
SCALE_FACTOR = 1.0

# ----------------------------------------------------------------------
# Utility: find pffexport / pffinfo executables
def find_pff_tool(tool_name):
    """Search for pffexport/pffinfo in the app folder and PATH."""
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
        self.conn.execute('CREATE INDEX IF NOT EXISTS idx_messages_folder ON messages(folder_id)')
        self.conn.execute('CREATE INDEX IF NOT EXISTS idx_messages_subject ON messages(subject)')
        self.conn.execute('CREATE INDEX IF NOT EXISTS idx_messages_sender ON messages(sender_email)')
        self.conn.execute('CREATE INDEX IF NOT EXISTS idx_messages_date ON messages(date_utc)')
        self.conn.execute('CREATE INDEX IF NOT EXISTS idx_attachments_message ON attachments(message_id)')
        self.conn.commit()

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
        return cur.lastrowid

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
        cur.execute('SELECT * FROM attachments WHERE message_id = ? ORDER BY filename', (message_id,))
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

    def _build_ui(self):
        layout = QVBoxLayout(self)

        # ---- Tool paths ----
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

        # ---- Input / Output ----
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

        # ---- pffexport options ----
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

        # ---- pffinfo options ----
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

        # ---- Buttons ----
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

    def _browse_export(self):
        path, _ = QFileDialog.getOpenFileName(
            self, 'Select pffexport executable', '',
            'Executables (*.exe);;All files (*)'
        )
        if path:
            self.export_path_edit.setText(path)

    def _browse_info(self):
        path, _ = QFileDialog.getOpenFileName(
            self, 'Select pffinfo executable', '',
            'Executables (*.exe);;All files (*)'
        )
        if path:
            self.info_path_edit.setText(path)

    def _browse_input(self):
        path, _ = QFileDialog.getOpenFileName(
            self, 'Select OST/PST file', '',
            'Outlook files (*.ost *.pst);;All files (*)'
        )
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
        export_path = self.export_path_edit.text().strip()

        if not input_file or not os.path.isfile(input_file):
            QMessageBox.warning(self, 'Error', 'Please select a valid input OST/PST file.')
            return
        if not output_dir or not os.path.isdir(output_dir):
            QMessageBox.warning(self, 'Error', 'Please select a valid output folder.')
            return
        if not export_path or not os.path.isfile(export_path):
            QMessageBox.warning(self, 'Error', 'Please select a valid pffexport executable.')
            return

        if self.chk_run_info.isChecked():
            info_path = self.info_path_edit.text().strip()
            if info_path and os.path.isfile(info_path):
                info_args = self._build_command(info_path, is_export=False)
                try:
                    result = subprocess.run(
                        info_args, capture_output=True, text=True, timeout=120
                    )
                    info_text = result.stdout + result.stderr
                    QMessageBox.information(
                        self, 'pffinfo Output',
                        f'Command: {" ".join(info_args)}\n\n{info_text[:3000]}'
                    )
                except Exception as e:
                    QMessageBox.warning(self, 'pffinfo Error', str(e))

        export_args = self._build_command(export_path, is_export=True)

        try:
            result = subprocess.run(
                export_args, capture_output=True, text=True,
                cwd=output_dir, timeout=600
            )
            if result.returncode == 0:
                QMessageBox.information(
                    self, 'Extraction Complete',
                    f'pffexport finished successfully.\n\n{result.stdout[:2000]}'
                )
                self.result_root = output_dir
                self.accept()
            else:
                QMessageBox.critical(
                    self, 'Extraction Failed',
                    f'pffexport returned code {result.returncode}.\n\n'
                    f'STDOUT:\n{result.stdout[:2000]}\n\n'
                    f'STDERR:\n{result.stderr[:2000]}'
                )
        except subprocess.TimeoutExpired:
            QMessageBox.critical(self, 'Timeout', 'pffexport timed out after 600 seconds.')
        except Exception as e:
            QMessageBox.critical(self, 'Error', str(e))

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
                    except Exception:
                        pass
                elif line.startswith('Size:'):
                    size_str = line.split(':', 1)[1].strip()
                    try:
                        data['size'] = int(size_str)
                    except Exception:
                        pass
                elif line.startswith('Flags:'):
                    if 'Has attachments' in line:
                        data['has_attachments'] = True
                elif line.startswith('Conversation topic:'):
                    data['conversation_topic'] = line.split(':', 1)[1].strip()
    except Exception:
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
    except Exception:
        pass
    return data

# ----------------------------------------------------------------------
# Scanner thread
class ScannerThread(QThread):
    progress = pyqtSignal(int, int)
    status = pyqtSignal(str)
    finished = pyqtSignal(bool)

    def __init__(self, root_path, db_path):
        super().__init__()
        self.root_path = root_path
        self.db_path = db_path
        self.db = None

    def run(self):
        try:
            self.db = MailDB(self.db_path)
            self.db.clear_all()
            self.status.emit('Scanning folders...')
            ipm_path = None
            for entry in os.listdir(self.root_path):
                if entry == 'Root - Mailbox':
                    mailbox_root = os.path.join(self.root_path, entry)
                    if os.path.isdir(mailbox_root):
                        sub = os.path.join(mailbox_root, 'IPM_SUBTREE')
                        if os.path.isdir(sub):
                            ipm_path = sub
                            break
            if not ipm_path:
                self.status.emit('IPM_SUBTREE not found!')
                self.finished.emit(False)
                return

            message_folders = []
            for dirpath, dirnames, filenames in os.walk(ipm_path):
                if 'OutlookHeaders.txt' in filenames:
                    message_folders.append(dirpath)

            self.status.emit('Building folder tree...')
            folder_id_map = {}
            root_id = self.db.insert_folder('IPM_SUBTREE', None, ipm_path)
            folder_id_map[ipm_path] = root_id

            parent_folders = set()
            for msg_path in message_folders:
                p = os.path.dirname(msg_path)
                while p != ipm_path and p not in folder_id_map:
                    parent_folders.add(p)
                    p = os.path.dirname(p)
            sorted_parents = sorted(parent_folders, key=lambda x: len(x.split(os.sep)))
            for folder_path in sorted_parents:
                parent_path = os.path.dirname(folder_path)
                parent_id = folder_id_map.get(parent_path)
                if parent_id is None:
                    continue
                folder_name = os.path.basename(folder_path)
                fid = self.db.insert_folder(folder_name, parent_id, folder_path)
                folder_id_map[folder_path] = fid

            total = len(message_folders)
            self.status.emit(f'Processing {total} messages...')
            for idx, msg_folder in enumerate(message_folders):
                self.progress.emit(idx+1, total)
                folder_id = folder_id_map.get(os.path.dirname(msg_folder))
                if folder_id is None:
                    continue
                outlook_path = os.path.join(msg_folder, 'OutlookHeaders.txt')
                outlook_data = parse_outlook_headers(outlook_path)
                internet_path = os.path.join(msg_folder, 'InternetHeaders.txt')
                internet_data = parse_internet_headers(internet_path) if os.path.exists(internet_path) else {}
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

            self.db.conn.commit()
            self.status.emit('Scan complete.')
            self.finished.emit(True)
        except Exception as e:
            self.status.emit(f'Error: {str(e)}')
            self.finished.emit(False)
        finally:
            if self.db:
                self.db.close()

# ----------------------------------------------------------------------
# Main Window
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.db = None
        self.current_folder_id = None
        self.current_message_id = None
        self.root_path = None
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

        extract_btn = QPushButton('Extract OST/PST...')
        extract_btn.clicked.connect(self.open_extract_dialog)
        top_bar_layout.addWidget(extract_btn)

        main_layout.addWidget(top_bar)

        # ---------------- Main horizontal splitter ----------------
        h_splitter = QSplitter(Qt.Horizontal)

        # ---------- LEFT: vertical splitter (folder tree | msg folder contents) ----------
        left_v_splitter = QSplitter(Qt.Vertical)

        # Folder tree (top-left)
        self.folder_tree = QTreeView()
        self.folder_tree.setHeaderHidden(True)
        self.folder_tree.clicked.connect(self.on_folder_selected)
        left_v_splitter.addWidget(self.folder_tree)

        # Message folder contents (bottom-left)
        self.msg_folder_view = QTreeWidget()
        self.msg_folder_view.setHeaderLabels(['Name', 'Size'])
        self.msg_folder_view.setIndentation(10)
        self.msg_folder_view.setContextMenuPolicy(Qt.CustomContextMenu)
        self.msg_folder_view.customContextMenuRequested.connect(self.msg_folder_context_menu)
        self.msg_folder_view.itemDoubleClicked.connect(self.open_msg_folder_item)
        left_v_splitter.addWidget(self.msg_folder_view)

        left_v_splitter.setSizes([400, 300])
        h_splitter.addWidget(left_v_splitter)

        # ---------- RIGHT: vertical splitter (messages | preview) ----------
        right_v_splitter = QSplitter(Qt.Vertical)

        self.message_table = QTableView()
        self.message_table.setSortingEnabled(True)
        self.message_table.horizontalHeader().setStretchLastSection(True)
        self.message_table.clicked.connect(self.on_message_selected)
        right_v_splitter.addWidget(self.message_table)

        preview_widget = QWidget()
        preview_layout = QVBoxLayout(preview_widget)
        preview_layout.setContentsMargins(0, 0, 0, 0)

        self.textview = QTextBrowser()
        self.textview.setOpenExternalLinks(True)
        self.textview.setOpenLinks(False)
        self.textview.setReadOnly(True)
        preview_layout.addWidget(self.textview, 1)

        self.attachments_widget = QWidget()
        att_layout = QHBoxLayout(self.attachments_widget)
        att_layout.setContentsMargins(5, 5, 5, 5)
        att_layout.setSpacing(10)
        self.attachments_widget.setVisible(False)
        preview_layout.addWidget(self.attachments_widget, 0)

        right_v_splitter.addWidget(preview_widget)
        right_v_splitter.setSizes([400, 500])

        h_splitter.addWidget(right_v_splitter)
        h_splitter.setSizes([250, 1150])
        main_layout.addWidget(h_splitter, 1)

        # ---------------- Menus ----------------
        menubar = self.menuBar()
        file_menu = menubar.addMenu('File')
        open_action = QAction('Open Folder', self)
        open_action.triggered.connect(self.select_root_folder)
        file_menu.addAction(open_action)

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
                self.root_path = dlg.result_root
                self.save_settings()
                self.load_db()

    # ------------------------------------------------------------------
    # Settings / DB loading
    def load_settings(self):
        settings = QSettings('OSTMailViewer', 'Settings')
        last_root = settings.value('last_root', '')
        if last_root and os.path.isdir(last_root):
            self.root_path = last_root
            self.load_db()
        else:
            self.select_root_folder()

    def save_settings(self):
        settings = QSettings('OSTMailViewer', 'Settings')
        settings.setValue('last_root', self.root_path)

    def select_root_folder(self):
        folder = QFileDialog.getExistingDirectory(self, 'Select root folder (containing Root - Mailbox)')
        if folder:
            self.root_path = folder
            self.save_settings()
            self.load_db()

    def load_db(self):
        db_path = os.path.join(self.root_path, 'mail_cache.db')
        if os.path.exists(db_path):
            self.db = MailDB(db_path)
            self.populate_tree()
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
            os.remove(db_path)
        self.scan_and_load()

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
                except Exception:
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
            self.message_model.appendRow([item_subject, item_sender, item_date, item_size, item_attach])
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

        html_path = row['message_html_path']
        if html_path and os.path.exists(html_path):
            with open(html_path, 'r', encoding='utf-8', errors='ignore') as f:
                html = f.read()
            msg_dir = os.path.dirname(html_path)
            self.textview.setSearchPaths([msg_dir])
            base_url = QUrl.fromLocalFile(msg_dir + os.sep)
            self.textview.setHtml(html, base_url)
        else:
            self.textview.setHtml('<p>No HTML content</p>')

        # Populate the bottom-left message folder view
        self.populate_message_folder(row)

        # Populate the attachments strip under the preview
        attachments = self.db.get_attachments_for_message(msg_id)
        att_widget = self.attachments_widget
        layout = att_widget.layout()
        while layout.count():
            child = layout.takeAt(0)
            if child.widget():
                child.widget().deleteLater()

        if attachments:
            att_widget.setVisible(True)
            for att in attachments:
                btn = QToolButton()
                icon = QFileIconProvider().icon(QFileInfo(att['filepath']))
                btn.setIcon(icon)
                btn.setToolTip(f"{att['filename']}\n{att['filepath']}")
                btn.setText(att['filename'])
                btn.setToolButtonStyle(Qt.ToolButtonTextUnderIcon)
                btn.setFixedSize(80, 80)
                btn.setProperty('att_id', att['id'])
                btn.clicked.connect(self.open_attachment_from_button)
                btn.setContextMenuPolicy(Qt.CustomContextMenu)
                btn.customContextMenuRequested.connect(
                    lambda pos, b=btn: self.attachment_context_menu(pos, b))
                layout.addWidget(btn)
        else:
            att_widget.setVisible(False)

    # ------------------------------------------------------------------
    # Bottom-left message folder contents
    def populate_message_folder(self, msg_row):
        self.msg_folder_view.clear()
        outlook_path = msg_row['outlook_headers_path']
        if not outlook_path or not os.path.exists(outlook_path):
            return
        msg_dir = os.path.dirname(outlook_path)

        outlook_item = QTreeWidgetItem(self.msg_folder_view)
        outlook_item.setText(0, 'Outlook Data')
        outlook_item.setExpanded(True)
        attachments_item = QTreeWidgetItem(self.msg_folder_view)
        attachments_item.setText(0, 'Attachments')
        attachments_item.setExpanded(True)

        # Files directly inside the message folder
        try:
            entries = os.listdir(msg_dir)
        except OSError:
            entries = []
        for fname in entries:
            full_path = os.path.join(msg_dir, fname)
            if os.path.isfile(full_path):
                if fname.lower() in ('internetheaders.txt', 'outlookheaders.txt',
                                     'conversationindex.txt', 'recipients.txt',
                                     'message.html'):
                    item = QTreeWidgetItem(outlook_item)
                    item.setText(0, fname)
                    item.setText(1, f'{os.path.getsize(full_path):,}')
                    item.setData(0, Qt.UserRole, full_path)

        # Attachments subfolder
        att_dir = os.path.join(msg_dir, 'Attachments')
        if os.path.isdir(att_dir):
            for fname in os.listdir(att_dir):
                full_path = os.path.join(att_dir, fname)
                if os.path.isfile(full_path):
                    item = QTreeWidgetItem(attachments_item)
                    item.setText(0, fname)
                    item.setText(1, f'{os.path.getsize(full_path):,}')
                    item.setData(0, Qt.UserRole, full_path)

        self.msg_folder_view.resizeColumnToContents(0)
        self.msg_folder_view.resizeColumnToContents(1)

    def open_msg_folder_item(self, item, column):
        filepath = item.data(0, Qt.UserRole)
        if filepath and os.path.exists(filepath):
            QDesktopServices.openUrl(QUrl.fromLocalFile(filepath))

    def msg_folder_context_menu(self, pos):
        item = self.msg_folder_view.itemAt(pos)
        if not item:
            return
        filepath = item.data(0, Qt.UserRole)
        if not filepath:
            return
        menu = QMenu()
        open_act = menu.addAction('Open')
        open_loc_act = menu.addAction('Open file location')
        copy_path_act = menu.addAction('Copy file path')
        action = menu.exec_(self.msg_folder_view.mapToGlobal(pos))
        if action == open_act:
            QDesktopServices.openUrl(QUrl.fromLocalFile(filepath))
        elif action == open_loc_act:
            if sys.platform == 'win32':
                subprocess.Popen(['explorer', '/select,', filepath])
            else:
                QDesktopServices.openUrl(QUrl.fromLocalFile(os.path.dirname(filepath)))
        elif action == copy_path_act:
            QApplication.clipboard().setText(filepath)

    # ------------------------------------------------------------------
    # Attachment actions
    def open_attachment_from_button(self):
        btn = self.sender()
        att_id = btn.property('att_id')
        self.open_attachment_by_id(att_id)

    def open_attachment_by_id(self, att_id):
        att = self.db.get_attachment_by_id(att_id)
        if att:
            path = att['filepath']
            if os.path.exists(path):
                QDesktopServices.openUrl(QUrl.fromLocalFile(path))
            else:
                QMessageBox.warning(self, 'Error', 'File not found.')

    def open_attachment_from_table(self, index):
        att_id = self.att_model.itemFromIndex(index.sibling(index.row(), 0)).data(Qt.UserRole)
        if att_id:
            self.open_attachment_by_id(att_id)

    def attachment_context_menu(self, pos, button):
        menu = QMenu()
        open_act = menu.addAction('Open')
        open_loc_act = menu.addAction('Open file location')
        copy_path_act = menu.addAction('Copy file path')
        action = menu.exec_(button.mapToGlobal(pos))
        att_id = button.property('att_id')
        if att_id is None:
            return
        att = self.db.get_attachment_by_id(att_id)
        if not att:
            return
        path = att['filepath']
        if action == open_act:
            self.open_attachment_by_id(att_id)
        elif action == open_loc_act:
            if os.path.exists(path):
                if sys.platform == 'win32':
                    subprocess.Popen(['explorer', '/select,', path])
                else:
                    QDesktopServices.openUrl(QUrl.fromLocalFile(os.path.dirname(path)))
        elif action == copy_path_act:
            QApplication.clipboard().setText(path)

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
            if os.path.exists(path):
                if sys.platform == 'win32':
                    subprocess.Popen(['explorer', '/select,', path])
                else:
                    QDesktopServices.openUrl(QUrl.fromLocalFile(os.path.dirname(path)))
        elif action == copy_path_act:
            QApplication.clipboard().setText(att['filepath'])

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
                except Exception:
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
        if hasattr(self, 'textview'):
            self.textview.setFont(font)

    def closeEvent(self, event):
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