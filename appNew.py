#!/usr/bin/env python3
"""
Spool Support Sync
Developer: Nivin Varkey
Version: V1.9
Contact: Please Contact IT
License: MIT (full text in About)

Changes in V1.9:
 - Full UI redesign matching Syncronizing PDF Files mockup (neon dark theme)
 - Dedicated header bar with logo + title
 - Sidebar action panel (Load Excel, Select Spools, Run Batch, View Logs, Settings)
 - Results table in Home tab (live batch output)
 - Session log view in Logs tab
 - BUG FIX: Manual attachment guard — empty merge_list returns failure, not success
 - BUG FIX: run_batch() validates manual folder before starting when selection mode is ON
 - All original logic preserved (Excel refresh, recursive search, patterns, single-instance)
"""

import os
import sys
import time
import socket
import json
import threading
import traceback
from datetime import datetime
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor
from force_updater import check_update_available, UpdateConfirmDialog
import re

try:
    from PySide6.QtWidgets import (
        QApplication, QWidget, QMainWindow, QLabel, QPushButton,
        QFileDialog, QVBoxLayout, QHBoxLayout, QListWidget, QMessageBox,
        QCheckBox, QLineEdit, QTextEdit, QTabWidget, QGroupBox, QSizePolicy,
        QInputDialog, QScrollArea, QComboBox, QFrame, QDialog,
        QDialogButtonBox, QListWidgetItem, QProgressBar, QTableWidget,
        QTableWidgetItem, QHeaderView, QAbstractItemView, QSplitter,
        QGraphicsOpacityEffect, QSpacerItem
    )
    from PySide6.QtGui import QIcon, QPixmap, QDropEvent, QColor, QMovie # <--- FIXED: QMovie imported from PySide6
    from PySide6.QtCore import QUrl, Qt, QTimer, Signal, QObject, Slot
    from PySide6.QtCore import QEvent, QPropertyAnimation, QEasingCurve
except Exception as e:
    print("PySide6 not installed or failed to import:", e)
    print("Install with: pip install PySide6")
    raise

import pandas as pd
try:
    from pypdf import PdfWriter as _PdfWriter, PdfReader as _PdfReader
    _USE_PYPDF = True
except ImportError:
    from PyPDF2 import PdfMerger as _PdfMerger  # type: ignore
    _USE_PYPDF = False

try:
    import win32com.client as win32
    WIN32_AVAILABLE = True
except Exception:
    WIN32_AVAILABLE = False

try:
    from openpyxl import load_workbook
    OPENPYXL_AVAILABLE = True
except Exception:
    OPENPYXL_AVAILABLE = False

# ═══════════════════════════════════════════════════════════════
#  CONFIG
# ═══════════════════════════════════════════════════════════════
APP_NAME             = "Spool Support Sync"
DEVELOPER            = "Nivin Varkey"
VERSION              = "V1.9"
EMAIL                = "Please Contact IT"
SINGLE_INSTANCE_PORT = 50555

EXCEL_SHEET           = "STO"
SPOOL_COLUMN          = "Spool Num"
SUPPORT_COLUMN        = "EpicTagNumber"
OUTPUT_SUBFOLDER_NAME = "Support Attached Spools"
SUPPORTS_SUBPATH      = os.path.join("3.0 SUPPORTS", "3.2 SUPPORT DETAILS")
DEFAULT_REFRESH_MINUTES = 30
DEFAULT_MAX_WORKERS   = min(16, os.cpu_count() or 8)

LOG_FOLDER = os.path.join(
    os.getenv("APPDATA") or str(Path.home()), "SpoolSupportSync", "logs")
os.makedirs(LOG_FOLDER, exist_ok=True)

USER_ICON_FOLDER = os.path.join(os.getcwd(), "assets", "icons")
ICON_MAP = {
    "run": "run.png", "exit": "exit.png", "clear": "clear.png",
    "select_spool": "select_spool.png", "browse_excel": "browse_excel.png",
    "browse_support": "browse_support.png", "open_logs": "open_logs.png",
    "save_txt": "save_txt.png", "export_xlsx": "export_xlsx.png",
    "save_settings": "save_settings.png", "home": "home.png",
    "logs": "logs.png", "settings": "settings.png", "about": "about.png"
}

# ═══════════════════════════════════════════════════════════════
#  STYLESHEETS
# ═══════════════════════════════════════════════════════════════
NEON_QSS = """
QWidget {
    background: #03040d;
    color: #c8e8ff;
    font-family: "Segoe UI", Arial;
    font-size: 10pt;
}
QMainWindow { background: #03040d; }

QTabWidget::pane {
    border: 1px solid #0d2a4a;
    border-radius: 6px;
    background: #04070f;
}
QTabBar::tab {
    background: #07111f;
    color: #6090b8;
    padding: 9px 22px;
    border: 1px solid #0d2040;
    border-bottom: none;
    border-top-left-radius: 6px;
    border-top-right-radius: 6px;
    margin-right: 2px;
    font-weight: 600;
}
QTabBar::tab:selected {
    background: #04070f;
    color: #2eb8ff;
    border: 1px solid #1a4a80;
    border-bottom: none;
}
QTabBar::tab:hover:!selected { color: #90c8e8; background: #091828; }

QGroupBox {
    border: 1px solid #0d2a4a;
    border-radius: 7px;
    margin-top: 14px;
    padding: 10px 8px 8px 8px;
    background: #050a18;
    font-weight: 700;
    color: #2eb8ff;
    font-size: 9pt;
}
QGroupBox::title {
    subcontrol-origin: margin;
    left: 12px;
    padding: 0 6px;
    color: #2eb8ff;
    font-size: 9pt;
}

QPushButton {
    background: #06101e;
    border: 1px solid #1a3a60;
    border-radius: 6px;
    padding: 8px 14px;
    color: #80b8d8;
    font-weight: 600;
    font-size: 10pt;
}
QPushButton:hover {
    background: #0a1e38;
    border: 1px solid #2eb8ff;
    color: #ffffff;
}
QPushButton:pressed { background: #081428; }
QPushButton:disabled { color: #2a4a60; border-color: #0d1e30; }

QPushButton#btnRunBatch {
    background: #060f28;
    border: 2px solid #1a5ab0;
    color: #40a8e8;
    font-weight: 700;
    font-size: 11pt;
    letter-spacing: 1px;
    padding: 10px 20px;
}
QPushButton#btnRunBatch:hover {
    background: #0a1e48;
    border-color: #2eb8ff;
    color: #ffffff;
}
QPushButton#btnRunBatch:disabled { color: #2a5a80; border-color: #0d2a50; }

QPushButton#btnExit {
    background: #140608;
    border: 2px solid #801818;
    color: #e04040;
    font-weight: 700;
    font-size: 11pt;
    letter-spacing: 1px;
    padding: 10px 20px;
}
QPushButton#btnExit:hover { background: #200808; border-color: #ff3333; color: #ffffff; }

QPushButton#btnStop {
    background: #1a0a00;
    border: 2px solid #a04000;
    color: #ff8030;
    font-weight: 700;
    font-size: 11pt;
    letter-spacing: 1px;
    padding: 10px 20px;
}
QPushButton#btnStop:hover { background: #2a1000; border-color: #ff6010; color: #ffffff; }
QPushButton#btnStop:disabled { color: #3a2010; border-color: #2a1800; }

QPushButton#btnSelectSpools {
    background: #06101e;
    border: 2px solid #1a4a80;
    color: #40a0e0;
    font-weight: 700;
    padding: 8px 16px;
}
QPushButton#btnSelectSpools:hover { border-color: #2eb8ff; color: #fff; background: #0a1e38; }

QPushButton#btnClear {
    background: #12060a;
    border: 1px solid #4a1a2a;
    color: #b06080;
    font-weight: 600;
    padding: 8px 16px;
}
QPushButton#btnClear:hover { border-color: #ff4060; color: #ffaac0; background: #1c080e; }

QPushButton#sideBtn {
    background: #040c1a;
    border: 2px solid #0d2a50;
    border-radius: 7px;
    color: #3a7ab0;
    font-weight: 700;
    font-size: 10pt;
    letter-spacing: 1px;
    padding: 12px 10px;
    text-align: left;
}
QPushButton#sideBtn:hover { background: #071428; border-color: #2eb8ff; color: #90d8ff; }

QPushButton#sideBtnRed {
    background: #100406;
    border: 2px solid #6a1010;
    border-radius: 7px;
    color: #d04040;
    font-weight: 700;
    font-size: 10pt;
    letter-spacing: 1px;
    padding: 12px 10px;
    text-align: left;
}
QPushButton#sideBtnRed:hover { background: #200608; border-color: #ff3333; color: #ff8888; }

QPushButton#sideBtnTeal {
    background: #03100c;
    border: 2px solid #0a5040;
    border-radius: 7px;
    color: #20a080;
    font-weight: 700;
    font-size: 10pt;
    letter-spacing: 1px;
    padding: 12px 10px;
    text-align: left;
}
QPushButton#sideBtnTeal:hover { background: #061a14; border-color: #20d0a0; color: #60ffd8; }

QLineEdit {
    background: #04091a;
    border: 1px solid #0d2a4a;
    border-radius: 5px;
    padding: 6px 10px;
    color: #80b8d8;
    selection-background-color: #1a4a80;
}
QLineEdit:focus { border-color: #2060a0; color: #c8e8ff; }

QTextEdit {
    background: #04091a;
    border: 1px solid #0d2a4a;
    border-radius: 5px;
    padding: 6px;
    color: #80b8d8;
}

QListWidget {
    background: #04070f;
    border: 1px solid #0d2040;
    border-radius: 7px;
    padding: 4px;
    color: #90c0e0;
    outline: none;
}
QListWidget::item {
    border-bottom: 1px solid #08172a;
    padding: 5px 8px;
    border-radius: 4px;
}
QListWidget::item:selected { background: #0a1e38; color: #2eb8ff; }
QListWidget::item:hover:!selected { background: #070f1e; }

QTableWidget {
    background: #04070f;
    border: 1px solid #0d2040;
    border-radius: 7px;
    color: #90c0e0;
    gridline-color: #08172a;
    outline: none;
}
QTableWidget::item { padding: 5px 8px; border: none; }
QTableWidget::item:selected { background: #0a1e38; color: #2eb8ff; }
QTableWidget::item:hover:!selected { background: #070f1e; }
QHeaderView::section {
    background: #060e1d;
    color: #4a7a9a;
    border: none;
    border-bottom: 1px solid #0d2040;
    padding: 6px 8px;
    font-weight: 700;
    font-size: 9pt;
    letter-spacing: 0.5px;
}

QProgressBar {
    background: #04070f;
    border: 1px solid #0d2040;
    border-radius: 5px;
    text-align: center;
    color: #2eb8ff;
    font-weight: 700;
    font-size: 9pt;
    height: 22px;
}
QProgressBar::chunk {
    background: qlineargradient(x1:0,y1:0,x2:1,y2:0,stop:0 #0a3a80,stop:1 #2090d0);
    border-radius: 4px;
}

QComboBox {
    background: #04091a;
    border: 1px solid #0d2a4a;
    border-radius: 5px;
    padding: 6px 10px;
    color: #80b8d8;
    min-width: 100px;
}
QComboBox::drop-down { border: none; width: 20px; }
QComboBox QAbstractItemView {
    background: #060e1d;
    border: 1px solid #1a3a60;
    color: #90c0e0;
    selection-background-color: #0a1e38;
}

QScrollBar:vertical { background: #04070f; width: 7px; border-radius: 3px; }
QScrollBar::handle:vertical { background: #1a3a60; border-radius: 3px; min-height: 20px; }
QScrollBar::handle:vertical:hover { background: #2060a0; }
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0; }

QCheckBox { color: #80b8d8; spacing: 8px; }
QCheckBox::indicator {
    width: 16px; height: 16px;
    border: 1px solid #1a3a60;
    border-radius: 3px;
    background: #04091a;
}
QCheckBox::indicator:checked {
    background: #0a50a0;
    border-color: #2eb8ff;
    image: url(data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHZpZXdCb3g9IjAgMCAxNiAxNiI+PHBhdGggZD0iTTIgOEw2IDEyTDE0IDQiIHN0cm9rZT0iIzJlYjhmZiIgc3Ryb2tlLXdpZHRoPSIyLjUiIGZpbGw9Im5vbmUiIHN0cm9rZS1saW5lY2FwPSJyb3VuZCIgc3Ryb2tlLWxpbmVqb2luPSJyb3VuZCIvPjwvc3ZnPg==);
}

QSplitter::handle { background: #0d2040; width: 2px; height: 2px; }

QStatusBar {
    background: #04070f;
    border-top: 1px solid #0d2040;
    color: #4a8ab0;
    font-size: 9pt;
    padding: 2px 8px;
}

QMessageBox { background: #060e1d; color: #c8e8ff; }
QMessageBox QPushButton { min-width: 80px; }
QDialog { background: #060e1d; color: #c8e8ff; }
QDialogButtonBox QPushButton { min-width: 80px; }

QLabel { background: transparent; }
QLabel#titleLabel {
    font-size: 17pt;
    font-weight: 700;
    color: #2eb8ff;
    letter-spacing: 2px;
}
QLabel#sectionLabel { color: #2eb8ff; font-weight: 700; font-size: 10pt; letter-spacing: 1px; }
QLabel#dimLabel { color: #3a6080; font-size: 9pt; }
QLabel#countLabel {
    color: #4a8ab0;
    font-size: 9pt;
    font-weight: 600;
    padding: 2px 8px;
    border: 1px solid #0d2040;
    border-radius: 4px;
    background: #04070f;
}
QLabel#headerBg {
    background: #060e1d;
    border-bottom: 1px solid #0d2a4a;
}
"""

LIGHT_QSS = """
QWidget { background: #f0f4f8; color: #1a2a3a; font-family: "Segoe UI", Arial; font-size: 10pt; }
QMainWindow { background: #f0f4f8; }
QTabWidget::pane { border: 1px solid #c0d0e0; border-radius: 6px; background: #ffffff; }
QTabBar::tab { background: #e0eaf4; color: #4a6a8a; padding: 9px 22px; border: 1px solid #c0d0e0; border-bottom: none; border-top-left-radius: 6px; border-top-right-radius: 6px; margin-right: 2px; font-weight: 600; }
QTabBar::tab:selected { background: #ffffff; color: #1060a0; border-color: #1060a0; border-bottom: none; }
QGroupBox { border: 1px solid #c0d0e0; border-radius: 7px; margin-top: 14px; padding: 10px 8px 8px 8px; background: #ffffff; font-weight: 700; color: #1060a0; font-size: 9pt; }
QGroupBox::title { subcontrol-origin: margin; left: 12px; padding: 0 6px; color: #1060a0; }
QPushButton { background: #ffffff; border: 1px solid #b0c8e0; border-radius: 6px; padding: 8px 14px; color: #2060a0; font-weight: 600; font-size: 10pt; }
QPushButton:hover { background: #e8f4ff; border-color: #1060a0; color: #003a80; }
QPushButton#btnRunBatch { background: #e8f4ff; border: 2px solid #1060a0; color: #003a80; font-weight: 700; font-size: 11pt; padding: 10px 20px; }
QPushButton#btnRunBatch:hover { background: #d0e8ff; }
QPushButton#btnExit { background: #fff0f0; border: 2px solid #c04040; color: #a02020; font-weight: 700; font-size: 11pt; padding: 10px 20px; }
QPushButton#btnSelectSpools { background: #e8f4ff; border: 2px solid #1060a0; color: #003a80; font-weight: 700; padding: 8px 16px; }
QPushButton#btnClear { background: #fff0f0; border: 1px solid #e0a0a0; color: #a04040; font-weight: 600; padding: 8px 16px; }
QPushButton#sideBtn { background: #e8f4ff; border: 2px solid #1060a0; border-radius: 7px; color: #003a80; font-weight: 700; padding: 12px 10px; text-align: left; }
QPushButton#sideBtn:hover { background: #d0e8ff; border-color: #0040a0; }
QPushButton#sideBtnRed { background: #fff0f0; border: 2px solid #c04040; border-radius: 7px; color: #a02020; font-weight: 700; padding: 12px 10px; text-align: left; }
QPushButton#sideBtnTeal { background: #e8fff8; border: 2px solid #108060; border-radius: 7px; color: #006040; font-weight: 700; padding: 12px 10px; text-align: left; }
QLineEdit { background: #ffffff; border: 1px solid #b0c8e0; border-radius: 5px; padding: 6px 10px; color: #1a2a3a; }
QLineEdit:focus { border-color: #1060a0; }
QTextEdit { background: #ffffff; border: 1px solid #b0c8e0; border-radius: 5px; padding: 6px; color: #1a2a3a; }
QListWidget { background: #ffffff; border: 1px solid #b0c8e0; border-radius: 7px; padding: 4px; color: #1a2a3a; outline: none; }
QListWidget::item { border-bottom: 1px solid #e8f0f8; padding: 5px 8px; border-radius: 4px; }
QListWidget::item:selected { background: #d0e8ff; color: #003a80; }
QTableWidget { background: #ffffff; border: 1px solid #b0c8e0; border-radius: 7px; color: #1a2a3a; gridline-color: #e0ecf8; outline: none; }
QTableWidget::item { padding: 5px 8px; }
QTableWidget::item:selected { background: #d0e8ff; color: #003a80; }
QHeaderView::section { background: #e8f0f8; color: #4a6a8a; border: none; border-bottom: 1px solid #c0d0e0; padding: 6px 8px; font-weight: 700; font-size: 9pt; }
QProgressBar { background: #e8f0f8; border: 1px solid #b0c8e0; border-radius: 5px; text-align: center; color: #1060a0; font-weight: 700; height: 22px; }
QProgressBar::chunk { background: qlineargradient(x1:0,y1:0,x2:1,y2:0,stop:0 #1060c0,stop:1 #40a0e0); border-radius: 4px; }
QComboBox { background: #ffffff; border: 1px solid #b0c8e0; border-radius: 5px; padding: 6px 10px; color: #1a2a3a; }
QScrollBar:vertical { background: #e8f0f8; width: 7px; border-radius: 3px; }
QScrollBar::handle:vertical { background: #a0c0e0; border-radius: 3px; min-height: 20px; }
QCheckBox { color: #1a2a3a; spacing: 8px; }
QCheckBox::indicator { width: 16px; height: 16px; border: 1px solid #a0c0e0; border-radius: 3px; background: #ffffff; }
QCheckBox::indicator:checked { background: #1060a0; border-color: #003a80; image: url(data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHZpZXdCb3g9IjAgMCAxNiAxNiI+PHBhdGggZD0iTTIgOEw2IDEyTDE0IDQiIHN0cm9rZT0iI2ZmZmZmZiIgc3Ryb2tlLXdpZHRoPSIyLjUiIGZpbGw9Im5vbmUiIHN0cm9rZS1saW5lY2FwPSJyb3VuZCIgc3Ryb2tlLWxpbmVqb2luPSJyb3VuZCIvPjwvc3ZnPg==); }
QStatusBar { background: #e8f0f8; border-top: 1px solid #c0d0e0; color: #4a6a8a; font-size: 9pt; padding: 2px 8px; }
QLabel { background: transparent; }
QLabel#titleLabel { font-size: 17pt; font-weight: 700; color: #003a80; letter-spacing: 2px; }
QLabel#sectionLabel { color: #1060a0; font-weight: 700; font-size: 10pt; }
QLabel#dimLabel { color: #7090b0; font-size: 9pt; }
QLabel#countLabel { color: #4a6a8a; font-size: 9pt; font-weight: 600; padding: 2px 8px; border: 1px solid #c0d0e0; border-radius: 4px; }
QLabel#headerBg { background: #ddeeff; border-bottom: 1px solid #a0c0e0; }
"""

# ═══════════════════════════════════════════════════════════════
#  HELPERS
# ═══════════════════════════════════════════════════════════════
def resource_path(relative_path: str) -> str:
    try:
        base_path = sys._MEIPASS  # type: ignore
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


def load_icon(key: str) -> QIcon:
    fname = ICON_MAP.get(key)
    if not fname:
        return QIcon()
    for p in [
        resource_path(os.path.join("assets", "icons", fname)),
        os.path.join(USER_ICON_FOLDER, fname),
        resource_path(fname),
    ]:
        if os.path.exists(p):
            return QIcon(p)
    return QIcon()


DEFAULT_PATTERNS = [
    r"(?P<job>\d{5})-(?:\d{2}-)?(?P<spool>\d{6})(?:[-_]?model)?(?:R\d+|r\d+|REV\d+|_R\d+|-R\d+)?$",
    r"(?P<job>\d{5})-(?P<spool>\d{6})(?:_model|-model)?(?:R\d+|r\d+)?$",
    r"(?P<job>\d{5})[-_](?P<spool>\d{6})(?:[-_]?model)?(?:R\d+)?$",
    r"(?P<job>\d{5})(?:\D+)?(?P<spool>\d{6})(?:[-_]?model)?(?:R\d+)?$"
]


def build_patterns_from_settings(settings) -> list:
    patterns = list(DEFAULT_PATTERNS)
    try:
        customs = settings.get("custom_patterns", [])
        if isinstance(customs, list):
            patterns.extend(customs)
    except Exception:
        pass
    compiled = []
    for p in patterns:
        try:
            compiled.append(re.compile(p, re.I))
        except Exception:
            continue
    return compiled


def parse_spool_filename(filename: str, settings: dict):
    name = os.path.basename(filename)
    name = re.sub(r'\.pdf$', '', name, flags=re.I).strip()
    for rx in build_patterns_from_settings(settings):
        m = rx.search(name)
        if m:
            job   = m.groupdict().get("job")
            spool = m.groupdict().get("spool")
            if job and spool:
                return job, spool.zfill(6)
    return None


def excel_needs_refresh(path: str, minutes: int = DEFAULT_REFRESH_MINUTES) -> bool:
    try:
        return time.time() - os.path.getmtime(path) > minutes * 60
    except Exception:
        return True


def excel_refresh_com(path: str, timeout: int = 20):
    if not WIN32_AVAILABLE:
        raise RuntimeError("pywin32 not available for COM refresh")
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    wb = None
    try:
        wb = excel.Workbooks.Open(path, UpdateLinks=3)
        wb.RefreshAll()
        start = time.time()
        while time.time() - start < timeout:
            try:
                excel.CalculateUntilAsyncQueriesDone()
            except Exception:
                pass
            time.sleep(0.5)
        wb.Save()
    finally:
        try:
            if wb:
                wb.Close(SaveChanges=False)
        except Exception:
            pass
        try:
            excel.Quit()
        except Exception:
            pass


def excel_refresh_fallback(path: str):
    if not OPENPYXL_AVAILABLE:
        raise RuntimeError("openpyxl not available for fallback refresh")
    wb = load_workbook(path, data_only=False)
    wb.save(path)


def find_job_folder(job_number: str, base_paths: list) -> str:
    for base in base_paths:
        if not os.path.isdir(base):
            continue
        try:
            for f in os.listdir(base):
                if f.startswith(job_number):
                    return os.path.join(base, f)
        except Exception:
            pass
    return None


def build_support_index(folder, recursive=False):
    pdf_index = {}
    try:
        if recursive:
            for root, _, files in os.walk(folder):
                for f in files:
                    name, ext = os.path.splitext(f)
                    if ext.lower() == ".pdf":
                        pdf_index[name.lower()] = os.path.join(root, f)
        else:
            for f in os.listdir(folder):
                name, ext = os.path.splitext(f)
                if ext.lower() == ".pdf":
                    pdf_index[name.lower()] = os.path.join(folder, f)
    except Exception:
        pass
    return pdf_index


def find_pdf_for_tag(folder: str, tag: str) -> str:
    if not os.path.isdir(folder):
        return None
    candidate = os.path.join(folder, f"{tag}.pdf")
    if os.path.exists(candidate):
        return candidate
    try:
        for fname in os.listdir(folder):
            name, ext = os.path.splitext(fname)
            if name.lower() == tag.lower() and ext.lower() == ".pdf":
                return os.path.join(folder, fname)
    except Exception:
        pass
    return None


# ═══════════════════════════════════════════════════════════════
#  SINGLE INSTANCE
# ═══════════════════════════════════════════════════════════════
_singleton_socket = None
_listener_thread  = None


def _listener_accept_loop(sock, on_activate):
    try:
        while True:
            conn, _ = sock.accept()
            try:
                try:
                    conn.recv(1024)
                except Exception:
                    pass
                if on_activate:
                    QTimer.singleShot(0, on_activate)
            finally:
                try:
                    conn.close()
                except Exception:
                    pass
    except Exception:
        pass


def ensure_single_instance(on_activate_callable=None) -> bool:
    global _singleton_socket, _listener_thread
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    s.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
    try:
        s.bind(("127.0.0.1", SINGLE_INSTANCE_PORT))
        s.listen(1)
        _singleton_socket = s
        t = threading.Thread(target=_listener_accept_loop,
                             args=(s, on_activate_callable), daemon=True)
        t.start()
        _listener_thread = t
        return True
    except OSError:
        try:
            c = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            c.connect(("127.0.0.1", SINGLE_INSTANCE_PORT))
            try:
                c.send(b"ACTIVATE")
            except Exception:
                pass
            try:
                c.close()
            except Exception:
                pass
        except Exception:
            pass
        return False


def release_single_instance():
    global _singleton_socket
    try:
        if _singleton_socket:
            _singleton_socket.close()
            _singleton_socket = None
    except Exception:
        pass


# ═══════════════════════════════════════════════════════════════
#  MERGE JOB
# ═══════════════════════════════════════════════════════════════
class JobSignals(QObject):
    status_update = Signal(str, str, str, str) # spool_path, status_text, detail_text

class MergeJob:
    def __init__(self, spool_path, supports_folder, tags, output_path,
                 recursive_search=False, selected_support_paths=None, signals=None):
        self.spool_path             = spool_path
        self.supports_folder        = supports_folder
        self.tags                   = tags
        self.output_path            = output_path
        self.recursive              = recursive_search
        self.selected_support_paths = selected_support_paths
        self.signals = signals

    def _emit(self, status, detail="", tags_text=""):
        if self.signals:
            self.signals.status_update.emit(self.spool_path, status, detail, tags_text)

    def run(self):
        missing = []
        try:
            self._emit("Processing")

            if not os.path.exists(self.spool_path):
                self._emit("Failed", "spool missing")
                return (self.spool_path, False, "spool missing", [])

            if not self.spool_path.lower().endswith(".pdf"):
                self._emit("Failed", "not a PDF")
                return (self.spool_path, False, "not a PDF", [])

            merge_list = [self.spool_path]

            # ── Support Collection Phase ──────────────────────────
            self._emit("Support Imported", "Searching for supports…")

            if self.selected_support_paths is not None:
                # Enable-selection mode: user already chose exact paths
                for p in self.selected_support_paths:
                    if os.path.exists(p):
                        merge_list.append(p)
                    else:
                        missing.append(os.path.basename(p))
            else:
                # Auto mode: find by tag
                for tag in self.tags:
                    if not tag or str(tag).lower() in ("nan", "none", ""):
                        continue
                    found = None
                    if self.recursive:
                        for root, _, files in os.walk(self.supports_folder):
                            for f in files:
                                name, ext = os.path.splitext(f)
                                if (ext.lower() == ".pdf"
                                        and name.lower() == str(tag).lower()):
                                    found = os.path.join(root, f)
                                    break
                            if found:
                                break
                    else:
                        found = find_pdf_for_tag(self.supports_folder, str(tag))
                    if found:
                        merge_list.append(found)
                    else:
                        missing.append(str(tag))

            # ── Merge Phase ───────────────────────────────────────
            self._emit("Attaching", "Merging PDFs…")
            os.makedirs(os.path.dirname(self.output_path), exist_ok=True)
            _tmp_output = self.output_path + ".tmp"
            try:
                if _USE_PYPDF:
                    # pypdf: open files in binary mode with buffered I/O for speed
                    writer = _PdfWriter()
                    _handles = []
                    try:
                        for p in merge_list:
                            fh = open(p, "rb", buffering=65536)
                            _handles.append(fh)
                            reader = _PdfReader(fh, strict=False)
                            for page in reader.pages:
                                writer.add_page(page)
                        with open(_tmp_output, "wb", buffering=65536) as out_fh:
                            writer.write(out_fh)
                    finally:
                        for fh in _handles:
                            try: fh.close()
                            except Exception: pass
                else:
                    merger = _PdfMerger()
                    for p in merge_list:
                        merger.append(p, import_outline=False)
                    with open(_tmp_output, "wb") as out_fh:
                        merger.write(out_fh)
                    merger.close()
                # Atomic rename to avoid partial writes on network drives
                if os.path.exists(self.output_path):
                    os.remove(self.output_path)
                os.rename(_tmp_output, self.output_path)
            except Exception:
                if os.path.exists(_tmp_output):
                    try: os.remove(_tmp_output)
                    except Exception: pass
                raise

            # ── Success Phase ─────────────────────────────────────
            attached_names = [
                os.path.splitext(os.path.basename(p))[0]
                for p in merge_list[1:]
            ]
            tags_str = ", ".join(attached_names) if attached_names else "—"

            if missing:
                self._emit("Completed",
                           "Partial — Missing: " + ", ".join(missing),
                           tags_str)
                return (self.spool_path, True,
                        "Missing: " + ", ".join(missing), attached_names)

            self._emit("Completed", "OK", tags_str)
            return (self.spool_path, True, "OK", attached_names)

        except Exception as e:
            self._emit("Error", str(e))
            return (self.spool_path, False, f"Exception: {e}", [])


# ═══════════════════════════════════════════════════════════════
#  BATCH SUPPORT SELECTION DIALOG
# ═══════════════════════════════════════════════════════════════
class BatchSupportSelectionDialog(QDialog):
    """Select supports for ALL spools at once (optimized batch dialog)."""

    def __init__(self, spools_data, parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"Select Supports — {len(spools_data)} Spools")
        self.setMinimumSize(860, 560)
        self.spools_data   = spools_data
        self.spool_widgets = {}

        layout = QVBoxLayout(self)
        layout.setSpacing(10)
        layout.setContentsMargins(16, 16, 16, 16)

        hdr = QLabel(f"Select Supports for {len(spools_data)} Spools")
        hdr.setObjectName("sectionLabel")
        layout.addWidget(hdr)

        info = QLabel("Check the supports to merge with each spool. Unchecked items are skipped.")
        info.setObjectName("dimLabel")
        layout.addWidget(info)

        global_btn_row = QHBoxLayout()
        sel_all_btn = QPushButton("✔  Select All")
        sel_all_btn.clicked.connect(self.select_all_supports)
        global_btn_row.addWidget(sel_all_btn)
        desel_all_btn = QPushButton("✕  Deselect All")
        desel_all_btn.clicked.connect(self.deselect_all_supports)
        global_btn_row.addWidget(desel_all_btn)
        global_btn_row.addStretch()
        layout.addLayout(global_btn_row)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        scroll_widget = QWidget()
        scroll_layout = QVBoxLayout(scroll_widget)

        for spool_path, spool_name, available_supports in spools_data:
            grp = QGroupBox(f"Spool: {spool_name}  ({len(available_supports)} supports)")
            gl  = QVBoxLayout(grp)
            lw  = QListWidget()
            lw.setMaximumHeight(150)
            for support_path, support_tag in available_supports:
                fn  = os.path.splitext(os.path.basename(support_path))[0]
                txt = fn if fn.lower() == str(support_tag).lower() else f"{support_tag} — {fn}"
                item = QListWidgetItem(txt)
                item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
                item.setCheckState(Qt.Checked)
                item.setData(Qt.UserRole, support_path)
                lw.addItem(item)
            gl.addWidget(lw)

            row = QHBoxLayout()
            sb = QPushButton("Select All")
            sb.clicked.connect(lambda _, w=lw: self.select_all_in_list(w))
            db = QPushButton("Deselect All")
            db.clicked.connect(lambda _, w=lw: self.deselect_all_in_list(w))
            row.addWidget(sb)
            row.addWidget(db)
            row.addStretch()
            gl.addLayout(row)
            scroll_layout.addWidget(grp)
            self.spool_widgets[spool_path] = lw

        scroll_layout.addStretch()
        scroll.setWidget(scroll_widget)
        layout.addWidget(scroll)

        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.accepted.connect(self.accept)
        bb.rejected.connect(self.reject)
        layout.addWidget(bb)

    def select_all_supports(self):
        for lw in self.spool_widgets.values():
            self.select_all_in_list(lw)

    def deselect_all_supports(self):
        for lw in self.spool_widgets.values():
            self.deselect_all_in_list(lw)

    def select_all_in_list(self, lw):
        for i in range(lw.count()):
            lw.item(i).setCheckState(Qt.Checked)

    def deselect_all_in_list(self, lw):
        for i in range(lw.count()):
            lw.item(i).setCheckState(Qt.Unchecked)

    def get_selections(self):
        result = {}
        for sp, lw in self.spool_widgets.items():
            result[sp] = [lw.item(i).data(Qt.UserRole)
                          for i in range(lw.count())
                          if lw.item(i).checkState() == Qt.Checked]
        return result


# ═══════════════════════════════════════════════════════════════
#  SINGLE-SPOOL SUPPORT SELECTION DIALOG (legacy, kept for reference)
# ═══════════════════════════════════════════════════════════════
class SupportSelectionDialog(QDialog):
    def __init__(self, spool_name, available_supports, parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"Select Supports for {spool_name}")
        self.setMinimumSize(500, 400)
        self.selected_supports = []

        layout = QVBoxLayout(self)
        layout.setSpacing(10)
        layout.setContentsMargins(15, 15, 15, 15)

        layout.addWidget(QLabel(f"<b>Available Supports ({len(available_supports)} found)</b>"))
        info = QLabel("Check the supports you want to merge with the spool:")
        info.setObjectName("dimLabel")
        layout.addWidget(info)

        self.list_widget = QListWidget()
        layout.addWidget(self.list_widget)

        for support_path, support_tag in available_supports:
            fn  = os.path.splitext(os.path.basename(support_path))[0]
            txt = fn if fn.lower() == str(support_tag).lower() else f"{support_tag} - {fn}"
            item = QListWidgetItem(txt)
            item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
            item.setCheckState(Qt.Checked)
            item.setData(Qt.UserRole, support_path)
            self.list_widget.addItem(item)

        btn_row = QHBoxLayout()
        sa = QPushButton("Select All")
        sa.clicked.connect(self.select_all)
        da = QPushButton("Deselect All")
        da.clicked.connect(self.deselect_all)
        btn_row.addWidget(sa)
        btn_row.addWidget(da)
        layout.addLayout(btn_row)

        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.accepted.connect(self.accept)
        bb.rejected.connect(self.reject)
        layout.addWidget(bb)

    def select_all(self):
        for i in range(self.list_widget.count()):
            self.list_widget.item(i).setCheckState(Qt.Checked)

    def deselect_all(self):
        for i in range(self.list_widget.count()):
            self.list_widget.item(i).setCheckState(Qt.Unchecked)

    def get_selected_supports(self):
        return [self.list_widget.item(i).data(Qt.UserRole)
                for i in range(self.list_widget.count())
                if self.list_widget.item(i).checkState() == Qt.Checked]


# ═══════════════════════════════════════════════════════════════
#  DRAG-DROP SPOOL LIST
# ═══════════════════════════════════════════════════════════════
class SpoolListWidget(QListWidget):
    def __init__(self, add_callback=None):
        super().__init__()
        self.add_callback = add_callback
        self.setAcceptDrops(True)
        self.setObjectName("spoolList")

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            event.ignore()

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            event.ignore()

    def dropEvent(self, event: QDropEvent):
        if not event.mimeData().hasUrls():
            event.ignore()
            return
        added = 0
        bad   = 0
        for url in event.mimeData().urls():
            path = url.toLocalFile()
            if os.path.splitext(path)[1].lower() == ".pdf":
                if self.add_callback:
                    self.add_callback(path)
                else:
                    if path not in [self.item(i).text() for i in range(self.count())]:
                        self.addItem(path)
                added += 1
            else:
                bad += 1
        event.acceptProposedAction()
        if bad > 0:
            QMessageBox.warning(self, "Invalid files",
                                "Only PDF files are accepted. Non-PDF files were ignored.")
        if added > 0:
            self.scrollToBottom()


# ═══════════════════════════════════════════════════════════════
#  SPINNER LABEL  (animated braille spinner for "Processing…")
# ═══════════════════════════════════════════════════════════════
class SpinnerLabel(QLabel):
    _FRAMES = ["⠋", "⠙", "⠹", "⠸", "⠼", "⠴", "⠦", "⠧", "⠇", "⠏"]

    def __init__(self, parent=None):
        super().__init__(parent)
        self._prefix = "Processing"
        self._frame  = 0
        self._timer  = QTimer(self)
        self._timer.setInterval(80)
        self._timer.timeout.connect(self._tick)
        self.setStyleSheet(
            "color:#2eb8ff; font-weight:700; font-size:10pt;"
            "background:transparent;")
        self.setText("")

    def start(self, prefix="Attaching supports"):
        self._prefix = prefix
        self._frame  = 0
        self._timer.start()
        self.setVisible(True)

    def stop(self):
        self._timer.stop()
        self.setText("")
        self.setVisible(False)

    def set_file(self, filename: str):
        self._prefix = f"Attaching: {filename}" if filename else "Attaching supports"

    def _tick(self):
        self._frame = (self._frame + 1) % len(self._FRAMES)
        self.setText(f"  {self._FRAMES[self._frame]}  {self._prefix}")


# ═══════════════════════════════════════════════════════════════
#  HEADER WIDGET
# ═══════════════════════════════════════════════════════════════
class HeaderWidget(QWidget):
    """Top banner: logo + neon title bar."""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFixedHeight(68)
        self.setObjectName("headerBg")
        self.setStyleSheet(
            "#headerBg { background: #060e1d; border-bottom: 1px solid #0d2a4a; }")

        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 16, 0)
        layout.setSpacing(0)

        # Logo box
        logo_box = QWidget()
        logo_box.setFixedSize(68, 68)
        logo_box.setStyleSheet(
            "background: #04091a; border-right: 1px solid #0d2a4a;")
        ll = QHBoxLayout(logo_box)
        ll.setContentsMargins(10, 10, 10, 10)
        logo_label = QLabel()
        logo_file = resource_path("logo.png")
        if os.path.exists(logo_file):
            try:
                pix = QPixmap(logo_file).scaled(
                    46, 46, Qt.KeepAspectRatio, Qt.SmoothTransformation)
                logo_label.setPixmap(pix)
            except Exception:
                logo_label.setText("S")
        else:
            logo_label.setText("S")
            logo_label.setStyleSheet(
                "color:#2eb8ff; font-size:22pt; font-weight:700; background:transparent;")
        logo_label.setAlignment(Qt.AlignCenter)
        ll.addWidget(logo_label)
        layout.addWidget(logo_box)

        # Title
        title_lbl = QLabel(f" {APP_NAME} ")
        title_lbl.setObjectName("titleLabel")
        title_lbl.setStyleSheet(
            "QLabel#titleLabel { background: transparent; font-size:17pt; "
            "font-weight:700; color:#2eb8ff; letter-spacing:2px; }")
        title_lbl.setAlignment(Qt.AlignVCenter | Qt.AlignLeft)
        layout.addWidget(title_lbl, 1)

        # Version chip
        ver_lbl = QLabel(VERSION)
        ver_lbl.setStyleSheet(
            "color:#2a5a80; font-size:9pt; background:transparent;")
        ver_lbl.setAlignment(Qt.AlignVCenter | Qt.AlignRight)
        layout.addWidget(ver_lbl)

        # --- ADD THIS SPACER ---
        # Add a horizontal spacer to push the guide button to the right
        spacer = QSpacerItem(20, 20, QSizePolicy.Expanding, QSizePolicy.Minimum)
        layout.addSpacerItem(spacer)
        # -----------------------
        
        # Guide button
        guide_btn = QPushButton("❓  Guide")
        guide_btn.setObjectName("sideBtn")
        guide_btn.setFixedHeight(36)
        guide_btn.setStyleSheet(
            "QPushButton { background:#04091a; border:1px solid #1a4a80; "
            "border-radius:6px; color:#40a8e8; font-weight:700; "
            "font-size:10pt; padding:4px 14px; }"
            "QPushButton:hover { border-color:#2eb8ff; color:#fff; background:#0a1e38; }")
        guide_btn.clicked.connect(self._open_guide)
        layout.addWidget(guide_btn)

    def _open_guide(self):
        """Open the User Guide dialog."""
        dlg = UserGuideDialog(self)
        dlg.exec()

    def start_miniscreen(self):
        """Start the miniscreen GIF animation."""
        if self._miniscreen_movie:
            self._miniscreen_movie.start()

    def stop_miniscreen(self):
        """Stop and clear the miniscreen GIF, ready to replay next time."""
        if self._miniscreen_movie:
            self._miniscreen_movie.stop()
            self._miniscreen_movie.jumpToFrame(0)
        self.miniscreen_lbl.clear()
        if self._miniscreen_movie:
            self.miniscreen_lbl.setMovie(self._miniscreen_movie)


# ═══════════════════════════════════════════════════════════════
#  SIDEBAR
# ═══════════════════════════════════════════════════════════════
class SidebarWidget(QWidget):
    """Right-side action buttons."""
    sig_load_excel    = Signal()
    sig_select_spools = Signal()
    sig_run_batch     = Signal()
    sig_view_logs     = Signal()
    sig_settings      = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFixedWidth(195)
        self.setStyleSheet(
            "background: #04070f; border-left: 1px solid #0d2040;")

        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 18, 12, 18)
        layout.setSpacing(10)

        self.btn_load   = self._btn("📂  Load Excel List",    "sideBtnRed")
        self.btn_spools = self._btn("📄  Select Spool PDFs",  "sideBtn")
        self.btn_run    = self._btn("▶   Run Batch",          "sideBtn")
        self.btn_logs   = self._btn("📋  View Merged Logs",   "sideBtnTeal")
        self.btn_sett   = self._btn("⚙   Settings",           "sideBtn")

        self.btn_load.clicked.connect(self.sig_load_excel)
        self.btn_spools.clicked.connect(self.sig_select_spools)
        self.btn_run.clicked.connect(self.sig_run_batch)
        self.btn_logs.clicked.connect(self.sig_view_logs)
        self.btn_sett.clicked.connect(self.sig_settings)

        for b in [self.btn_load, self.btn_spools, self.btn_run,
                  self.btn_logs, self.btn_sett]:
            layout.addWidget(b)
        layout.addStretch()

    def _btn(self, text, obj):
        b = QPushButton(text)
        b.setObjectName(obj)
        b.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        return b

    def set_running(self, running: bool):
        self.btn_run.setEnabled(not running)
        self.btn_load.setEnabled(not running)
        self.btn_spools.setEnabled(not running)


# ═══════════════════════════════════════════════════════════════
#  USER GUIDE DIALOG
# ═══════════════════════════════════════════════════════════════
class UserGuideDialog(QDialog):
    """Interactive User Guide dialog with step-by-step instructions."""

    GUIDE_STEPS = [
        ("📂  Step 1 — Load Your Excel File",
         "Click  'Load Excel List'  in the sidebar or use 'Browse' in the Excel File panel.\n\n"
         "• The Excel file must contain a sheet named  STO.\n"
         "• The sheet must have columns:  'Spool Num'  and  'EpicTagNumber'.\n"
         "• If your file has formulas or linked data, enable auto-refresh (leave 'Skip Excel auto-refresh' unchecked).\n\n"
         "Tip: The last loaded Excel path is remembered between sessions."),

        ("📄  Step 2 — Select Spool PDFs",
         "Add your spool PDF files in either of two ways:\n\n"
         "• Click  'Select Spool PDFs'  to browse and multi-select PDF files.\n"
         "• Drag & drop PDF files directly onto the 'Selected Spools' list.\n\n"
         "Filenames must follow a recognised pattern, e.g.:\n"
         "  01555-00-000123.pdf\n"
         "  01555-000123R1.pdf\n"
         "  01555-000123_model.pdf\n\n"
         "If a filename is not recognised, a prompt will offer to open Settings where you can add a custom regex pattern."),

        ("📁  Step 3 — (Optional) Manual Support Folder",
         "If your support PDFs are stored in a specific folder:\n\n"
         "• Enter the folder path in 'Manual Support Folder', or click 'Browse'.\n"
         "• The folder will be searched recursively for support PDF files.\n\n"
         "NOTE: If 'Enable Support Selection' is turned ON in Settings, you MUST set a manual support folder before running the batch."),

        ("⚙   Step 4 — Configure Settings (Optional)",
         "Open the  Settings  tab to adjust:\n\n"
         "• Excel refresh interval — how often the Excel file is re-calculated (default: 30 min).\n"
         "• Max worker threads — how many spools are processed in parallel (default: CPU cores).\n"
         "• Enable Support Selection — shows a dialog to choose exactly which supports attach to each spool.\n"
         "• Custom filename patterns — add your own regex if filenames don't match built-in patterns.\n\n"
         "Always click  'Save Settings'  after making changes."),

        ("▶   Step 5 — Run the Batch",
         "Click  'Run Batch'  (or the sidebar button).\n\n"
         "During processing:\n"
         "• Each spool row in the Results table will animate in the Status column.\n"
         "• The miniscreen indicator in the header plays while the batch is running.\n"
         "• The progress bar shows overall completion percentage.\n\n"
         "When finished:\n"
         "• Status changes to  ✔ Completed  or  ✕ Failed  per row.\n"
         "• The miniscreen indicator stops.\n"
         "• A summary dialog shows success / failure / partial counts."),

        ("📋  Step 6 — Review Results & Logs",
         "After the batch completes:\n\n"
         "• The  Batch Results  table shows each spool, its status, attached tags, and any details.\n"
         "• Failed spools remain in the Spool list so you can re-run them.\n"
         "• Open the  Logs  tab to view the full session log.\n"
         "• Use  'Save current log as .txt'  or  'Export as .xlsx'  to keep records.\n"
         "• Logs are also automatically saved to your AppData folder.\n\n"
         "Merged output files are saved to a subfolder called  'Support Attached Spools'  next to each original spool PDF."),

        ("💡  Tips & Troubleshooting",
         "Common issues and solutions:\n\n"
         "• 'Unrecognised filename' — Go to Settings → add a regex with named groups (job) and (spool).\n"
         "• 'spool missing' — The PDF path no longer exists; re-select the file.\n"
         "• 'not a PDF' — Only .pdf files are supported as spool inputs.\n"
         "• Excel refresh fails — Make sure Excel is installed, or tick 'Skip Excel auto-refresh'.\n"
         "• Partial success / Missing tags — Some support files could not be found for the listed tags.\n\n"
         f"For further help contact:  {EMAIL}"),
    ]

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("User Guide — Spool Support Sync")
        self.setMinimumSize(680, 500)
        self._step = 0

        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 18, 20, 18)
        layout.setSpacing(12)

        # Header
        hdr = QLabel("📖  User Guide")
        hdr.setObjectName("sectionLabel")
        hdr.setStyleSheet("font-size:14pt; font-weight:700; color:#2eb8ff;")
        layout.addWidget(hdr)

        sub = QLabel(f"Step  1  of  {len(self.GUIDE_STEPS)}")
        sub.setObjectName("dimLabel")
        self._step_label = sub
        layout.addWidget(sub)

        # Step title
        self._title_lbl = QLabel()
        self._title_lbl.setObjectName("sectionLabel")
        self._title_lbl.setWordWrap(True)
        self._title_lbl.setStyleSheet("font-size:11pt; font-weight:700; color:#40c0ff;")
        layout.addWidget(self._title_lbl)

        # Content area
        self._content = QTextEdit()
        self._content.setReadOnly(True)
        self._content.setMinimumHeight(260)
        self._content.setStyleSheet(
            "font-size:10pt; color:#c8e8ff; background:#050a18; "
            "border:1px solid #0d2a4a; border-radius:6px; padding:10px;")
        layout.addWidget(self._content, 1)

        # Navigation row
        nav = QHBoxLayout()
        self._btn_prev = QPushButton("◀  Previous")
        self._btn_prev.clicked.connect(self._prev)
        self._btn_next = QPushButton("Next  ▶")
        self._btn_next.clicked.connect(self._next)
        self._btn_close = QPushButton("✕  Close")
        self._btn_close.clicked.connect(self.accept)

        nav.addWidget(self._btn_prev)
        nav.addStretch()
        nav.addWidget(self._btn_next)
        nav.addSpacing(20)
        nav.addWidget(self._btn_close)
        layout.addLayout(nav)

        self._show_step(0)

    def _show_step(self, idx):
        self._step = idx
        title, body = self.GUIDE_STEPS[idx]
        self._title_lbl.setText(title)
        self._content.setPlainText(body)
        self._step_label.setText(f"Step  {idx + 1}  of  {len(self.GUIDE_STEPS)}")
        self._btn_prev.setEnabled(idx > 0)
        self._btn_next.setEnabled(idx < len(self.GUIDE_STEPS) - 1)
        self._btn_next.setText("Next  ▶" if idx < len(self.GUIDE_STEPS) - 1 else "—  End  —")

    def _prev(self):
        if self._step > 0:
            self._show_step(self._step - 1)

    def _next(self):
        if self._step < len(self.GUIDE_STEPS) - 1:
            self._show_step(self._step + 1)


# ═══════════════════════════════════════════════════════════════
#  MAIN WINDOW
# ═══════════════════════════════════════════════════════════════
class MainWindow(QMainWindow):
    _show_batch_dialog_signal = Signal()

    def __init__(self):
        super().__init__()
        ico = resource_path("logo.ico")
        if os.path.exists(ico):
            try:
                self.setWindowIcon(QIcon(ico))
            except Exception:
                pass

        if not ensure_single_instance(on_activate_callable=self._bring_to_front):
            print("Another instance is already running — activating it and exiting.")
            sys.exit(0)

        self.setWindowTitle(f"{APP_NAME}  —  {DEVELOPER}")
        self.setMinimumSize(1100, 720)

        self.executor             = None
        self.stop_event           = threading.Event()
        self.selected_spools      = []
        _settings_dir = Path(os.getenv("APPDATA") or str(Path.home())) / "SpoolSupportSync"
        _settings_dir.mkdir(parents=True, exist_ok=True)
        self.settings_path        = _settings_dir / "settings.json"
        self.settings             = self._load_settings()
        self.threadpool           = ThreadPoolExecutor(
            max_workers=self.settings.get("max_workers", DEFAULT_MAX_WORKERS))
        self._pending_dialog_data = None
        self._dialog_result       = None
        self._last_run_results    = []
        self._row_anim_timers     = {}   # spool_path -> QTimer for status animation
        self._row_anim_frames     = {}   # spool_path -> frame index

        self._show_batch_dialog_signal.connect(self._show_batch_selection_dialog)

        self._build_ui()
        self._apply_theme(self.settings.get("theme", "neon"))

    def force_update_check(self):
        try:
            available, info = check_update_available()

            if available and info:
                dlg = UpdateConfirmDialog(info, self)
                dlg.exec()

        except Exception as ex:
            print("Update check failed:", ex)

    def manual_check_updates(self):
        self.force_update_check()

    # ── settings ──────────────────────────────────────────────
    def _load_settings(self):
        default = {
            "theme": "neon",
            "refresh_minutes": DEFAULT_REFRESH_MINUTES,
            "skip_excel_refresh": False,
            "manual_support_path": "",
            "max_workers": DEFAULT_MAX_WORKERS,
            "last_excel": "",
            "custom_patterns": [],
            "enable_support_selection": False,
        }
        try:
            if self.settings_path.exists():
                with open(self.settings_path, "r", encoding="utf-8") as f:
                    default.update(json.load(f))
        except Exception:
            pass
        return default

    def _save_settings(self):
        try:
            with open(self.settings_path, "w", encoding="utf-8") as f:
                json.dump(self.settings, f, indent=2)
        except Exception:
            pass

    def _apply_theme(self, theme):
        self.setStyleSheet(LIGHT_QSS if theme == "light" else NEON_QSS)

    # ── build UI ──────────────────────────────────────────────
    def _build_ui(self):
        root = QWidget()
        self.setCentralWidget(root)

        root_vbox = QVBoxLayout(root)
        root_vbox.setContentsMargins(0, 0, 0, 0)
        root_vbox.setSpacing(0)

        self.header_widget = HeaderWidget()
        root_vbox.addWidget(self.header_widget)

        body = QHBoxLayout()
        body.setContentsMargins(0, 0, 0, 0)
        body.setSpacing(0)

        self.tabs = QTabWidget()
        self.tabs.setDocumentMode(True)
        self.tabs.tabBar().setStyleSheet("QTabBar { background: transparent; }")

        self.home_tab     = QWidget()
        self.logs_tab     = QWidget()
        self.settings_tab = QWidget()
        self.about_tab    = QWidget()

        for tab in (self.home_tab, self.logs_tab, self.settings_tab, self.about_tab):
            tab.setAttribute(Qt.WA_TranslucentBackground, False)

        self.tabs.addTab(self.home_tab,     load_icon("home"),     "  Home")
        self.tabs.addTab(self.logs_tab,     load_icon("logs"),     "  Logs")
        self.tabs.addTab(self.settings_tab, load_icon("settings"), "  Settings")
        self.tabs.addTab(self.about_tab,    load_icon("about"),    "  About")

        body.addWidget(self.tabs, 1)
        root_vbox.addLayout(body, 1)

        self.statusBar().showMessage("Ready")

        self._build_home()
        self._build_logs()
        self._build_settings()
        self._build_about()

    # ── Home tab ──────────────────────────────────────────────
    def _build_home(self):
        layout = QVBoxLayout(self.home_tab)
        layout.setContentsMargins(12, 10, 12, 10)
        layout.setSpacing(10)

        # ── Welcome banner ──
        import getpass
        try:
            _username = getpass.getuser()
        except Exception:
            _username = os.getenv("USERNAME") or os.getenv("USER") or "User"
        welcome_lbl = QLabel(f"👋  Welcome, <b>{_username}</b>  —  ready to sync supports.")
        welcome_lbl.setStyleSheet(
            "color: #2eb8ff; font-size: 11pt; font-weight: 600; "
            "background: #04091a; border: 1px solid #0d2a4a; "
            "border-radius: 6px; padding: 6px 14px;")
        layout.addWidget(welcome_lbl)

        # ── Top row: Excel + Controls ──
        top_row = QHBoxLayout()
        top_row.setSpacing(10)

        excel_grp = QGroupBox("Excel File  (STO Sheet)")
        excel_gl  = QVBoxLayout(excel_grp)

        excel_hl = QHBoxLayout()
        self.excel_line = QLineEdit(self.settings.get("last_excel", ""))
        self.excel_line.setPlaceholderText("Path to Support_List Excel file…")
        excel_hl.addWidget(self.excel_line, 1)
        btn_bexcel = QPushButton("Browse")
        btn_bexcel.setIcon(load_icon("browse_excel"))
        btn_bexcel.clicked.connect(self.browse_excel)
        excel_hl.addWidget(btn_bexcel)
        excel_gl.addLayout(excel_hl)

        supp_hl = QHBoxLayout()
        supp_lbl = QLabel("Manual Support Folder:")
        supp_lbl.setObjectName("dimLabel")
        supp_hl.addWidget(supp_lbl)
        self.manual_support_line = QLineEdit(
            self.settings.get("manual_support_path", ""))
        self.manual_support_line.setPlaceholderText(
            "Optional — leave blank for auto-detect")
        supp_hl.addWidget(self.manual_support_line, 1)
        btn_bsupp = QPushButton("Browse")
        btn_bsupp.setIcon(load_icon("browse_support"))
        btn_bsupp.clicked.connect(self.browse_support)
        supp_hl.addWidget(btn_bsupp)
        excel_gl.addLayout(supp_hl)

        self.skip_refresh_chk = QCheckBox("Skip Excel auto-refresh")
        self.skip_refresh_chk.setChecked(
            self.settings.get("skip_excel_refresh", False))
        excel_gl.addWidget(self.skip_refresh_chk)

        top_row.addWidget(excel_grp, 2)

        # Controls group
        ctrl_grp = QGroupBox("Spool Controls")
        ctrl_gl  = QVBoxLayout(ctrl_grp)

        btn_row1 = QHBoxLayout()
        self.btn_select = QPushButton("Select Spool PDFs")
        self.btn_select.setObjectName("btnSelectSpools")
        self.btn_select.setIcon(load_icon("select_spool"))
        self.btn_select.clicked.connect(self.select_spools_dialog)
        self.btn_clear = QPushButton("Clear List")
        self.btn_clear.setObjectName("btnClear")
        self.btn_clear.setIcon(load_icon("clear"))
        self.btn_clear.clicked.connect(self.clear_spools)
        btn_row1.addWidget(self.btn_select, 2)
        btn_row1.addWidget(self.btn_clear, 1)
        ctrl_gl.addLayout(btn_row1)

        self.count_label = QLabel("Selected: 0 spools")
        self.count_label.setObjectName("countLabel")
        ctrl_gl.addWidget(self.count_label)

        # Animated progress bar (0 → 100, smooth easing)
        self.progress = QProgressBar()
        self.progress.setRange(0, 100)
        self.progress.setValue(0)
        self.progress.setFormat("Ready")
        self.progress.setFixedHeight(24)
        ctrl_gl.addWidget(self.progress)

        # Spinner shown while batch is active
        self.spinner_lbl = SpinnerLabel()
        self.spinner_lbl.setVisible(False)
        ctrl_gl.addWidget(self.spinner_lbl)

        run_row = QHBoxLayout()
        self.btn_run = QPushButton("Run Batch")
        self.btn_run.setObjectName("btnRunBatch")
        self.btn_run.setIcon(load_icon("run"))
        self.btn_run.clicked.connect(self.run_batch)
        self.btn_stop = QPushButton("⏹  Stop")
        self.btn_stop.setObjectName("btnStop")
        self.btn_stop.clicked.connect(self._stop_batch)
        self.btn_stop.setEnabled(False)
        self.btn_exit = QPushButton("Exit")
        self.btn_exit.setObjectName("btnExit")
        self.btn_exit.setIcon(load_icon("exit"))
        self.btn_exit.clicked.connect(self.close_event)
        run_row.addWidget(self.btn_run, 3)
        run_row.addWidget(self.btn_stop, 2)
        run_row.addWidget(self.btn_exit, 1)
        ctrl_gl.addLayout(run_row)

        top_row.addWidget(ctrl_grp, 1)
        layout.addLayout(top_row)

        # ── Bottom: spool list + results table ──
        splitter = QSplitter(Qt.Horizontal)

        spool_grp = QGroupBox("Selected Spools  (drag & drop PDFs here)")
        spool_gl  = QVBoxLayout(spool_grp)
        self.spool_list = SpoolListWidget(add_callback=self._add_spool_from_drag)
        self.spool_list.setWordWrap(True)
        self.spool_list.setSpacing(4)
        spool_gl.addWidget(self.spool_list)
        drop_hint = QLabel("Drag & drop spool PDFs or use 'Select Spool PDFs' above")
        drop_hint.setObjectName("dimLabel")
        drop_hint.setAlignment(Qt.AlignCenter)
        spool_gl.addWidget(drop_hint)
        splitter.addWidget(spool_grp)

        res_grp = QGroupBox("Batch Results")
        res_gl  = QVBoxLayout(res_grp)
        self.results_table = QTableWidget(0, 4)
        self.results_table.setHorizontalHeaderLabels(
            ["Spool File", "Status", "Attached Tags", "Details"])
        self.results_table.horizontalHeader().setSectionResizeMode(
            0, QHeaderView.ResizeToContents)
        self.results_table.horizontalHeader().setSectionResizeMode(
            1, QHeaderView.ResizeToContents)
        self.results_table.horizontalHeader().setSectionResizeMode(
            2, QHeaderView.Stretch)
        self.results_table.horizontalHeader().setSectionResizeMode(
            3, QHeaderView.ResizeToContents)
        self.results_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.results_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.results_table.verticalHeader().setVisible(False)
        self.results_table.setWordWrap(True)
        res_gl.addWidget(self.results_table)
        splitter.addWidget(res_grp)

        splitter.setSizes([420, 580])
        layout.addWidget(splitter, 1)

    # ── Logs tab ──────────────────────────────────────────────
    def _build_logs(self):
        layout = QVBoxLayout(self.logs_tab)
        layout.setContentsMargins(16, 14, 16, 14)
        layout.setSpacing(10)

        hdr = QLabel("Run Logs")
        hdr.setObjectName("sectionLabel")
        layout.addWidget(hdr)

        info = QLabel(f"Logs are saved automatically to:  {LOG_FOLDER}")
        info.setObjectName("dimLabel")
        info.setWordWrap(True)
        layout.addWidget(info)

        btn_open = QPushButton("Open Logs Folder")
        btn_open.setIcon(load_icon("open_logs"))
        btn_open.clicked.connect(lambda: os.startfile(LOG_FOLDER))
        layout.addWidget(btn_open)

        btn_save_txt = QPushButton("Save current log as .txt")
        btn_save_txt.setIcon(load_icon("save_txt"))
        btn_save_txt.clicked.connect(self.save_log_txt)
        layout.addWidget(btn_save_txt)

        btn_save_xlsx = QPushButton("Export current log as .xlsx")
        btn_save_xlsx.setIcon(load_icon("export_xlsx"))
        btn_save_xlsx.clicked.connect(self.save_log_xlsx)
        layout.addWidget(btn_save_xlsx)

        btn_del = QPushButton("🗑  Delete logs older than N days…")
        btn_del.clicked.connect(self._delete_old_logs_prompt)
        layout.addWidget(btn_del)

        layout.addWidget(QLabel("Current Session Log"))
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setPlaceholderText(
            "Batch results will appear here after running…")
        self.log_text.setStyleSheet(
            "QTextEdit {"
            "  background: #04070f;"
            "  color: #a8c8e8;"
            "  font-family: 'Consolas', 'Courier New', monospace;"
            "  font-size: 10pt;"
            "  border: 1px solid #0d2040;"
            "  border-radius: 6px;"
            "  padding: 8px;"
            "}")
        layout.addWidget(self.log_text, 1)

    def _delete_old_logs_prompt(self):
        days, ok = QInputDialog.getInt(
            self, "Delete old logs", "Delete logs older than (days):", 30, 1, 365, 1)
        if ok:
            cutoff  = time.time() - (days * 86400)
            removed = 0
            for fn in os.listdir(LOG_FOLDER):
                p = os.path.join(LOG_FOLDER, fn)
                try:
                    if os.path.isfile(p) and os.path.getmtime(p) < cutoff:
                        os.remove(p)
                        removed += 1
                except Exception:
                    pass
            QMessageBox.information(self, "Deleted", f"Removed {removed} log file(s).")
            self.statusBar().showMessage(f"Deleted {removed} old logs")

    # ── Settings tab ──────────────────────────────────────────
    def _build_settings(self):
        layout = QVBoxLayout(self.settings_tab)
        layout.setContentsMargins(16, 14, 16, 14)
        layout.setSpacing(10)

        hdr = QLabel("Settings")
        hdr.setObjectName("sectionLabel")
        layout.addWidget(hdr)

        theme_row = QHBoxLayout()
        theme_row.addWidget(QLabel("Theme:"))
        self.theme_select = QComboBox()
        self.theme_select.addItems(["neon", "light"])
        self.theme_select.setCurrentText(self.settings.get("theme", "neon"))
        self.theme_select.currentTextChanged.connect(self._on_theme_changed)
        theme_row.addWidget(self.theme_select)
        theme_row.addStretch()
        layout.addLayout(theme_row)

        self.enable_support_selection_chk = QCheckBox(
            "Enable Support Selection  (choose which supports to merge per batch)")
        self.enable_support_selection_chk.setChecked(
            self.settings.get("enable_support_selection", False))
        self.enable_support_selection_chk.setToolTip(
            "When enabled, a dialog appears before each batch so you can choose which supports to attach.")
        layout.addWidget(self.enable_support_selection_chk)

        layout.addWidget(QLabel("Excel refresh interval (minutes):"))
        self.refresh_line = QLineEdit(
            str(self.settings.get("refresh_minutes", DEFAULT_REFRESH_MINUTES)))
        layout.addWidget(self.refresh_line)

        layout.addWidget(QLabel("Max worker threads:"))
        self.max_workers_line = QLineEdit(
            str(self.settings.get("max_workers", DEFAULT_MAX_WORKERS)))
        layout.addWidget(self.max_workers_line)

        layout.addWidget(QLabel(
            "Custom filename patterns (regex).\n"
            "Must include named groups  (?P<job>…)  and  (?P<spool>…)\n"
            "Example:  (?P<job>\\d{5})-(?P<spool>\\d{6})"))
        self.pattern_line = QLineEdit()
        self.pattern_line.setPlaceholderText("Enter regex pattern here…")
        layout.addWidget(self.pattern_line)

        btn_add = QPushButton("Add Pattern")
        btn_add.clicked.connect(self._add_custom_pattern_from_ui)
        layout.addWidget(btn_add)

        btn_save = QPushButton("Save Settings")
        btn_save.setIcon(load_icon("save_settings"))
        btn_save.clicked.connect(self._save_settings_from_ui)
        layout.addWidget(btn_save)
        layout.addStretch()

    def _add_custom_pattern_from_ui(self):
        pat = self.pattern_line.text().strip()
        if not pat:
            QMessageBox.information(self, "Empty", "Enter a regex pattern first.")
            return
        try:
            rx = re.compile(pat, re.I)
            if "job" not in rx.groupindex or "spool" not in rx.groupindex:
                QMessageBox.warning(self, "Invalid pattern",
                    "Pattern must contain named groups 'job' and 'spool'.")
                return
        except Exception as e:
            QMessageBox.warning(self, "Invalid regex", f"Compilation failed: {e}")
            return
        customs = self.settings.get("custom_patterns", [])
        if pat not in customs:
            customs.append(pat)
            self.settings["custom_patterns"] = customs
            self._save_settings()
            QMessageBox.information(self, "Saved", "Pattern added.")
            self.pattern_line.clear()
            self.statusBar().showMessage("Custom pattern added.")
        else:
            QMessageBox.information(self, "Exists", "Pattern already present.")

    def _save_settings_from_ui(self):
        try:
            self.settings["refresh_minutes"] = int(self.refresh_line.text())
        except Exception:
            self.settings["refresh_minutes"] = DEFAULT_REFRESH_MINUTES
        try:
            self.settings["max_workers"] = int(self.max_workers_line.text())
        except Exception:
            self.settings["max_workers"] = DEFAULT_MAX_WORKERS
        self.settings["enable_support_selection"] = \
            self.enable_support_selection_chk.isChecked()
        self._save_settings()
        QMessageBox.information(self, "Saved", "Settings saved.")
        self.statusBar().showMessage("Settings saved")

    def _on_theme_changed(self, text):
        self.settings["theme"] = text
        self._apply_theme(text)
        self._save_settings()
        self.statusBar().showMessage(f"Theme: {text}")

    # ── About tab ─────────────────────────────────────────────
    def _build_about(self):
        layout = QVBoxLayout(self.about_tab)
        layout.setContentsMargins(16, 14, 16, 14)
        layout.setSpacing(12)

        top = QHBoxLayout()
        logo_lbl = QLabel()
        lf = resource_path("logo.png")
        if os.path.exists(lf):
            try:
                pix = QPixmap(lf).scaled(
                    120, 120, Qt.KeepAspectRatio, Qt.SmoothTransformation)
                logo_lbl.setPixmap(pix)
            except Exception:
                pass
        logo_lbl.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        top.addWidget(logo_lbl)

        info_vbox = QVBoxLayout()
        t = QLabel(APP_NAME)
        t.setObjectName("titleLabel")
        info_vbox.addWidget(t)
        for txt in [f"Version: {VERSION}", f"Developer: {DEVELOPER}",
                    f"Contact: {EMAIL}"]:
            lbl = QLabel(txt)
            lbl.setObjectName("dimLabel")
            info_vbox.addWidget(lbl)
        dev_msg = QTextEdit()
        dev_msg.setReadOnly(True)
        dev_msg.setFixedHeight(68)
        dev_msg.setPlainText(
            "Thank you for using Spool Support Sync.\nDeveloped for EPIC PIPING.")
        info_vbox.addWidget(dev_msg)
        top.addLayout(info_vbox)
        top.addStretch()
        layout.addLayout(top)

        layout.addWidget(QLabel("Copyright and License Notice"))
        instr = QTextEdit()
        instr.setReadOnly(True)
        instr.setMinimumHeight(260)
        instr.setPlainText(
            "Copyright © 2025 Nivin Varkey\n\n"
            "Spool Support Sync™ was developed and provided to Epic Piping, LLC\n"
            "free of charge by Nivin Varkey.\n"
            "Epic Piping, LLC is granted a perpetual, royalty-free, non-exclusive right\n"
            "to use, modify, maintain, enhance, and distribute this software for its\n"
            "internal and business operations without any development or licensing fees.\n\n"
            "Any modified or derivative versions of this software shall retain:\n\n"
            "1. The original software name: Spool Support Sync(TM)\n"
            "2. The developer attribution: Developed by Nivin Varkey\n\n"
            "unless written permission is obtained from the original developer.\n\n"
            "This software is provided AS IS, without warranty of any kind, express\n"
            "or implied, including but not limited to the warranties of merchantability,\n"
            "fitness for a particular purpose, and noninfringement. In no event shall\n"
            "the developer be liable for any claim, damages, or other liability arising\n"
            "from the use of this software.\n"
            f"For more help contact: {EMAIL}\n"
        )
        layout.addWidget(instr, 1)
        self.btn_about_update = QPushButton("Check For Updates")
        self.btn_about_update.clicked.connect(self.manual_check_updates)
        layout.addWidget(self.btn_about_update)

        # Check GitHub silently and change About button text if an update exists.
        QTimer.singleShot(800, self.refresh_about_update_button)

    def refresh_about_update_button(self):
        try:
            available, info = check_update_available()

            if available and info:
                latest = info.get("version", "")
                self.btn_about_update.setText(f"Update Available - Install V{latest}")
                self.btn_about_update.setStyleSheet(
                    "background:#0a3a80; color:white; font-weight:bold; padding:8px;"
                )
            else:
                self.btn_about_update.setText("Check For Updates")
                self.btn_about_update.setStyleSheet("")

        except Exception:
            self.btn_about_update.setText("Check For Updates")
            self.btn_about_update.setStyleSheet("")

    # ── UI helpers ────────────────────────────────────────────
    def _add_spool_from_drag(self, path):
        parsed = parse_spool_filename(path, self.settings)
        if parsed:
            if path not in self.selected_spools:
                self.selected_spools.append(path)
                item = QListWidgetItem(
                    f"{os.path.basename(path)}\n{os.path.dirname(path)}")
                item.setData(Qt.UserRole, path)
                self.spool_list.addItem(item)
                self._update_selected_count()
                self.statusBar().showMessage(
                    f"Added: {os.path.basename(path)}")
        else:
            self.statusBar().showMessage(
                f"Unrecognised filename: {os.path.basename(path)}")
            reply = QMessageBox.question(
                self, "Unrecognised filename",
                f"'{os.path.basename(path)}' didn't match known patterns.\n"
                "Open Settings to add a pattern for this filename?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
            if reply == QMessageBox.StandardButton.Yes:
                self.tabs.setCurrentWidget(self.settings_tab)

    def select_spools_dialog(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, "Select Spool PDFs", "", "PDF Files (*.pdf)")
        added = 0
        for f in files:
            if parse_spool_filename(f, self.settings):
                if f not in self.selected_spools:
                    self.selected_spools.append(f)
                    item = QListWidgetItem(
                        f"{os.path.basename(f)}\n{os.path.dirname(f)}")
                    item.setData(Qt.UserRole, f)
                    self.spool_list.addItem(item)
                    added += 1
            else:
                self.statusBar().showMessage(
                    f"Unrecognised: {os.path.basename(f)}")
        if added:
            self._update_selected_count()
            self.statusBar().showMessage(f"Added {added} spool(s)")

    def clear_spools(self):
        self.selected_spools = []
        self.spool_list.clear()
        self.results_table.setRowCount(0)
        # Stop any lingering row animations
        for sp in list(self._row_anim_timers.keys()):
            self._stop_row_animation(sp)
        self.progress.setValue(0)
        self.progress.setFormat("")
        self._update_selected_count()
        self.statusBar().showMessage("Spool list cleared")

    def _update_selected_count(self):
        n = len(self.selected_spools)
        self.count_label.setText(f"Selected: {n} spool{'s' if n != 1 else ''}")

    def browse_excel(self):
        p, _ = QFileDialog.getOpenFileName(
            self, "Select Support_List Excel", "",
            "Excel Files (*.xlsx *.xls *.xlsb)")
        if p:
            self.excel_line.setText(p)
            self.settings["last_excel"] = p
            self._save_settings()
            self.statusBar().showMessage(f"Excel: {p}")

    def browse_support(self):
        p = QFileDialog.getExistingDirectory(
            self, "Select manual support folder")
        if p:
            self.manual_support_line.setText(p)
            self.settings["manual_support_path"] = p
            self._save_settings()
            self.statusBar().showMessage(f"Support folder: {p}")

    # ── Log formatting ────────────────────────────────────────
    def _format_log(self, results):
        """Return a fully formatted log string matching the standard layout."""
        W  = 100          # total line width
        DIVIDER  = "=" * W
        SEPARATOR = "-" * W

        ok      = sum(1 for r in results if r[1] and "Missing" not in str(r[2]))
        fail    = sum(1 for r in results if not r[1])
        partial = sum(1 for r in results if r[1] and "Missing" in str(r[2]))

        batch_ts  = getattr(self, "_batch_start_ts", datetime.now())
        finish_ts = datetime.now()

        lines = []
        lines.append(DIVIDER)
        lines.append(f" Batch Run : {batch_ts.strftime('%Y-%m-%d %H:%M:%S')}")
        lines.append(f" Success   : {ok}")
        lines.append(f" Failed    : {fail}")
        lines.append(f" Partial   : {partial}")
        lines.append(DIVIDER)
        lines.append("")

        for r in results:
            spool_path = r[0]
            success    = r[1]
            detail     = str(r[2])
            tags       = r[3] if len(r) > 3 else []

            # Determine status label
            if success and "Missing" in detail:
                status_lbl = "PARTIAL"
            elif success:
                status_lbl = "SUCCESS"
            else:
                status_lbl = "FAILED"

            # Tags line
            if tags:
                tags_str = ", ".join(tags)
                attached = len(tags)
                selected = detail.split("Selected:")[-1].split("|")[0].strip() \
                    if "Selected:" in detail else str(attached)
                detail_str = f"Selected:{selected} | Attached:{attached}"
                if "Missing" in detail:
                    missing_part = detail.split("Missing:")[-1].strip()
                    detail_str += f" | Missing:{missing_part}"
            else:
                tags_str   = "None"
                detail_str = detail

            ts = finish_ts.strftime("%Y-%m-%d %H:%M:%S")
            lines.append(f"[{ts}] Status   : {status_lbl}")
            lines.append(f"          Spool No : {os.path.basename(spool_path)}")
            lines.append(f"          Tags     : {tags_str}")
            lines.append(f"          Details  : {detail_str}")
            lines.append(SEPARATOR)
            lines.append("")

        return "\n".join(lines)

    def _format_log_html(self, results):
        """Return colour-coded HTML log for display in the UI QTextEdit."""
        # ── Colour palette ──
        C_DIVIDER   = "#1a3a5c"
        C_HEADER_LB = "#2eb8ff"   # label
        C_HEADER_VL = "#ffffff"   # value
        C_TS        = "#5a8aaa"   # timestamp
        C_SUCCESS   = "#30c080"   # green
        C_FAILED    = "#e04040"   # red
        C_PARTIAL   = "#e0a030"   # amber
        C_LABEL     = "#6090b0"   # field names
        C_SPOOL     = "#d0e8ff"   # spool number
        C_TAGS      = "#60d8a8"   # tag list
        C_DETAIL    = "#a8c8e8"   # detail text
        C_MISSING   = "#e0a030"   # missing detail
        C_SEP       = "#0d2a40"   # row separator

        ok      = sum(1 for r in results if r[1] and "Missing" not in str(r[2]))
        fail    = sum(1 for r in results if not r[1])
        partial = sum(1 for r in results if r[1] and "Missing" in str(r[2]))
        batch_ts = getattr(self, "_batch_start_ts", datetime.now())

        def span(color, text):
            return f'<span style="color:{color}">{text}</span>'

        def div_line(color):
            return (f'<span style="color:{color}">'
                    f'{"=" * 96}</span><br>')

        def sep_line(color):
            return (f'<span style="color:{color}">'
                    f'{"-" * 96}</span><br>')

        html_parts = ['<pre style="margin:0;padding:0;line-height:1.5;">']

        # ── Header block ──
        html_parts.append(div_line(C_DIVIDER))
        html_parts.append(
            f' {span(C_HEADER_LB, "Batch Run")} : '
            f'{span(C_HEADER_VL, batch_ts.strftime("%Y-%m-%d %H:%M:%S"))}<br>')
        html_parts.append(
            f' {span(C_HEADER_LB, "Success  ")} : '
            f'{span(C_SUCCESS, str(ok))}<br>')
        html_parts.append(
            f' {span(C_HEADER_LB, "Failed   ")} : '
            f'{span(C_FAILED, str(fail)) if fail else span(C_DETAIL, "0")}<br>')
        html_parts.append(
            f' {span(C_HEADER_LB, "Partial  ")} : '
            f'{span(C_PARTIAL, str(partial)) if partial else span(C_DETAIL, "0")}<br>')
        html_parts.append(div_line(C_DIVIDER))
        html_parts.append('<br>')

        finish_ts = datetime.now()

        for r in results:
            spool_path = r[0]
            success    = r[1]
            detail     = str(r[2])
            tags       = r[3] if len(r) > 3 else []

            if success and "Missing" in detail:
                status_lbl   = "PARTIAL"
                status_color = C_PARTIAL
            elif success:
                status_lbl   = "SUCCESS"
                status_color = C_SUCCESS
            else:
                status_lbl   = "FAILED"
                status_color = C_FAILED

            if tags:
                tags_str  = ", ".join(tags)
                attached  = len(tags)
                selected  = detail.split("Selected:")[-1].split("|")[0].strip() \
                    if "Selected:" in detail else str(attached)
                detail_str = f"Selected:{selected} | Attached:{attached}"
                if "Missing" in detail:
                    missing_part = detail.split("Missing:")[-1].strip()
                    detail_str  += f" | Missing:{missing_part}"
                    detail_color = C_MISSING
                else:
                    detail_color = C_DETAIL
            else:
                tags_str     = "None"
                detail_str   = detail
                detail_color = C_MISSING if "missing" in detail.lower() else C_DETAIL

            ts = finish_ts.strftime("%Y-%m-%d %H:%M:%S")

            html_parts.append(
                f'{span(C_TS, f"[{ts}]")} '
                f'{span(C_LABEL, "Status   :")} '
                f'{span(status_color, status_lbl)}<br>')
            html_parts.append(
                f'          {span(C_LABEL, "Spool No :")} '
                f'{span(C_SPOOL, os.path.basename(spool_path))}<br>')
            html_parts.append(
                f'          {span(C_LABEL, "Tags     :")} '
                f'{span(C_TAGS if tags else C_DETAIL, tags_str)}<br>')
            html_parts.append(
                f'          {span(C_LABEL, "Details  :")} '
                f'{span(detail_color, detail_str)}<br>')
            html_parts.append(sep_line(C_SEP))
            html_parts.append('<br>')

        html_parts.append('</pre>')
        return "".join(html_parts)

    def save_log_txt(self):
        if not self._last_run_results:
            QMessageBox.information(self, "No results", "Run a batch first.")
            return
        fn = os.path.join(
            LOG_FOLDER,
            f"log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt")
        try:
            log_text = self._format_log(self._last_run_results)
            with open(fn, "w", encoding="utf-8") as fh:
                fh.write(log_text)
            QMessageBox.information(self, "Saved", f"Log saved:\n{fn}")
        except Exception as e:
            QMessageBox.warning(self, "Error", str(e))

    def save_log_xlsx(self):
        QMessageBox.information(
            self, "Log",
            f"Logs are saved automatically as .txt files in:\n{LOG_FOLDER}")

    # ── Orchestration ─────────────────────────────────────────
    def run_batch(self):
        if not self.selected_spools:
            QMessageBox.warning(self, "No spools", "Select one or more spool PDFs first.")
            return
        
        excel = self.excel_line.text().strip() or self.settings.get("last_excel", "")
        if not excel or not os.path.exists(excel):
            QMessageBox.warning(self, "Excel missing", "Select a valid Excel file first.")
            return

        self.settings["skip_excel_refresh"]  = bool(self.skip_refresh_chk.isChecked())
        self.settings["manual_support_path"] = self.manual_support_line.text().strip()
        self.settings["last_excel"] = excel
        self._save_settings()

        # Initialize the Signal communicator and connect it
        self.job_signals = JobSignals()
        self.job_signals.status_update.connect(self._update_table_row)

        self.btn_run.setEnabled(False)
        self.btn_stop.setEnabled(True)
        self.btn_exit.setEnabled(False)
        self.btn_clear.setEnabled(False)
        self.results_table.setRowCount(0)
        self.statusBar().showMessage("Starting batch…")
        self.progress.setValue(0)
        self.progress.setFormat("Starting…")
        self.spinner_lbl.start("Attaching supports…")

        # Start miniscreen GIF
        try:
            self.header_widget.start_miniscreen()
        except Exception:
            pass

        self._anim_target  = 0
        self._anim_current = 0.0
        self._anim_timer   = QTimer(self)
        self._anim_timer.setInterval(30)
        self._anim_timer.timeout.connect(self._tick_anim)
        self._anim_timer.start()

        # Clear any leftover row animation state
        for sp in list(self._row_anim_timers.keys()):
            self._stop_row_animation(sp)
        self._row_anim_timers.clear()
        self._row_anim_frames.clear()

        self.stop_event.clear()
        self._batch_start_ts = datetime.now()   # captured for log header
        self.executor = ThreadPoolExecutor(max_workers=self.settings.get("max_workers", DEFAULT_MAX_WORKERS))

        self._future = self.executor.submit(
            self._orchestrate, list(self.selected_spools), excel, self.job_signals)

        self._poll_timer = QTimer()
        self._poll_timer.setInterval(300)
        self._poll_timer.timeout.connect(self._poll_future)
        self._poll_timer.start()

    # Spinner frames for the Status column cell animation
    _STATUS_FRAMES = ["⠋ Processing", "⠙ Processing", "⠹ Processing",
                      "⠸ Processing", "⠼ Processing", "⠴ Processing",
                      "⠦ Processing", "⠧ Processing", "⠇ Processing",
                      "⠏ Processing"]

    def _start_row_animation(self, spool_path: str):
        """Kick off a per-row status-cell spinner for spool_path."""
        if spool_path in self._row_anim_timers:
            return  # already running
        self._row_anim_frames[spool_path] = 0
        timer = QTimer(self)
        timer.setInterval(80)
        timer.timeout.connect(lambda sp=spool_path: self._tick_row_anim(sp))
        timer.start()
        self._row_anim_timers[spool_path] = timer

    def _tick_row_anim(self, spool_path: str):
        fname = os.path.basename(spool_path)
        idx   = (self._row_anim_frames.get(spool_path, 0) + 1) % len(
            self._STATUS_FRAMES)
        self._row_anim_frames[spool_path] = idx
        frame_text = self._STATUS_FRAMES[idx]
        for row in range(self.results_table.rowCount()):
            item = self.results_table.item(row, 0)
            if item and item.text() == fname:
                si = self.results_table.item(row, 1)
                if si:
                    si.setText(frame_text)
                    si.setForeground(QColor("#2eb8ff"))
                break

    def _stop_row_animation(self, spool_path: str):
        timer = self._row_anim_timers.pop(spool_path, None)
        if timer:
            timer.stop()
        self._row_anim_frames.pop(spool_path, None)

    @Slot(str, str, str, str)
    def _update_table_row(self, spool_path, status, detail, tags_text):
        fname = os.path.basename(spool_path)

        # Determine final-state statuses
        is_final = status.lower() in ("completed", "failed", "error",
                                      "✔  ok", "✕  fail")

        # Build a display status item
        def _make_status_item(text, color=None):
            it = QTableWidgetItem(text)
            if color:
                it.setForeground(QColor(color))
            return it

        # Find existing row
        target_row = -1
        for row in range(self.results_table.rowCount()):
            item = self.results_table.item(row, 0)
            if item and item.text() == fname:
                target_row = row
                break

        if target_row == -1:
            # Insert new row
            target_row = self.results_table.rowCount()
            self.results_table.insertRow(target_row)
            self.results_table.setItem(target_row, 0, QTableWidgetItem(fname))

        if is_final:
            # Stop animation
            self._stop_row_animation(spool_path)
            # Set coloured final status
            if "completed" in status.lower() or "ok" in status.lower():
                si = _make_status_item("✔  Completed", "#30c080")
            elif "error" in status.lower() or "fail" in status.lower():
                si = _make_status_item("✕  Failed", "#e04040")
            else:
                si = _make_status_item(status, "#e0a030")
            self.results_table.setItem(target_row, 1, si)
        else:
            # Start spinner animation if not already running
            if spool_path not in self._row_anim_timers:
                # Pre-populate the row with a placeholder
                si = _make_status_item("⠋ Processing", "#2eb8ff")
                self.results_table.setItem(target_row, 1, si)
                self._start_row_animation(spool_path)

        self.results_table.setItem(target_row, 2, QTableWidgetItem(tags_text))
        self.results_table.setItem(target_row, 3, QTableWidgetItem(detail))

    def _stop_batch(self):
        """Immediately signal the background thread to stop after the current spool."""
        self.stop_event.set()
        self.btn_stop.setEnabled(False)
        self.btn_stop.setText("Stopping…")
        self.statusBar().showMessage("⏹  Stop requested — finishing current spool…")
        self.spinner_lbl.stop()

    def _poll_future(self):
        if not self._future.done():
            return
        try:
            results = self._future.result()
        except Exception as e:
            self.statusBar().showMessage(f"Batch failed: {e}")
            results = []

        self._poll_timer.stop()
        try:
            self.executor.shutdown(wait=False)
        except Exception:
            pass

        # Stop animation → snap to 100
        try:
            self._anim_timer.stop()
        except Exception:
            pass
        self.progress.setValue(100)
        self.progress.setFormat("100%  —  Done")
        self.spinner_lbl.stop()

        # Stop header miniscreen GIF (kept in header for visual feedback)
        try:
            self.header_widget.stop_miniscreen()
        except Exception:
            pass

        # Stop all remaining row animations
        for sp in list(self._row_anim_timers.keys()):
            self._stop_row_animation(sp)

        self.btn_run.setEnabled(True)
        self.btn_stop.setEnabled(False)
        self.btn_stop.setText("⏹  Stop")
        self.btn_select.setEnabled(True)
        self.btn_clear.setEnabled(True)
        self.btn_exit.setEnabled(True)
        self._last_run_results = results

        ok   = sum(1 for r in results if r[1])
        fail = sum(1 for r in results if not r[1])
        miss = sum(1 for r in results if r[1] and "Missing" in str(r[2]))

        # ── Populate 4-column table ──
        self.results_table.setRowCount(len(results))
        for idx, res in enumerate(results):
            path    = res[0]
            success = res[1]
            detail  = res[2]
            tags    = res[3] if len(res) > 3 else []

            fi = QTableWidgetItem(os.path.basename(path))
            si = QTableWidgetItem("✔  OK" if success else "✕  FAIL")
            si.setForeground(QColor("#30c080") if success else QColor("#e04040"))

            tags_text = ", ".join(tags) if tags else "—"
            ti = QTableWidgetItem(tags_text)
            ti.setForeground(QColor("#60d8a8") if tags else QColor("#3a6080"))

            di = QTableWidgetItem(str(detail))
            if "Missing" in str(detail):
                di.setForeground(QColor("#e0a030"))

            self.results_table.setItem(idx, 0, fi)
            self.results_table.setItem(idx, 1, si)
            self.results_table.setItem(idx, 2, ti)
            self.results_table.setItem(idx, 3, di)

        self.results_table.resizeRowsToContents()

        self.statusBar().showMessage(
            f"Batch complete — ✔ {ok} succeeded  ✕ {fail} failed  "
            f"⚠ {miss} partial")

        # ── Auto-save structured log ──────────────────────────
        try:
            fn = os.path.join(
                LOG_FOLDER,
                f"log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt")
            log_text = self._format_log(results)
            with open(fn, "w", encoding="utf-8") as fh:
                fh.write(log_text)
            # Update session log view with colour-coded HTML
            self.log_text.setHtml(self._format_log_html(results))
        except Exception:
            pass

        self.progress.setValue(
            int(ok / max(len(results), 1) * 100) if results else 0)
        self.progress.setFormat(f"{ok}/{len(results)} completed")

        QMessageBox.information(
            self, "Batch Complete",
            f"Run finished.\n\n✔ Success: {ok}\n✕ Failed:  {fail}"
            f"\n⚠ Partial: {miss}")

        # Keep failed spools in the list
        remaining = [r[0] for r in results if not r[1]]
        self.selected_spools = remaining
        self.spool_list.clear()
        for path in remaining:
            item = QListWidgetItem(
                f"{os.path.basename(path)}\n{os.path.dirname(path)}")
            item.setData(Qt.UserRole, path)
            self.spool_list.addItem(item)
        self._update_selected_count()

    def _orchestrate(self, spool_paths, excel_path, signals):
        results = []
        try:
            if not self.settings.get("skip_excel_refresh", False):
                if excel_needs_refresh(
                        excel_path,
                        self.settings.get("refresh_minutes",
                                          DEFAULT_REFRESH_MINUTES)):
                    self.statusBar().showMessage("Refreshing Excel…")
                    try:
                        excel_refresh_com(excel_path)
                        self.statusBar().showMessage("Excel refreshed via COM.")
                    except Exception as e:
                        self.statusBar().showMessage(
                            f"COM refresh failed: {e}; trying fallback.")
                        try:
                            excel_refresh_fallback(excel_path)
                            self.statusBar().showMessage(
                                "Excel fallback save done.")
                        except Exception as e2:
                            self.statusBar().showMessage(
                                f"Excel fallback failed: {e2}")
                else:
                    self.statusBar().showMessage(
                        "Excel is recent; skipping refresh.")
            else:
                self.statusBar().showMessage(
                    "Skipping Excel refresh (user choice).")

            try:
                df = pd.read_excel(excel_path, sheet_name=EXCEL_SHEET,
                                   dtype=str)
            except Exception as e:
                self.statusBar().showMessage(f"Failed to read Excel: {e}")
                return results

            df[SPOOL_COLUMN] = df[SPOOL_COLUMN].astype(str).str.strip()
            spool_lookup = {}
            for _, row in df.iterrows():
                key = row[SPOOL_COLUMN]
                if key not in spool_lookup:
                    spool_lookup[key] = []
                tag = str(row.get(SUPPORT_COLUMN, "")).strip()
                if tag and tag.lower() not in ("nan", "none"):
                    spool_lookup[key].append(tag)

            total     = len(spool_paths)
            completed = 0
            start_t   = time.time()

            enable_selection   = self.settings.get(
                "enable_support_selection", False)
            support_selections = {}

            # ── Pre-collect supports & show batch selection dialog ──
            if enable_selection:
                self.statusBar().showMessage("Collecting available supports for all spools…")
                spools_data = []

                for sp in spool_paths:
                    if self.stop_event.is_set():
                        break
                    try:
                        spool_filename = os.path.basename(sp)
                        parsed = parse_spool_filename(
                            spool_filename, self.settings)
                        if not parsed:
                            continue
                        job_number, spool_num = parsed

                        manual = self.settings.get(
                            "manual_support_path", "").strip()
                        recursive_search = False
                        if manual:
                            supports_folder  = manual
                            recursive_search = True
                        else:
                            base_paths = [
                                r"\\epicuaefs02\ENGINEERING\US Projects",
                                r"\\epicuaefs02\ENGINEERING\BendTec Projects"
                            ]
                            jf = find_job_folder(job_number, base_paths)
                            supports_folder = (
                                os.path.join(jf, SUPPORTS_SUBPATH)
                                if jf else os.path.dirname(sp))

                        if not os.path.isdir(supports_folder):
                            continue

                        lookup_key = f"{job_number}-{spool_num}"
                        tags = spool_lookup.get(lookup_key, [])
                        if not tags:
                            if ("Job" in df.columns
                                    and "Spool" in df.columns):
                                rows = df[
                                    (df["Job"].astype(str).str.strip()
                                     == job_number) &
                                    (df["Spool"].astype(str).str.strip()
                                     .str.zfill(6) == spool_num)]
                                tags = []
                                for _, r in rows.iterrows():
                                    t = str(r.get(SUPPORT_COLUMN, "")).strip()
                                    if t and t.lower() not in ("nan", "none"):
                                        tags.append(t)
                        if not tags:
                            continue

                        available_supports = []
                        for tag in tags:
                            if not tag or str(tag).lower() in ("nan", "none"):
                                continue
                            found = None
                            if recursive_search:
                                for root, _, files in os.walk(
                                        supports_folder):
                                    for f in files:
                                        name, ext = os.path.splitext(f)
                                        if (ext.lower() == ".pdf"
                                                and name.lower()
                                                == str(tag).lower()):
                                            found = os.path.join(root, f)
                                            break
                                    if found:
                                        break
                            else:
                                found = find_pdf_for_tag(
                                    supports_folder, str(tag))
                            if found:
                                available_supports.append((found, tag))

                        if available_supports:
                            spools_data.append(
                                (sp, spool_filename, available_supports))
                            self.statusBar().showMessage(
                                f"Found {len(available_supports)} supports "
                                f"for {spool_filename}")
                    except Exception as e:
                        self.statusBar().showMessage(
                            f"Error collecting supports for "
                            f"{os.path.basename(sp)}: {e}")
                        continue

                self.statusBar().showMessage(
                    f"Total spools with supports: {len(spools_data)}")

                if spools_data:
                    self.statusBar().showMessage(
                        f"Showing support selection dialog for "
                        f"{len(spools_data)} spools…")
                    self._pending_dialog_data = spools_data
                    self._dialog_result       = None
                    from PySide6.QtCore import QMetaObject
                    QMetaObject.invokeMethod(
                        self, "_show_batch_selection_dialog",
                        Qt.ConnectionType.BlockingQueuedConnection)
                    support_selections = (
                        self._dialog_result if self._dialog_result else {})
                    self.statusBar().showMessage(
                        f"User made selections for "
                        f"{len(support_selections)} spools")
                else:
                    self.statusBar().showMessage("No supports found for any spool")

            # ── Process each spool ──
            for sp in spool_paths:
                if self.stop_event.is_set():
                    break
                try:
                    spool_filename = os.path.basename(sp)
                    parsed = parse_spool_filename(
                        spool_filename, self.settings)
                    if not parsed:
                        self.statusBar().showMessage(
                            f"File not in config list: {spool_filename}")
                        results.append((sp, False, "file not in config list"))
                        completed += 1
                        self._update_simple_progress(
                            completed, total, start_t)
                        continue

                    job_number, spool_num = parsed

                    manual = self.settings.get(
                        "manual_support_path", "").strip()
                    recursive_search = False
                    if manual:
                        supports_folder  = manual
                        recursive_search = True
                    else:
                        base_paths = [
                            r"\\epicuaefs02\ENGINEERING\US Projects",
                            r"\\epicuaefs02\ENGINEERING\BendTec Projects"
                        ]
                        jf = find_job_folder(job_number, base_paths)
                        supports_folder = (
                            os.path.join(jf, SUPPORTS_SUBPATH)
                            if jf else os.path.dirname(sp))

                    if not os.path.isdir(supports_folder):
                        self.statusBar().showMessage(
                            f"Supports folder missing for "
                            f"{job_number}-{spool_num}: {supports_folder}")
                        results.append((sp, False, "supports folder missing"))
                        completed += 1
                        self._update_simple_progress(
                            completed, total, start_t)
                        continue

                    lookup_key = f"{job_number}-{spool_num}"
                    tags = spool_lookup.get(lookup_key, [])
                    if not tags:
                        if ("Job" in df.columns and "Spool" in df.columns):
                            rows = df[
                                (df["Job"].astype(str).str.strip()
                                 == job_number) &
                                (df["Spool"].astype(str).str.strip()
                                 .str.zfill(6) == spool_num)]
                            tags = []
                            for _, r in rows.iterrows():
                                t = str(r.get(SUPPORT_COLUMN, "")).strip()
                                if t and t.lower() not in ("nan", "none"):
                                    tags.append(t)

                    if not tags:
                        self.statusBar().showMessage(
                            f"No Excel entries for "
                            f"{job_number}-{spool_num}")
                        results.append((sp, False, "no excel entries"))
                        completed += 1
                        self._update_simple_progress(
                            completed, total, start_t)
                        continue

                    out_folder = os.path.join(
                        os.path.dirname(sp), OUTPUT_SUBFOLDER_NAME)
                    os.makedirs(out_folder, exist_ok=True)
                    dest = os.path.join(out_folder, os.path.basename(sp))

                    selected_support_paths = None
                    if enable_selection:
                        if sp in support_selections:
                            selected_support_paths = support_selections[sp]
                            if not selected_support_paths:
                                # User explicitly unchecked all supports
                                self.statusBar().showMessage(
                                    f"All supports deselected for "
                                    f"{spool_filename} — skipping")
                                results.append(
                                    (sp, False, "no supports selected"))
                                completed += 1
                                self._update_simple_progress(
                                    completed, total, start_t)
                                continue
                        else:
                            # Spool was not in the dialog — no supports were
                            # found for it during pre-collection, skip it
                            self.statusBar().showMessage(
                                f"No supports found for {spool_filename} — skipping")
                            results.append(
                                (sp, False, "no supports found"))
                            completed += 1
                            self._update_simple_progress(
                                completed, total, start_t)
                            continue

                    job = MergeJob(sp, supports_folder, tags, dest,
                       recursive_search, selected_support_paths, signals=signals)
                    res = job.run()
                    results.append(res)
                    self.statusBar().showMessage(
                        f"{os.path.basename(sp)} → "
                        f"{'OK' if res[1] else 'ERROR'}  {res[2]}")

                except Exception as e:
                    results.append((sp, False, str(e), []))
                    self.statusBar().showMessage(
                        f"Exception: {sp}: {e}\n{traceback.format_exc()}")
                finally:
                    completed += 1
                    self._update_simple_progress(
                        completed, total, start_t,
                        current_file=os.path.basename(sp))

            return results

        except Exception as e:
            self.statusBar().showMessage(
                f"Orchestrator error: {e}\n{traceback.format_exc()}")
            return results

    def _tick_anim(self):
        """Smoothly ease progress bar toward _anim_target."""
        try:
            target = getattr(self, "_anim_target", 0)
            cur    = getattr(self, "_anim_current", 0.0)
            if cur < target:
                step = max(0.3, (target - cur) * 0.10)
                cur  = min(cur + step, float(target))
                self._anim_current = cur
                self.progress.setValue(int(cur))
                self.progress.setFormat(f"{int(cur)}%")
        except Exception:
            pass

    def _update_simple_progress(self, completed, total, start_time,
                                current_file=""):
        try:
            pct     = int((completed / total) * 100) if total else 100
            elapsed = time.time() - start_time if start_time else 0
            if completed > 0:
                remaining = int((elapsed / completed) * (total - completed))
            else:
                remaining = 0
            if remaining >= 3600:
                eta = f"{remaining // 3600}h {(remaining % 3600) // 60}m"
            elif remaining >= 60:
                eta = f"{remaining // 60}m {remaining % 60}s"
            else:
                eta = f"{remaining}s"

            fname = os.path.basename(current_file) if current_file else ""

            def update_ui():
                self._anim_target = pct
                if fname:
                    self.spinner_lbl.set_file(fname)
                self.statusBar().showMessage(
                    f"Processing {completed}/{total}  ({pct}%)  ETA: {eta}")

            QTimer.singleShot(0, update_ui)
        except Exception:
            pass

    @Slot()
    def _show_batch_selection_dialog(self):
        try:
            dlg = BatchSupportSelectionDialog(
                self._pending_dialog_data, self)
            if dlg.exec() == QDialog.Accepted:
                self._dialog_result = dlg.get_selections()
            else:
                self._dialog_result = {}
        except Exception as e:
            self.statusBar().showMessage(f"Dialog error: {e}")
            self._dialog_result = {}

    # ── Shutdown ──────────────────────────────────────────────
    def close_event(self):
        """Called by the Exit button — delegates to Qt's closeEvent."""
        self.close()

    def closeEvent(self, event):
        """Qt native close handler — called by X button AND self.close()."""
        reply = QMessageBox.question(
            self, "Exit", "Are you sure you want to exit?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No)
        if reply == QMessageBox.StandardButton.Yes:
            event.accept()
            self._shutdown_and_exit()
        else:
            event.ignore()

    def _shutdown_and_exit(self):
        try:
            self.stop_event.set()
            if self.executor:
                self.executor.shutdown(wait=False)
        except Exception:
            pass
        # Stop all row animations
        for sp in list(getattr(self, "_row_anim_timers", {}).keys()):
            try:
                self._stop_row_animation(sp)
            except Exception:
                pass
        release_single_instance()
        self._save_settings()
        QApplication.quit()

    def _bring_to_front(self):
        try:
            if self.windowState() & Qt.WindowState.WindowMinimized:
                self.setWindowState(
                    self.windowState() & ~Qt.WindowState.WindowMinimized)
            self.show()
            self.raise_()
            self.activateWindow()
        except Exception:
            pass


# ═══════════════════════════════════════════════════════════════
#  SPLASH (video intro)
# ═══════════════════════════════════════════════════════════════
def show_video_splash(app):
    screen_geo = app.primaryScreen().geometry()
    sw, sh     = screen_geo.width(), screen_geo.height()
    splash_w   = max(900, int(sw * 0.70))
    splash_h   = max(560, int(sh * 0.65))

    splash = QWidget()
    splash.setWindowFlags(Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint)
    splash.setFixedSize(splash_w, splash_h)
    splash.move(
        screen_geo.left() + (sw - splash_w) // 2,
        screen_geo.top()  + (sh - splash_h) // 2,
    )

    splash.setStyleSheet("""
        QWidget#ContainerWidget {
            background-color: transparent;
            border: 2px solid #2eb8ff;
            border-radius: 12px;
        }
        QLabel { background: transparent; }
        QLabel#title {
            color: #2eb8ff;
            font-size: 24pt;
            font-weight: 900;
            letter-spacing: 2px;
        }
        QLabel#dev {
            color: #8ab4cc;
            font-size: 11pt;
            font-weight: bold;
        }
        QLabel#status {
            color: #4a7a9a;
            font-size: 10pt;
        }
        QProgressBar {
            background: #000000;
            border: 1px solid #0d2040;
            border-radius: 6px;
            text-align: center;
            color: white;
            font-weight: bold;
            height: 22px;
        }
        QProgressBar::chunk {
            background: qlineargradient(
                x1:0, y1:0, x2:1, y2:0,
                stop:0 #ff0000,
                stop:1 #00ffcc
            );
            border-radius: 5px;
        }
    """)

    # Layer 1: Background GIF
    gif_label = QLabel(splash)
    gif_label.setGeometry(0, 0, splash_w, splash_h)
    gif_label.setScaledContents(True)

    movie = QMovie(resource_path("background.gif"))
    gif_label.setMovie(movie)
    movie.start()

    # Layer 2: Translucent Tint Overlay
    overlay = QWidget(splash)
    overlay.setGeometry(0, 0, splash_w, splash_h)
    overlay.setStyleSheet("background-color: rgba(5, 10, 24, 190); border-radius: 10px;")

    # Layer 3: Text & Progress Container (Top Layer)
    content_container = QWidget(splash)
    content_container.setObjectName("ContainerWidget")
    content_container.setGeometry(0, 0, splash_w, splash_h)
    content_container.setAttribute(Qt.WA_TranslucentBackground, True)
    content_container.raise_()

    layout = QVBoxLayout(content_container)
    layout.setContentsMargins(40, 40, 40, 40)

    title = QLabel("SPOOL SUPPORT SYNC")
    title.setObjectName("title")
    title.setAlignment(Qt.AlignCenter)
    layout.addWidget(title)

    dev = QLabel("By Nivin Varkey")
    dev.setObjectName("dev")
    dev.setAlignment(Qt.AlignCenter)
    layout.addWidget(dev)

    layout.addStretch()

    status = QLabel("Initializing...")
    status.setObjectName("status")
    layout.addWidget(status)

    progress = QProgressBar()
    layout.addWidget(progress)

    splash.show()

    milestones = {
        0: "Initializing application...",
        20: "Loading PDF engine...",
        40: "Reading configuration...",
        60: "Preparing workspace...",
        80: "Finalizing modules...",
        100: "Launching..."
    }

    effect = QGraphicsOpacityEffect(status)
    status.setGraphicsEffect(effect)

    anim = QPropertyAnimation(effect, b"opacity")
    anim.setDuration(800)
    anim.setStartValue(1.0)
    anim.setEndValue(0.4)
    anim.setEasingCurve(QEasingCurve.InOutQuad)

    for i in range(101):
        progress.setValue(i)

        if i in milestones:
            status.setText(milestones[i])

            if anim.state() != QPropertyAnimation.Running:
                anim.setDirection(
                    QPropertyAnimation.Forward
                    if anim.direction() == QPropertyAnimation.Backward
                    else QPropertyAnimation.Backward
                )
                anim.start()

        app.processEvents()
        time.sleep(0.020)  # Reduced from 0.055 → faster splash, less delay before main window

    anim.stop()
    movie.stop()  
    splash.close()

# ═══════════════════════════════════════════════════════════════
#  ENTRY POINT
# ═══════════════════════════════════════════════════════════════
def main():
    app = QApplication(sys.argv)

    ico = resource_path("logo.ico")
    if os.path.exists(ico):
        try:
            app.setWindowIcon(QIcon(ico))
        except Exception:
            pass

    w = MainWindow()

    if os.path.exists(ico):
        try:
            w.setWindowIcon(QIcon(ico))
        except Exception:
            pass

    # Check update before splash and before main window opens
    try:
        available, info = check_update_available()

        if available and info:
            dlg = UpdateConfirmDialog(info, w)
            dlg.exec()

    except Exception as ex:
        QMessageBox.critical(
            None,
            "Update Check Failed",
            str(ex)
        )

    show_video_splash(app)

    w.show()
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())