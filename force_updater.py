import os
import sys
import json
import shutil
import tempfile
import subprocess

from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QLabel, QProgressBar,
    QPushButton, QHBoxLayout, QApplication, QMessageBox
)
from PySide6.QtCore import Qt, QThread, Signal


VERSION_FILE = r"\\epicuaefs03\UAE Data\Document_Control\Engineering\U.S-Jobs\09 Spool Support Sync\version.json"
CURRENT_VERSION = "1.9"


def version_tuple(v):
    return tuple(int(x) for x in str(v).replace("V", "").split("."))


def check_update_available():
    try:
        if not os.path.exists(VERSION_FILE):
            return False, None

        with open(VERSION_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)

        latest = data.get("version", CURRENT_VERSION)

        if version_tuple(latest) > version_tuple(CURRENT_VERSION):
            return True, data

        return False, data

    except Exception as e:
        print("Update check failed:", e)
        return False, None


class UpdateWorker(QThread):
    progress = Signal(int, str)
    finished = Signal(str)
    failed = Signal(str)

    def __init__(self, source_exe):
        super().__init__()
        self.source_exe = source_exe

    def run(self):
        try:
            if not os.path.exists(self.source_exe):
                self.failed.emit("Update EXE not found on network.")
                return

            temp_exe = os.path.join(tempfile.gettempdir(), "SpoolSupportSync_new.exe")

            total_size = os.path.getsize(self.source_exe)
            copied = 0

            with open(self.source_exe, "rb") as src, open(temp_exe, "wb") as dst:
                while True:
                    chunk = src.read(1024 * 1024)
                    if not chunk:
                        break

                    dst.write(chunk)
                    copied += len(chunk)

                    percent = int((copied / total_size) * 100)
                    self.progress.emit(percent, f"Updating... {percent}%")

            self.finished.emit(temp_exe)

        except Exception as e:
            self.failed.emit(str(e))


class UpdateConfirmDialog(QDialog):
    def __init__(self, update_info, parent=None):
        super().__init__(parent)

        self.update_info = update_info
        self.worker = None

        self.setWindowTitle("Update Required")
        self.setFixedSize(480, 260)

        self.setWindowFlags(
            Qt.Dialog |
            Qt.CustomizeWindowHint |
            Qt.WindowTitleHint
        )

        layout = QVBoxLayout(self)

        title = QLabel("Spool Support Sync Update Available")
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("font-size: 14pt; font-weight: bold; color: #2eb8ff;")
        layout.addWidget(title)

        latest = update_info.get("version", "")
        notes = "\n".join(update_info.get("notes", []))

        self.info_label = QLabel(
            f"Current Version: V{CURRENT_VERSION}\n"
            f"Latest Version : V{latest}\n\n"
            f"Changes:\n{notes}\n\n"
            "This update is required to continue using the application."
        )
        self.info_label.setWordWrap(True)
        layout.addWidget(self.info_label)

        self.progress = QProgressBar()
        self.progress.setRange(0, 100)
        self.progress.setValue(0)
        self.progress.setFormat("Waiting for confirmation...")
        layout.addWidget(self.progress)

        self.status_label = QLabel("Do you want to update now?")
        self.status_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.status_label)

        btn_row = QHBoxLayout()

        self.btn_update = QPushButton("Update Now")
        self.btn_update.clicked.connect(self.start_update)

        self.btn_exit = QPushButton("Exit App")
        self.btn_exit.clicked.connect(self.exit_app)

        btn_row.addWidget(self.btn_update)
        btn_row.addWidget(self.btn_exit)

        layout.addLayout(btn_row)

    def closeEvent(self, event):
        event.ignore()

    def exit_app(self):
        QApplication.quit()

    def start_update(self):
        self.btn_update.setEnabled(False)
        self.btn_exit.setEnabled(False)

        self.status_label.setText("Starting update...")
        self.progress.setFormat("Starting update...")

        source_exe = self.update_info.get("exe_path")

        self.worker = UpdateWorker(source_exe)
        self.worker.progress.connect(self.on_progress)
        self.worker.finished.connect(self.on_finished)
        self.worker.failed.connect(self.on_failed)
        self.worker.start()

    def on_progress(self, percent, text):
        self.progress.setValue(percent)
        self.progress.setFormat(text)
        self.status_label.setText(text)

    def on_failed(self, error):
        self.status_label.setText(f"Update failed: {error}")
        self.progress.setFormat("Update failed")

        QMessageBox.critical(
            self,
            "Update Failed",
            f"Update failed:\n\n{error}\n\nApplication will close."
        )

        QApplication.quit()

    def on_finished(self, new_exe):
        self.progress.setValue(100)
        self.progress.setFormat("Installing update...")
        self.status_label.setText("Installing update and restarting...")

        current_exe = sys.executable

        bat_path = os.path.join(tempfile.gettempdir(), "SpoolSupportSync_Update.bat")

        bat = f"""
@echo off
timeout /t 2 /nobreak >nul
copy /Y "{new_exe}" "{current_exe}"
start "" "{current_exe}"
del "{new_exe}"
del "%~f0"
"""

        with open(bat_path, "w", encoding="utf-8") as f:
            f.write(bat)

        subprocess.Popen(
            bat_path,
            shell=True,
            creationflags=subprocess.CREATE_NO_WINDOW
        )

        QApplication.quit()