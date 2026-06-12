import os
import sys
import json
import tempfile
import subprocess
import urllib.request
import urllib.error

from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QLabel, QProgressBar,
    QPushButton, QHBoxLayout, QApplication, QMessageBox
)
from PySide6.QtCore import Qt, QThread, Signal


# =====================================================================
#  GITHUB UPDATE SETTINGS
# =====================================================================
CURRENT_VERSION = "1.9"

# CHANGE THESE 3 VALUES AFTER YOU CREATE YOUR GITHUB REPOSITORY
GITHUB_OWNER = "nivinvarkey"
GITHUB_REPO = "SpoolSupportSync"
ASSET_NAME = "SpoolSupportSync.exe"

LATEST_RELEASE_API = (
    f"https://api.github.com/repos/{GITHUB_OWNER}/{GITHUB_REPO}/releases/latest"
)


# =====================================================================
#  HELPERS
# =====================================================================
def version_tuple(version_text):
    """
    Converts v1.9 / V1.9 / 1.9 into comparable tuple: (1, 9).
    """
    clean = str(version_text).strip().replace("V", "").replace("v", "")
    parts = []
    for part in clean.split("."):
        digits = "".join(ch for ch in part if ch.isdigit())
        parts.append(int(digits or 0))
    return tuple(parts)


def check_update_available():
    """
    Checks latest GitHub Release.

    Required release setup:
      Tag name: v2.0 / v2.1 / v3.0 etc.
      Asset name: SpoolSupportSync.exe

    Returns:
      (True, info)  -> update available
      (False, info) -> no update
      (False, None) -> check failed
    """
    try:
        request = urllib.request.Request(
            LATEST_RELEASE_API,
            headers={
                "Accept": "application/vnd.github+json",
                "User-Agent": "SpoolSupportSync-Updater"
            },
        )

        with urllib.request.urlopen(request, timeout=15) as response:
            data = json.loads(response.read().decode("utf-8"))

        latest_version = str(data.get("tag_name", CURRENT_VERSION)).replace("v", "").replace("V", "")
        release_notes = data.get("body", "") or ""

        download_url = None
        for asset in data.get("assets", []):
            if asset.get("name") == ASSET_NAME:
                download_url = asset.get("browser_download_url")
                break

        info = {
            "version": latest_version,
            "notes": release_notes.splitlines() if release_notes else [],
            "download_url": download_url,
            "release_url": data.get("html_url", ""),
            "asset_name": ASSET_NAME,
        }

        if not download_url:
            print(f"Update asset not found in GitHub release: {ASSET_NAME}")
            return False, info

        if version_tuple(latest_version) > version_tuple(CURRENT_VERSION):
            return True, info

        return False, info

    except Exception as e:
        print("Update check failed:", e)
        return False, None


# =====================================================================
#  DOWNLOAD WORKER
# =====================================================================
class UpdateWorker(QThread):
    progress = Signal(int, str)
    finished = Signal(str)
    failed = Signal(str)

    def __init__(self, download_url):
        super().__init__()
        self.download_url = download_url

    def run(self):
        try:
            if not self.download_url:
                self.failed.emit("GitHub download URL was not found.")
                return

            temp_exe = os.path.join(tempfile.gettempdir(), "SpoolSupportSync_new.exe")

            request = urllib.request.Request(
                self.download_url,
                headers={"User-Agent": "SpoolSupportSync-Updater"},
            )

            with urllib.request.urlopen(request, timeout=60) as response:
                total_size = int(response.headers.get("Content-Length", 0) or 0)
                copied = 0

                with open(temp_exe, "wb") as dst:
                    while True:
                        chunk = response.read(1024 * 1024)
                        if not chunk:
                            break

                        dst.write(chunk)
                        copied += len(chunk)

                        if total_size > 0:
                            percent = int((copied / total_size) * 100)
                        else:
                            percent = 0

                        self.progress.emit(percent, f"Downloading update... {percent}%")

            if not os.path.exists(temp_exe) or os.path.getsize(temp_exe) == 0:
                self.failed.emit("Downloaded update file is empty.")
                return

            self.finished.emit(temp_exe)

        except urllib.error.HTTPError as e:
            self.failed.emit(f"GitHub download failed: HTTP {e.code}")
        except urllib.error.URLError as e:
            self.failed.emit(f"GitHub download failed: {e.reason}")
        except Exception as e:
            self.failed.emit(str(e))


# =====================================================================
#  UPDATE DIALOG
# =====================================================================
class UpdateConfirmDialog(QDialog):
    def __init__(self, update_info, parent=None, required=False):
        super().__init__(parent)

        self.update_info = update_info or {}
        self.worker = None
        self.required = required

        self.setWindowTitle("Update Available")
        self.setFixedSize(520, 300)

        if required:
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

        latest = self.update_info.get("version", "")
        notes = "\n".join(self.update_info.get("notes", []))
        if not notes.strip():
            notes = "No release notes provided."

        self.info_label = QLabel(
            f"Current Version: V{CURRENT_VERSION}\n"
            f"Latest Version : V{latest}\n\n"
            f"Changes:\n{notes}\n\n"
            "Click Install Update to download from GitHub and restart the app."
        )
        self.info_label.setWordWrap(True)
        layout.addWidget(self.info_label)

        self.progress = QProgressBar()
        self.progress.setRange(0, 100)
        self.progress.setValue(0)
        self.progress.setFormat("Waiting...")
        layout.addWidget(self.progress)

        self.status_label = QLabel("Update is ready to install.")
        self.status_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.status_label)

        btn_row = QHBoxLayout()

        self.btn_update = QPushButton("Install Update")
        self.btn_update.clicked.connect(self.start_update)

        self.btn_later = QPushButton("Later")
        self.btn_later.clicked.connect(self.reject)

        btn_row.addWidget(self.btn_update)
        btn_row.addWidget(self.btn_later)

        layout.addLayout(btn_row)

    def closeEvent(self, event):
        if self.required:
            event.ignore()
        else:
            event.accept()

    def start_update(self):
        self.btn_update.setEnabled(False)
        self.btn_later.setEnabled(False)

        self.status_label.setText("Downloading update from GitHub...")
        self.progress.setFormat("Downloading update...")

        download_url = self.update_info.get("download_url")

        self.worker = UpdateWorker(download_url)
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
            f"Update failed:\n\n{error}\n\nPlease check the GitHub release asset and internet connection."
        )

        self.btn_update.setEnabled(True)
        self.btn_later.setEnabled(True)

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