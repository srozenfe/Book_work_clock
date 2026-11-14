"""Private GitHub Releases updater for Book Work Clock.

The user's read-only GitHub token is stored in Windows Credential Manager via
keyring, never in the SQLite database or source tree. Only a published GitHub
release with a SHA-256 digest is accepted as an update.
"""
from __future__ import annotations

import hashlib
import os
import re
import subprocess
import sys
from pathlib import Path
from urllib.parse import urlparse

import keyring
import requests
from packaging.version import Version, InvalidVersion
from PySide6.QtWidgets import QDialog, QDialogButtonBox, QFormLayout, QLineEdit, QMessageBox

APP_VERSION = "0.3.0"
REPOSITORY = "srozenfe/Book_work_clock"
API_URL = f"https://api.github.com/repos/{REPOSITORY}/releases/latest"
CREDENTIAL_SERVICE = "BookWorkClock-GitHub-Readonly"
CREDENTIAL_ACCOUNT = REPOSITORY


class CredentialDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Connect private GitHub repository")
        self.input = QLineEdit()
        self.input.setEchoMode(QLineEdit.EchoMode.Password)
        self.input.setPlaceholderText("Fine-grained GitHub read-only token")
        form = QFormLayout(self)
        form.addRow("GitHub token:", self.input)
        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Save |
            QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        form.addRow(buttons)


def get_token():
    return keyring.get_password(CREDENTIAL_SERVICE, CREDENTIAL_ACCOUNT)


def save_token(token):
    keyring.set_password(CREDENTIAL_SERVICE, CREDENTIAL_ACCOUNT, token.strip())


def github_headers(token, binary=False):
    headers = {
        "Authorization": f"Bearer {token}",
        "X-GitHub-Api-Version": "2022-11-28",
        "User-Agent": f"BookWorkClock/{APP_VERSION}",
        "Accept": "application/octet-stream" if binary
                  else "application/vnd.github+json",
    }
    return headers


def latest_release(token):
    response = requests.get(API_URL, headers=github_headers(token), timeout=15)
    response.raise_for_status()
    release = response.json()
    try:
        latest_version = Version(str(release["tag_name"]).lstrip("vV"))
        current_version = Version(APP_VERSION)
    except (KeyError, InvalidVersion) as exc:
        raise RuntimeError("Invalid GitHub release version") from exc
    if latest_version <= current_version:
        return None
    asset = next(
        (a for a in release.get("assets", [])
         if a.get("name") == "BookWorkClock-Setup.exe"), None
    )
    if not asset:
        raise RuntimeError("Latest release has no BookWorkClock-Setup.exe asset")
    digest = asset.get("digest") or ""
    if not re.fullmatch(r"sha256:[0-9a-fA-F]{64}", digest):
        raise RuntimeError("Latest release has no verifiable SHA-256 digest")
    return release, asset, digest.split(":", 1)[1].lower()


def download_installer(token, asset, expected_digest):
    # Never download from a user-supplied URL: only the GitHub release asset API.
    url = asset.get("url", "")
    expected_prefix = f"https://api.github.com/repos/{REPOSITORY}/releases/assets/"
    if not url.startswith(expected_prefix) or not url[len(expected_prefix):].isdigit():
        raise RuntimeError("Unexpected release asset URL")
    folder = Path(os.getenv("LOCALAPPDATA", str(Path.home()))) / "BookWorkClock" / "updates"
    folder.mkdir(parents=True, exist_ok=True)
    destination = folder / "BookWorkClock-Setup.exe"
    temporary = folder / "BookWorkClock-Setup.exe.part"
    digest = hashlib.sha256()
    try:
        with requests.get(
            url, headers=github_headers(token, binary=True),
            stream=True, timeout=(15, 90)
        ) as response:
            response.raise_for_status()
            # requests strips Authorization when redirected to another host.
            with temporary.open("wb") as handle:
                for chunk in response.iter_content(chunk_size=1024 * 1024):
                    if chunk:
                        digest.update(chunk)
                        handle.write(chunk)
        if digest.hexdigest() != expected_digest:
            raise RuntimeError("Downloaded installer SHA-256 mismatch; update cancelled")
        with temporary.open("rb") as handle:
            if handle.read(2) != b"MZ":
                raise RuntimeError("Downloaded asset is not a Windows executable")
        temporary.replace(destination)
        return destination
    finally:
        temporary.unlink(missing_ok=True)


def launch_installer(installer):
    """Exit the old application, install, and reopen the updated executable."""
    if not getattr(sys, "frozen", False):
        raise RuntimeError("In-app installation requires the installed Windows executable")
    installed_exe = Path(sys.executable)
    # The script runs outside the Python application, allowing its executable
    # and Qt libraries to be replaced after the application closes.
    script = installer.parent / "finish-update.cmd"
    script.write_text(
        "@echo off\n"
        "timeout /t 3 /nobreak >nul\n"
        f'start /wait "" "{installer}" /SILENT /SUPPRESSMSGBOXES /NORESTART /CLOSEAPPLICATIONS\n'
        "if errorlevel 1 exit /b 1\n"
        f'start "" "{installed_exe}"\n',
        encoding="utf-8",
    )
    subprocess.Popen(
        ["cmd.exe", "/c", str(script)],
        creationflags=subprocess.CREATE_NO_WINDOW | subprocess.CREATE_NEW_PROCESS_GROUP,
        close_fds=True,
    )


def check_for_updates(parent):
    if not getattr(sys, "frozen", False):
        QMessageBox.information(
            parent, "Developer mode",
            "Update installation is supported in the installed Windows executable."
        )
        return False
    try:
        token = get_token()
        if not token:
            if QMessageBox.question(
                parent, "Private GitHub access",
                "This repository is private. To check for updates, enter a fine-grained "
                "read-only GitHub token once. Open the token entry window?"
            ) != QMessageBox.StandardButton.Yes:
                return False
            dialog = CredentialDialog(parent)
            if dialog.exec() != QDialog.DialogCode.Accepted or not dialog.input.text().strip():
                return False
            token = dialog.input.text().strip()
            save_token(token)
        result = latest_release(token)
        if result is None:
            QMessageBox.information(parent, "Book Work Clock", "You're using the latest published version.")
            return False
        release, asset, digest = result
        if QMessageBox.question(
            parent, "Software update available",
            f"Version {release['tag_name']} is available (installed: {APP_VERSION}).\n\n"
            "Download the verified installer, update, and restart the application?"
        ) != QMessageBox.StandardButton.Yes:
            return False
        installer = download_installer(token, asset, digest)
        launch_installer(installer)
        return True
    except (requests.RequestException, OSError, RuntimeError, keyring.errors.KeyringError) as exc:
        QMessageBox.warning(parent, "Update check failed", str(exc))
        return False
