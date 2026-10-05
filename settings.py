"""Persist the selected location, keeping compatibility with version 1.0."""

from __future__ import annotations

import ctypes
import json
import os
import sys
import tempfile
from pathlib import Path

APP_NAME = 'Docx Meta Cleaner'


def settings_file_path() -> Path:
    base = Path(os.environ.get('APPDATA', str(Path.home()))) if sys.platform == 'win32' else Path.home() / '.config'
    return base / APP_NAME / 'settings.json'


def desktop_directory() -> Path:
    if sys.platform == 'win32':
        # CSIDL_DESKTOPDIRECTORY resolves the physical folder, including redirection.
        buffer = ctypes.create_unicode_buffer(32768)
        if ctypes.windll.shell32.SHGetFolderPathW(None, 0x10, None, 0, buffer) == 0:
            path = Path(buffer.value)
            if path.is_dir():
                return path
    return Path.home() / 'Desktop'


def load_settings(path: Path) -> dict:
    try:
        data = json.loads(path.read_text(encoding='utf-8'))
        return data if isinstance(data, dict) else {}
    except (OSError, ValueError):
        return {}


def save_settings(path: Path, data: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, name = tempfile.mkstemp(prefix='.settings-', suffix='.json', dir=path.parent)
    temporary = Path(name)
    try:
        with os.fdopen(fd, 'w', encoding='utf-8') as output:
            json.dump(data, output, ensure_ascii=False, indent=2)
            output.flush()
            os.fsync(output.fileno())
        os.replace(temporary, path)
    finally:
        temporary.unlink(missing_ok=True)


def restored_selection(data: dict) -> tuple[Path | None, str]:
    selected = data.get('selected_path') or data.get('last_directory')
    if not isinstance(selected, str) or not selected:
        return None, ''
    path = Path(selected).expanduser()
    if path.is_dir():
        return path.resolve(), 'folder'
    if path.is_file() and data.get('selection_mode') == 'file':
        return path.resolve(), 'file'
    return None, ''


def dialog_directory(data: dict) -> str:
    value = data.get('last_directory')
    if isinstance(value, str) and value and Path(value).is_dir():
        return str(Path(value).resolve())
    desktop = desktop_directory()
    return str(desktop if desktop.is_dir() else Path.home())
