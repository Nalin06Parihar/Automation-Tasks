import os
import subprocess
from typing import Any


def safe_text(value: Any, default: str = "Unavailable") -> str:
    if value is None:
        return default
    if isinstance(value, str):
        cleaned = value.strip()
        return cleaned if cleaned else default
    return str(value)


def format_bytes(value: Any) -> str:
    if value in (None, "", "Unavailable"):
        return "Unavailable"
    try:
        number = float(value)
    except (TypeError, ValueError):
        return "Unavailable"
    units = ["B", "KB", "MB", "GB", "TB", "PB"]
    size = number
    unit_index = 0
    while size >= 1024 and unit_index < len(units) - 1:
        size /= 1024
        unit_index += 1
    if unit_index == 0:
        return f"{int(size)} {units[unit_index]}"
    return f"{size:.2f} {units[unit_index]}"


def run_command(command: list[str] | tuple[str, ...], shell: bool = False) -> str:
    try:
        result = subprocess.run(command, shell=shell, capture_output=True, text=True, check=False)
    except (OSError, ValueError):
        return ""
    if result.stdout:
        return result.stdout.strip()
    return result.stderr.strip()


def read_text_file(path: str) -> str:
    try:
        with open(path, "r", encoding="utf-8", errors="replace") as handle:
            return handle.read()
    except OSError:
        return ""


def first_available(*values: Any) -> Any:
    for value in values:
        if value not in (None, "", "Unavailable"):
            return value
    return "Unavailable"


def windows_reg_value(raw_value: Any) -> str:
    if raw_value is None:
        return "Unavailable"
    if isinstance(raw_value, (int, float)):
        return str(raw_value)
    return str(raw_value).strip() or "Unavailable"
