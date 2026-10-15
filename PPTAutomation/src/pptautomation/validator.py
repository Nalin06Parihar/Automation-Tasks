from __future__ import annotations

from typing import BinaryIO


def validate_uploaded_pptx(file_obj: BinaryIO | None, filename: str | None = None) -> None:
    if file_obj is None:
        raise ValueError("No file was uploaded.")

    name = (filename or getattr(file_obj, "name", "")).lower()
    if not name.endswith(".pptx"):
        raise ValueError("Please upload a .pptx file.")


def validate_color(color_value: str | None) -> str | None:
    if color_value in (None, ""):
        return None

    value = str(color_value).strip()
    if not value.startswith("#"):
        raise ValueError("Color values must use a hex format such as #0078D4.")

    value = value[1:]
    if len(value) not in {3, 6}:
        raise ValueError("Color values must use a 3-digit or 6-digit hex format.")

    try:
        int(value, 16)
    except ValueError as exc:
        raise ValueError("Color value is not valid hex.") from exc

    if len(value) == 3:
        value = "".join(ch * 2 for ch in value)

    return f"#{value.upper()}"


def validate_numeric(value: object, field_name: str) -> float | None:
    if value in (None, ""):
        return None

    try:
        return float(value)
    except (TypeError, ValueError) as exc:
        raise ValueError(f"{field_name} must be numeric.") from exc
