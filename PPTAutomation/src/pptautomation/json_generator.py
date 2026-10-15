from __future__ import annotations

from typing import Any


FIELD_LABELS = {
    "text": "Text on slide",
    "font_size": "Font size (pt)",
    "font_color": "Text color",
    "bold": "Bold",
    "italic": "Italic",
    "fill_color": "Fill color",
    "line_color": "Border color",
}

FIELD_ORDER = [
    "text",
    "font_size",
    "font_color",
    "bold",
    "italic",
    "fill_color",
    "line_color",
]


def supported_editable_fields(shape: dict[str, Any]) -> list[str]:
    properties = shape.get("properties", {})
    shape_type = shape.get("type")
    fields: list[str] = []

    for field_name in FIELD_ORDER:
        if field_name in properties:
            fields.append(field_name)

    return fields


def build_default_input(template: dict[str, Any]) -> dict[str, Any]:
    default_input: dict[str, Any] = {}

    for slide in template.get("slides", []):
        slide_key = f"slide_{slide.get('slide')}"
        default_input[slide_key] = {}
        for shape in slide.get("shapes", []):
            shape_name = shape.get("name", f"Shape {shape.get('id')}")
            if not supported_editable_fields(shape):
                continue
            default_input[slide_key][shape_name] = {}
            for field_name in supported_editable_fields(shape):
                default_input[slide_key][shape_name][field_name] = shape.get("properties", {}).get(field_name)

    return default_input
