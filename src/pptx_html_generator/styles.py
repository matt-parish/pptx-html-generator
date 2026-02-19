"""Style and unit parsing helpers."""

from __future__ import annotations

import re
from typing import Any

from pptx.dml.color import RGBColor
from pptx.util import Cm, Emu, Inches, Pt

_DIMENSION_RE = re.compile(
    r"^\s*(?P<value>-?\d+(?:\.\d+)?)\s*(?P<unit>in|cm|pt|emu)\s*$",
    re.IGNORECASE,
)
_HEX_COLOR_RE = re.compile(r"^#(?P<hex>[0-9a-f]{3}|[0-9a-f]{6})$", re.IGNORECASE)
_RGB_COLOR_RE = re.compile(
    r"^rgb\(\s*(?P<r>\d{1,3})\s*,\s*(?P<g>\d{1,3})\s*,\s*(?P<b>\d{1,3})\s*\)$",
    re.IGNORECASE,
)
_FONT_SIZE_RE = re.compile(r"^\s*(?P<value>-?\d+(?:\.\d+)?)\s*(?P<unit>pt|px)\s*$", re.IGNORECASE)

_NAMED_COLORS = {
    "black": "000000",
    "silver": "C0C0C0",
    "gray": "808080",
    "white": "FFFFFF",
    "maroon": "800000",
    "red": "FF0000",
    "purple": "800080",
    "fuchsia": "FF00FF",
    "green": "008000",
    "lime": "00FF00",
    "olive": "808000",
    "yellow": "FFFF00",
    "navy": "000080",
    "blue": "0000FF",
    "teal": "008080",
    "aqua": "00FFFF",
    "orange": "FFA500",
}


def parse_dimension(value: Any):
    """Parse a dimension string like '1in' into a python-pptx length."""
    if not isinstance(value, str):
        raise ValueError(f"Dimension must be a string, got {type(value).__name__}")

    match = _DIMENSION_RE.match(value)
    if not match:
        raise ValueError(
            f"Invalid dimension '{value}'. Expected format like '1in', '2.5cm', '18pt', '914400emu'."
        )

    number = float(match.group("value"))
    unit = match.group("unit").lower()

    if unit == "in":
        return Inches(number)
    if unit == "cm":
        return Cm(number)
    if unit == "pt":
        return Pt(number)
    if unit == "emu":
        return Emu(int(number))

    raise ValueError(f"Unsupported unit '{unit}'")


def parse_css_inline_style(style_value: str | None) -> dict[str, str]:
    """Parse CSS declarations from a style attribute."""
    if not style_value:
        return {}

    declarations: dict[str, str] = {}
    for chunk in style_value.split(";"):
        if not chunk.strip() or ":" not in chunk:
            continue
        key, value = chunk.split(":", 1)
        key = key.strip().lower()
        value = value.strip()
        if key:
            declarations[key] = value
    return declarations


def parse_css_color(color_value: str) -> RGBColor:
    """Parse CSS color string into RGBColor."""
    value = color_value.strip().lower()

    match = _HEX_COLOR_RE.match(value)
    if match:
        hex_value = match.group("hex")
        if len(hex_value) == 3:
            hex_value = "".join(char * 2 for char in hex_value)
        return RGBColor.from_string(hex_value.upper())

    match = _RGB_COLOR_RE.match(value)
    if match:
        red = _clamp_rgb(int(match.group("r")))
        green = _clamp_rgb(int(match.group("g")))
        blue = _clamp_rgb(int(match.group("b")))
        return RGBColor(red, green, blue)

    named = _NAMED_COLORS.get(value)
    if named:
        return RGBColor.from_string(named)

    raise ValueError(f"Unsupported color value '{color_value}'")


def parse_css_font_size(size_value: str) -> Pt:
    """Parse CSS font-size supporting pt and px."""
    match = _FONT_SIZE_RE.match(size_value)
    if not match:
        raise ValueError(f"Unsupported font-size '{size_value}'")

    value = float(match.group("value"))
    unit = match.group("unit").lower()

    if unit == "pt":
        return Pt(value)
    if unit == "px":
        return Pt(value * 0.75)
    raise ValueError(f"Unsupported font-size unit '{unit}'")


def _clamp_rgb(value: int) -> int:
    return max(0, min(255, value))
