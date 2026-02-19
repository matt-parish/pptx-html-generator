from pptx.dml.color import RGBColor
from pptx.util import Cm, Emu, Inches, Pt

from pptx_html_generator.styles import (
    parse_css_color,
    parse_css_font_size,
    parse_css_inline_style,
    parse_dimension,
)


def test_parse_dimension_inches():
    assert parse_dimension("1in") == Inches(1)


def test_parse_dimension_centimeters():
    assert parse_dimension("2.5cm") == Cm(2.5)


def test_parse_dimension_points():
    assert parse_dimension("18pt") == Pt(18)


def test_parse_dimension_emu():
    assert parse_dimension("914400emu") == Emu(914400)


def test_parse_dimension_invalid_unit():
    try:
        parse_dimension("12px")
    except ValueError as exc:
        assert "Invalid dimension" in str(exc)
    else:
        raise AssertionError("Expected ValueError")


def test_parse_css_inline_style():
    result = parse_css_inline_style("color: #f00; font-size: 24pt;")
    assert result == {"color": "#f00", "font-size": "24pt"}


def test_parse_css_color_hex_and_rgb_and_named():
    assert parse_css_color("#F00") == RGBColor(255, 0, 0)
    assert parse_css_color("rgb(255, 0, 0)") == RGBColor(255, 0, 0)
    assert parse_css_color("red") == RGBColor(255, 0, 0)


def test_parse_css_font_size_pt_and_px():
    assert parse_css_font_size("24pt") == Pt(24)
    assert parse_css_font_size("16px") == Pt(12)
