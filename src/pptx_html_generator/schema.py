"""JSON schema validation for presentation input."""

from __future__ import annotations

from typing import Any

from .styles import parse_css_color, parse_css_font_size, parse_dimension

VALID_LAYOUT_NAMES = {"blank", "title", "title_and_content", "section_header"}
VALID_ALIGNMENT = {"left", "center", "right", "justify"}
VALID_VERTICAL_ANCHOR = {"top", "middle", "bottom"}


def validate_spec(spec: Any) -> dict[str, Any]:
    """Validate incoming spec and return it when valid."""
    if not isinstance(spec, dict):
        raise ValueError("Root JSON must be an object.")

    presentation = spec.get("presentation")
    slides = spec.get("slides")

    if not isinstance(presentation, dict):
        raise ValueError("'presentation' is required and must be an object.")
    if not isinstance(slides, list):
        raise ValueError("'slides' is required and must be an array.")

    _validate_presentation(presentation)

    for slide_index, slide in enumerate(slides):
        _validate_slide(slide, slide_index)

    return spec


def _validate_presentation(presentation: dict[str, Any]) -> None:
    for dim_name in ("width", "height"):
        if dim_name in presentation:
            _validate_dimension(presentation[dim_name], f"presentation.{dim_name}")

    defaults = presentation.get("defaults")
    if defaults is not None and not isinstance(defaults, dict):
        raise ValueError("'presentation.defaults' must be an object when provided.")
    if isinstance(defaults, dict):
        _validate_style(defaults, "presentation.defaults")


def _validate_slide(slide: Any, slide_index: int) -> None:
    if not isinstance(slide, dict):
        raise ValueError(f"slides[{slide_index}] must be an object.")

    layout = slide.get("layout", "blank")
    if not (isinstance(layout, int) or layout in VALID_LAYOUT_NAMES):
        raise ValueError(
            f"slides[{slide_index}].layout must be one of {sorted(VALID_LAYOUT_NAMES)} or an integer."
        )

    elements = slide.get("elements")
    if not isinstance(elements, list):
        raise ValueError(f"slides[{slide_index}].elements is required and must be an array.")

    for element_index, element in enumerate(elements):
        _validate_element(element, slide_index, element_index)


def _validate_element(element: Any, slide_index: int, element_index: int) -> None:
    if not isinstance(element, dict):
        raise ValueError(f"slides[{slide_index}].elements[{element_index}] must be an object.")

    element_type = element.get("type")
    if element_type != "textbox":
        raise ValueError(
            f"slides[{slide_index}].elements[{element_index}].type must be 'textbox' for v1."
        )

    position = element.get("position")
    if not isinstance(position, dict):
        raise ValueError(
            f"slides[{slide_index}].elements[{element_index}].position is required and must be an object."
        )

    for key in ("left", "top", "width", "height"):
        if key not in position:
            raise ValueError(
                f"slides[{slide_index}].elements[{element_index}].position.{key} is required."
            )
        _validate_dimension(
            position[key],
            f"slides[{slide_index}].elements[{element_index}].position.{key}",
        )

    content = element.get("content")
    if content is not None and not isinstance(content, str):
        raise ValueError(
            f"slides[{slide_index}].elements[{element_index}].content must be a string or null."
        )

    name = element.get("name")
    if name is not None and not isinstance(name, str):
        raise ValueError(
            f"slides[{slide_index}].elements[{element_index}].name must be a string when provided."
        )

    style = element.get("style")
    if style is not None and not isinstance(style, dict):
        raise ValueError(f"slides[{slide_index}].elements[{element_index}].style must be an object.")
    if isinstance(style, dict):
        _validate_style(style, f"slides[{slide_index}].elements[{element_index}].style")


def _validate_dimension(value: Any, path: str) -> None:
    try:
        parse_dimension(value)
    except ValueError as exc:
        raise ValueError(f"{path}: {exc}") from exc


def _validate_style(style: dict[str, Any], path: str) -> None:
    font_name = style.get("font_name")
    if font_name is not None and not isinstance(font_name, str):
        raise ValueError(f"{path}.font_name must be a string.")

    font_size = style.get("font_size")
    if font_size is not None:
        if not isinstance(font_size, str):
            raise ValueError(f"{path}.font_size must be a string.")
        try:
            parse_css_font_size(font_size)
        except ValueError as exc:
            raise ValueError(f"{path}.font_size: {exc}") from exc

    font_color = style.get("font_color")
    if font_color is not None:
        if not isinstance(font_color, str):
            raise ValueError(f"{path}.font_color must be a string.")
        try:
            parse_css_color(font_color)
        except ValueError as exc:
            raise ValueError(f"{path}.font_color: {exc}") from exc

    alignment = style.get("alignment")
    if alignment is not None:
        if not isinstance(alignment, str):
            raise ValueError(f"{path}.alignment must be a string.")
        if alignment.lower() not in VALID_ALIGNMENT:
            raise ValueError(f"{path}.alignment must be one of {sorted(VALID_ALIGNMENT)}.")

    vertical_anchor = style.get("vertical_anchor")
    if vertical_anchor is not None:
        if not isinstance(vertical_anchor, str):
            raise ValueError(f"{path}.vertical_anchor must be a string.")
        if vertical_anchor.lower() not in VALID_VERTICAL_ANCHOR:
            raise ValueError(
                f"{path}.vertical_anchor must be one of {sorted(VALID_VERTICAL_ANCHOR)}."
            )

    word_wrap = style.get("word_wrap")
    if word_wrap is not None and not isinstance(word_wrap, bool):
        raise ValueError(f"{path}.word_wrap must be a boolean.")
