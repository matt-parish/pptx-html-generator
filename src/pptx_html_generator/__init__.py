"""PPTX HTML generator package."""

from .html_parser import render_html_to_text_frame
from .generator import generate_pptx
from .reverse import extract_shape_html, list_slide_elements

__all__ = [
    "render_html_to_text_frame",
    "generate_pptx",
    "extract_shape_html",
    "list_slide_elements",
]
