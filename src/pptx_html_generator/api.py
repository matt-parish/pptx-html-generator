"""Stable public API surface for pptx_html_generator."""

from .html_parser import render_html_to_text_frame
from .generator import generate_pptx, generate_pptx_from_file
from .reverse import extract_shape_html, list_slide_elements

__all__ = [
    "generate_pptx",
    "generate_pptx_from_file",
    "render_html_to_text_frame",
    "list_slide_elements",
    "extract_shape_html",
]
