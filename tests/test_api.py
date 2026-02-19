from pptx_html_generator import render_html_to_text_frame
from pptx_html_generator.api import render_html_to_text_frame as api_render


def test_primary_api_imports_are_stable():
    assert callable(render_html_to_text_frame)
    assert callable(api_render)
