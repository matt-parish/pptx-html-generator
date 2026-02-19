from pathlib import Path

from pptx_html_generator.generator import generate_pptx
from pptx_html_generator.reverse import extract_shape_html, list_slide_elements


def _build_demo(tmp_path: Path) -> Path:
    spec = {
        "presentation": {"width": "10in", "height": "7.5in"},
        "slides": [
            {
                "layout": "blank",
                "elements": [
                    {
                        "type": "textbox",
                        "name": "BodyContent",
                        "position": {
                            "left": "1in",
                            "top": "1in",
                            "width": "8in",
                            "height": "4in",
                        },
                        "content": "<h2>Heading</h2><p>Hello <b>bold</b> <span style=\"color:#008000\">green</span> <a href=\"https://example.com\">link</a></p><p>Line A<br/>Line B</p><ul><li>One</li><li>Two</li></ul><ol><li>First</li><li>Second</li></ol>",
                    }
                ],
            }
        ],
    }
    output = tmp_path / "reverse-demo.pptx"
    generate_pptx(spec, output)
    return output


def test_list_elements_uses_shape_names(tmp_path: Path):
    output = _build_demo(tmp_path)
    items = list_slide_elements(output, 1)
    names = [item["name"] for item in items]
    assert "BodyContent" in names


def test_extract_shape_html_by_name(tmp_path: Path):
    output = _build_demo(tmp_path)
    result = extract_shape_html(output, 1, shape_name="BodyContent")

    assert "<h2>" in result and "Heading" in result and "</h2>" in result
    assert "<b>bold</b>" in result
    assert "color:#008000" in result
    assert '<a href="https://example.com">' in result
    assert "<br/>" in result
    assert "<ul><li>One</li><li>Two</li></ul>" in result
    assert "<ol><li>First</li><li>Second</li></ol>" in result
