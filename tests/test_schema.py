import pytest

from pptx_html_generator.schema import validate_spec


def valid_spec():
    return {
        "presentation": {"width": "10in", "height": "7.5in"},
        "slides": [
            {
                "layout": "blank",
                "elements": [
                    {
                        "type": "textbox",
                        "position": {
                            "left": "1in",
                            "top": "1in",
                            "width": "5in",
                            "height": "2in",
                        },
                        "content": "Hello",
                    }
                ],
            }
        ],
    }


def test_validate_spec_accepts_valid_spec():
    spec = valid_spec()
    assert validate_spec(spec) is spec


def test_validate_spec_requires_slides_list():
    with pytest.raises(ValueError, match="slides"):
        validate_spec({"presentation": {}})


def test_validate_spec_rejects_unknown_element_type():
    spec = valid_spec()
    spec["slides"][0]["elements"][0]["type"] = "image"
    with pytest.raises(ValueError, match="textbox"):
        validate_spec(spec)


def test_validate_spec_rejects_invalid_dimension():
    spec = valid_spec()
    spec["slides"][0]["elements"][0]["position"]["left"] = "100px"
    with pytest.raises(ValueError, match="Invalid dimension"):
        validate_spec(spec)


def test_validate_spec_rejects_invalid_style_alignment():
    spec = valid_spec()
    spec["slides"][0]["elements"][0]["style"] = {"alignment": "diagonal"}
    with pytest.raises(ValueError, match="alignment"):
        validate_spec(spec)
