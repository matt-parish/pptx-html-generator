# Public API (v1)

This document defines the stable `v1` API for `pptx-html-generator`.

## Design goal

`render_html_to_text_frame(...)` is the first-class API.

Most users should integrate at the content-rendering boundary of their existing
`python-pptx` pipeline, where strings are assigned to a text frame.

## Stable imports

```python
from pptx_html_generator.api import (
    generate_pptx,
    generate_pptx_from_file,
    render_html_to_text_frame,
    list_slide_elements,
    extract_shape_html,
    install_html_extension,
    uninstall_html_extension,
)
```

Top-level package exports are also stable:

```python
from pptx_html_generator import (
    generate_pptx,
    render_html_to_text_frame,
    list_slide_elements,
    extract_shape_html,
    install_html_extension,
    uninstall_html_extension,
)
```

## First-class API

### `render_html_to_text_frame(text_frame, content, base_styles=None) -> None`

Render HTML content into an existing `python-pptx` `TextFrame`.

Use this for drop-in integration when your app already builds slides/shapes.

Typical usage:

```python
from pptx_html_generator import render_html_to_text_frame

render_html_to_text_frame(shape.text_frame, html_string, base_styles={...})
```

## Secondary APIs

### `generate_pptx(spec, output_path) -> pathlib.Path`

Generate a full PPTX from JSON-like spec.

### `generate_pptx_from_file(input_path, output_path) -> pathlib.Path`

File-based helper wrapper around `generate_pptx`.

## Reverse API

### `list_slide_elements(pptx_path, slide_number) -> list[dict]`

Return selectable shape metadata for a slide:
- `shape_id`
- `name` (PowerPoint Selection Pane name)
- `has_text`
- `text_preview`

### `extract_shape_html(pptx_path, slide_number, shape_name=None, shape_id=None) -> str`

Extract HTML from a text-bearing shape selected by:
- `shape_name` (recommended; matches Selection Pane), or
- `shape_id`

## Extension API

### `install_html_extension() -> None`

Install `TextFrame.html` monkey-patch on `python-pptx` text frames.

### `uninstall_html_extension() -> None`

Remove `TextFrame.html` monkey-patch (primarily for test isolation).

When the package is imported from `pptx_html_generator`, the extension is
installed automatically.

## Non-goals for v1 stability

These are intentionally internal and may change:
- `schema.py` helper internals
- `styles.py` parser internals
- `html_parser.py` recursive traversal internals
- exact ordering/normalization of generated reverse HTML tags
