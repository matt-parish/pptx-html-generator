# PPTX HTML Generator

Python library for rendering rich HTML text into `python-pptx` text frames.

## Primary API

Recommended: use `TextFrame.html`.

```python
import pptx_html_generator  # installs TextFrame.html extension

shape.text_frame.html = "<p><b>Hello</b> world</p>"  # render HTML -> PPTX text frame
html = shape.text_frame.html                         # extract PPTX text frame -> HTML
```

Alternative (explicit function call): `render_html_to_text_frame(...)`.

See:
- API contract: `docs/API.md`

## Features

- JSON schema validation with useful error messages
- Unit parsing for `in`, `cm`, `pt`, `emu`
- Presentation generation with text box elements
- HTML rich-text support:
  - Inline formatting: `<b>`, `<strong>`, `<i>`, `<em>`, `<u>`, `<s>`, `<del>`, `<strike>`, `<code>`, `<sup>`, `<sub>`
  - Links and styled spans: `<a>`, `<span style>`
  - Block structure: `<p>`, `<br>`, `<ul>`, `<ol>`, `<li>`, `<h1>`-`<h6>`
- JSON style/default support:
  - `font_name`, `font_size`, `font_color`
  - `alignment`, `vertical_anchor`, `word_wrap`
- CLI: `json_in -> pptx_out`
- Optional `TextFrame.html` extension for ergonomic get/set HTML on text frames

## Quick start

```bash
python -m pip install pptx-html-generator
```

Quick start: set HTML (`TextFrame.html`)

```python
from pptx import Presentation
import pptx_html_generator  # installs TextFrame.html extension

prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank
shape = slide.shapes.add_textbox(914400, 914400, 7315200, 1828800)
html = "<p><strong>Hello</strong> <em>world</em> with <a href='https://example.com'>a link</a>.</p>"

shape.text_frame.html = html
prs.save("example.pptx")
```

Quick start: get HTML (`TextFrame.html`)

```python
from pptx import Presentation
import pptx_html_generator  # installs TextFrame.html extension

prs = Presentation("example.pptx")
slide = prs.slides[0]
shape = slide.shapes[0]

html = shape.text_frame.html
print(html)
```

Forward generation CLI (optional helper for JSON-driven generation):

```bash
pptx-html-generator generate examples/full_implementation.json output/full_implementation_demo.pptx
```

Alternative usage (`render_html_to_text_frame`):

```python
from pptx import Presentation
from pptx_html_generator import render_html_to_text_frame

prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[6])
shape = slide.shapes.add_textbox(914400, 914400, 7315200, 1828800)
html = "<p><strong>Hello</strong> <em>world</em></p>"

render_html_to_text_frame(shape.text_frame, html)
```

Development setup:

```bash
python -m pip install -e ".[dev]"
pytest
```

## CLI

Generate PPTX:

```bash
pptx-html-generator generate examples/full_implementation.json output/full_implementation_demo.pptx
```

List selectable elements on a slide (uses PowerPoint Selection Pane names):

```bash
pptx-html-generator list-elements output/full_implementation_demo.pptx --slide 1
```

Extract HTML from a selected shape by Selection Pane name:

```bash
pptx-html-generator extract-html output/full_implementation_demo.pptx --slide 1 --shape-name "BodyContent"
```

## JSON shape (optional high-level API)

```json
{
  "presentation": {
    "width": "13.333in",
    "height": "7.5in"
  },
  "slides": [
    {
      "layout": "blank",
      "elements": [
        {
          "type": "textbox",
          "position": {
            "left": "1in",
            "top": "1in",
            "width": "8in",
            "height": "2in"
          },
          "content": "Plain text content"
        }
      ]
    }
  ]
}
```
