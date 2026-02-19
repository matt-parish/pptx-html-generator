---
id: 20260219000000
title: "PPTX HTML Generator"
created: 2026-02-19
modified: 2026-02-19
tags:
  - type/project
  - topic/python
  - topic/powerpoint
  - topic/automation
  - topic/html
type: project
status: draft
start_date:
target_date:
completion_date:
---

# PPTX HTML Generator

## Goal

Build a Python library that generates PPTX presentations from a JSON schema, with support for rich text formatting via a subset of HTML tags. The JSON defines slide structure (layouts, text boxes, positions), and text content within those boxes uses HTML for inline/block formatting — enabling callers to write `"<b>Revenue</b> grew <span style=\"color:green\">12%</span>"` and get properly formatted PowerPoint output.

## Success Criteria

- [ ] Define and document the JSON schema for slides, text boxes, and content
- [ ] Implement HTML-to-rich-text parser with all Tier 1 tags
- [ ] Implement Tier 2 tags (strikethrough, code, hyperlinks, spans, sub/sup)
- [ ] Implement Tier 3 block-level tags (lists, headings)
- [ ] Handle all identified edge cases with tests
- [ ] Plain text passthrough (no HTML = no processing overhead)
- [ ] Comprehensive test suite covering every tag and edge case
- [ ] Working CLI or function entry point: `json_in → pptx_out`

## Context & Motivation

Extends the ideas in [[powerpoint-automation-ideas]] from template-based modification to full generation. The HTML-in-JSON approach provides a clean separation: JSON handles structure/layout, HTML handles text formatting. This makes it easy for both humans and LLMs to author slide content without needing to understand python-pptx internals.

## Current Status

**Planning Phase** (2026-02-19) — Research complete, design decisions made, ready to implement.

---

## Architecture

### High-Level Flow

```
JSON input → Validate schema → For each slide:
  → Create slide with layout
  → For each text box:
    → Create shape at position/size
    → Parse HTML content
    → Convert to python-pptx paragraphs/runs with formatting
  → Save .pptx
```

### Dependencies

| Package | Purpose | Version |
|---------|---------|---------|
| `python-pptx` | PPTX generation | ≥1.0.0 |
| `beautifulsoup4` | HTML parsing | ≥4.12 |
| `lxml` | BS4 parser backend (fast, robust) | ≥4.9 |

### Project Structure

```
pptx-html-generator/
├── src/
│   ├── __init__.py
│   ├── generator.py        # Main entry: JSON → PPTX
│   ├── schema.py           # JSON schema definition & validation
│   ├── html_parser.py      # HTML → python-pptx rich text
│   └── styles.py           # Default styles, color parsing, units
├── tests/
│   ├── test_html_parser.py # Tag-by-tag + edge cases
│   ├── test_generator.py   # End-to-end JSON → PPTX
│   ├── test_schema.py      # Schema validation
│   └── test_styles.py      # Color/unit parsing
├── examples/
│   ├── simple.json         # Minimal example
│   ├── rich_text.json      # All HTML tags demonstrated
│   └── edge_cases.json     # Weird scenarios
├── pyproject.toml
└── README.md
```

---

## JSON Schema Design

### Top Level

```json
{
  "presentation": {
    "width": "13.333in",
    "height": "7.5in",
    "defaults": {
      "font_name": "Calibri",
      "font_size": "11pt",
      "font_color": "#333333"
    }
  },
  "slides": [...]
}
```

### Slide

```json
{
  "layout": "blank",
  "background": "#FFFFFF",
  "elements": [...]
}
```

`layout` values: `"blank"`, `"title"`, `"title_and_content"`, `"section_header"`, or an integer index into the slide master's layouts.

### Element (Text Box)

```json
{
  "type": "textbox",
  "position": { "left": "1in", "top": "1in", "width": "8in", "height": "2in" },
  "content": "<p>Hello <b>world</b></p>",
  "style": {
    "font_name": "Arial",
    "font_size": "18pt",
    "font_color": "#000000",
    "alignment": "center",
    "vertical_anchor": "middle",
    "word_wrap": true
  }
}
```

### Unit Format

All dimensions use string values with unit suffixes: `"1in"`, `"2.5cm"`, `"72pt"`, `"914400emu"`. Parsed into python-pptx `Inches()`, `Cm()`, `Pt()`, or `Emu()` respectively.

---

## HTML Tag Support

### Tier 1 — Core Formatting (Must Have)

These are the minimum viable tags. Without these, the library isn't useful.

| HTML | python-pptx | Rationale |
|------|-------------|-----------|
| `<b>`, `<strong>` | `run.font.bold = True` | Most fundamental emphasis |
| `<i>`, `<em>` | `run.font.italic = True` | Second most common emphasis |
| `<u>` | `run.font.underline = True` | Common in presentations |
| `<br>` / `<br/>` | New line within paragraph (soft return via `\n` or `\v`) | Essential for line control within a text box |
| `<p>` | New paragraph (`text_frame.add_paragraph()`) | Fundamental block structure |
| Plain text | Direct run with inherited styles | Must handle content with zero HTML tags |

**Design note on `<br>` vs `<p>`**: A `<br>` produces a line break *within* the current paragraph (same spacing/bullet level). A `<p>` creates a *new paragraph* (which can have different spacing, alignment, bullet level). This distinction matters for PowerPoint output. In python-pptx, a soft return is a vertical tab character `\v` (not `\n`).

### Tier 2 — Extended Formatting (Should Have)

These add significant value and cover most real-world use cases.

| HTML | python-pptx | Rationale |
|------|-------------|-----------|
| `<s>`, `<del>`, `<strike>` | `run.font._element.attrib['strike'] = 'sngStrike'` | Useful but needs OOXML workaround — no public API in python-pptx (PR #606 pending) |
| `<code>` | `run.font.name = 'Courier New'` | Technical/code content common in decks |
| `<a href="...">` | `run.hyperlink.address = url` | Links are very common in presentations |
| `<span style="...">` | Parse inline CSS → apply to run | The flexible escape hatch for color, size, font changes |
| `<sup>` | `run.font._element.set('baseline', '30000')` | Footnotes, ordinals (1st, 2nd) — needs OOXML workaround |
| `<sub>` | `run.font._element.set('baseline', '-25000')` | Chemical formulas, footnotes — needs OOXML workaround |

**Supported inline CSS properties** (via `style="..."` attribute on any tag):

| CSS Property | python-pptx Mapping |
|-------------|---------------------|
| `color` | `run.font.color.rgb = RGBColor(...)` |
| `font-size` | `run.font.size = Pt(...)` |
| `font-family` | `run.font.name = ...` |
| `font-weight: bold` | `run.font.bold = True` |
| `font-style: italic` | `run.font.italic = True` |
| `text-decoration: underline` | `run.font.underline = True` |
| `text-decoration: line-through` | Strikethrough via OOXML workaround |
| `background-color` | **Not supported** — PPTX runs don't have background highlight (unlike Word). Explicitly unsupported. |

**Supported color formats** (for `color` and any color CSS property):

| Format | Example | Notes |
|--------|---------|-------|
| 6-digit hex | `#FF0000` | Primary format |
| 3-digit hex | `#F00` | Expand to `#FF0000` |
| `rgb()` | `rgb(255, 0, 0)` | Parse components |
| Named colors | `red`, `blue`, `green` | Map a set of ~17 CSS named colors (the standard HTML color keywords) |

### Tier 3 — Block-Level Elements (Nice to Have)

These handle structured content within a text box.

| HTML | python-pptx | Rationale |
|------|-------------|-----------|
| `<ul>` / `<li>` | Paragraphs with `p.level` set + bullet character | Bullet lists are extremely common in presentations |
| `<ol>` / `<li>` | Paragraphs with `p.level` set + numbering | Numbered lists less common but still useful |
| `<h1>` – `<h6>` | Paragraph with scaled font size + bold | Useful for structure within large text boxes |

**List implementation details**:
- Each `<li>` becomes a new paragraph
- `<ul>` nesting increments `p.level` (0, 1, 2...)
- Bullet character: `•` at level 0, `–` at level 1, `◦` at level 2
- `<ol>` numbering: managed via paragraph numbering format (if available) or text prefix fallback (`1.`, `2.`, etc.)
- Rich text *within* `<li>` elements is fully supported (e.g., `<li>Item with <b>bold</b></li>`)

### Explicitly Not Supported

| HTML | Why Not |
|------|---------|
| `<table>` | Tables are a separate shape type in PPTX, not text runs. Would need its own element type in the JSON schema. Out of scope for v1. |
| `<img>` | Images are separate shapes. Would need its own element type. Out of scope for v1. |
| `<div>` | Treat as transparent container — just process children. No formatting effect. |
| `<style>` | Only inline `style=""` attributes supported. No CSS rule processing. |
| `<script>` | No use case. |
| `background-color` CSS | PPTX text runs don't support character-level highlighting the way Word does. The OOXML `<a:highlight>` element exists only in newer specs and isn't supported by python-pptx. |

---

## HTML Parser Design

### Core Algorithm: Recursive Style Accumulation

The parser walks the HTML tree recursively, accumulating formatting state as it descends into nested tags, and creating python-pptx runs only when it reaches text nodes (NavigableString).

```
function process_node(node, paragraph, accumulated_styles):
    if node is NavigableString:
        if node.text is not empty/whitespace-only:
            run = paragraph.add_run()
            run.text = node.text
            apply_styles(run, accumulated_styles)
        return

    # It's a Tag — merge this tag's formatting into accumulated styles
    new_styles = accumulated_styles.copy()
    merge_tag_styles(node.tag_name, node.attrs, new_styles)

    if node is block-level (<p>, <li>, <h1-h6>):
        paragraph = create_or_advance_paragraph(...)

    for child in node.children:
        process_node(child, paragraph, new_styles)
```

### Style Accumulation Dictionary

```python
{
    "bold": True,          # from <b>/<strong>/font-weight
    "italic": False,       # from <i>/<em>/font-style
    "underline": False,    # from <u>/text-decoration
    "strikethrough": False, # from <s>/<del>/text-decoration
    "font_name": None,     # from <code>/font-family (None = inherit)
    "font_size": None,     # from font-size CSS (None = inherit)
    "font_color": None,    # from color CSS (None = inherit)
    "superscript": False,  # from <sup>
    "subscript": False,    # from <sub>
    "hyperlink": None,     # from <a href="...">
}
```

### TextFrame Initialisation Gotcha

python-pptx TextFrames always come with one default empty paragraph. The parser must use `text_frame.paragraphs[0]` for the first paragraph, then `text_frame.add_paragraph()` for subsequent ones. Failing to account for this creates a blank line at the top.

---

## Edge Cases & Weird Scenarios

These are the scenarios that will break a naive implementation. Each needs explicit handling and a test case.

### 1. Plain Text (No HTML)

**Input**: `"Revenue grew 12% this quarter"`
**Expected**: Single run, no formatting, just the text.
**Strategy**: Check if content contains any `<` characters. If not, skip HTML parsing entirely and insert as plain text. This is both a performance optimisation and a correctness guarantee.

### 2. Empty or Null Content

**Input**: `""`, `null`, or missing `content` field
**Expected**: Empty text frame (keep the shape, just no text).
**Strategy**: Early return. Don't call the parser.

### 3. Nested Duplicate Tags

**Input**: `"<b>bold <b>still bold</b> bold</b>"`
**Expected**: All text is bold. The inner `<b>` is a no-op.
**Strategy**: Style accumulation is naturally idempotent — setting `bold: True` when it's already `True` does nothing.

### 4. Overlapping/Interleaved Formatting

**Input**: `"<b>bold <i>bold-italic</i> bold</b> <i>italic</i>"`
**Expected**: Four runs: "bold " (B), "bold-italic" (B+I), " bold" (B), " italic" (I).
**Strategy**: The recursive style accumulation handles this naturally because each branch of the tree carries its own copy of the style state.

### 5. Whitespace Between Inline Elements

**Input**: `"<b>hello</b> <i>world</i>"`
**Expected**: "hello" (bold) + " " (plain) + "world" (italic).
**Gotcha**: The space between `</b>` and `<i>` is a NavigableString `" "` at the parent level. The parser must not strip/collapse whitespace-only text nodes between inline elements.
**Strategy**: Only skip truly empty strings (`""`), not whitespace strings. Whitespace-only NavigableStrings between inline elements produce runs with just spaces.

### 6. Leading/Trailing Whitespace in Tags

**Input**: `"<b> hello </b>"`
**Expected**: " hello " (bold, preserving spaces).
**Strategy**: Don't strip text content. Preserve exactly what's in the HTML.

### 7. Whitespace Normalisation

**Input**: `"<p>  Multiple   spaces   here  </p>"`
**Decision**: **Preserve literal whitespace**. Unlike browsers, PowerPoint doesn't collapse whitespace. What you put in is what you get. If the caller wants collapsed spaces, they should normalise before passing to us.
**Alternative considered**: Collapsing whitespace like a browser. Rejected because PPTX is not a browser and callers would find it surprising if their carefully placed spaces disappeared.

### 8. `<br>` Handling

**Input**: `"Line one<br>Line two<br/>Line three<br />Line four"`
**Expected**: All four lines in the same paragraph, separated by soft returns.
**Strategy**: When encountering a `<br>` tag, insert `\v` (vertical tab) into the current run's text. This produces a soft return in PowerPoint. All three `<br>` variants (`<br>`, `<br/>`, `<br />`) are handled identically by BeautifulSoup.

### 9. `<p>` Creates New Paragraph, Not Nested

**Input**: `"<p>First paragraph</p><p>Second paragraph</p>"`
**Expected**: Two separate paragraphs in the text frame.
**Gotcha**: What about `<p>outer <p>inner</p></p>`? This is malformed HTML. BeautifulSoup will auto-close the outer `<p>` before the inner one, producing two sibling paragraphs. That's fine — we follow BeautifulSoup's interpretation.

### 10. Mixed Block and Inline at Top Level

**Input**: `"Some text <b>bold</b> <p>A paragraph</p> more text"`
**Expected**: The bare text and inline elements before the `<p>` go into the first (default) paragraph. The `<p>` creates a new paragraph. The trailing text after `</p>` creates yet another paragraph.
**Strategy**: Track current paragraph. When a block-level element is encountered, advance to a new paragraph. After the block element closes, any further content goes to a new paragraph.

### 11. Deeply Nested Tags

**Input**: `"<b><i><u><span style=\"color:red\"><a href=\"...\">text</a></span></u></i></b>"`
**Expected**: Single run with bold + italic + underline + red color + hyperlink.
**Strategy**: The recursive approach handles arbitrary depth naturally. No special case needed. The style dict just accumulates more entries.

### 12. HTML Entities

**Input**: `"5 &gt; 3 &amp; 2 &lt; 4, also &nbsp; non-breaking"`
**Expected**: `5 > 3 & 2 < 4, also   non-breaking` (with a non-breaking space).
**Strategy**: BeautifulSoup automatically decodes all HTML entities. `&nbsp;` becomes `\xa0` (U+00A0). Preserve it as-is — PowerPoint renders it as a non-breaking space.

### 13. Escaped HTML (Literal Tags in Output)

**Input**: `"Use &lt;b&gt; for bold"`
**Expected**: `Use <b> for bold` rendered as plain text (not actually bold).
**Strategy**: BeautifulSoup decodes entities before parsing. Since `&lt;b&gt;` becomes the text `<b>` *after* entity decoding during the first parse, it won't be re-parsed as a tag. This just works.

### 14. Unknown/Unsupported Tags

**Input**: `"<blink>some text</blink>"` or `"<div class=\"foo\">content</div>"`
**Expected**: Process children, ignore the tag itself. No error.
**Strategy**: The default case in the tag handler is "process children with no style changes". Unknown tags are transparent wrappers.

### 15. `style` Attribute on Non-Span Tags

**Input**: `"<b style=\"color: blue\">bold blue</b>"`
**Expected**: Bold and blue.
**Strategy**: Parse inline `style` attribute on *any* tag, not just `<span>`. The tag's own formatting (bold from `<b>`) merges with the CSS formatting (color from `style`).

### 16. Conflicting CSS and Tag Formatting

**Input**: `"<b style=\"font-weight: normal\">not bold?</b>"`
**Expected**: Not bold. Inline CSS overrides the tag's implicit formatting.
**Strategy**: Process tag-level formatting first, then CSS overrides. CSS wins (this matches browser behavior).

### 17. Empty Tags

**Input**: `"Before<b></b>After"` or `"<p></p>"`
**Expected**: For inline: "BeforeAfter" with no extra run. For block: an empty paragraph.
**Strategy**: Empty inline tags produce no runs (no text nodes inside). Empty `<p>` produces an empty paragraph (which is valid — it's a blank line).

### 18. Unicode, Emoji, and Special Characters

**Input**: `"<b>café</b> 🚀 <i>naïve résumé</i>"`
**Expected**: All characters preserved exactly. Emoji rendered by PowerPoint's font fallback.
**Strategy**: python-pptx handles Unicode natively. No special handling needed. Just pass the text through.

### 19. Very Large Content

**Input**: A text box with thousands of words and hundreds of tags.
**Expected**: Works correctly, perhaps slowly.
**Strategy**: BeautifulSoup + lxml is fast enough for reasonable content. No hard limit, but document that extremely large HTML may be slow. Consider if this ever becomes a real issue (it probably won't for presentation content).

### 20. Self-Closing Tags in Wrong Places

**Input**: `"<b/>text"` (self-closing bold)
**Expected**: "text" is not bold — `<b/>` is treated as an empty element.
**Strategy**: BeautifulSoup handles this. `<b/>` creates a Tag with no children. The "text" after it is a sibling NavigableString, not a child, so it gets no bold formatting. This is correct HTML behavior.

### 21. Multiple `style` Properties

**Input**: `"<span style=\"color: red; font-size: 24pt; font-family: Arial\">styled</span>"`
**Expected**: Red, 24pt, Arial.
**Strategy**: Split `style` value by `;`, then each property by `:`. Trim whitespace. Map each to the appropriate python-pptx property.

### 22. Colour Edge Cases

**Input**: `"<span style=\"color: #F00\">short hex</span>"` or `"<span style=\"color: rgb(255,0,0)\">rgb</span>"`
**Expected**: Both produce red text.
**Strategy**: The color parser must handle:
- `#RRGGBB` → direct mapping
- `#RGB` → expand each digit (`#F00` → `#FF0000`)
- `rgb(r, g, b)` → extract integers, clamp to 0-255
- Named colors → lookup table (17 standard HTML colors + any extras we want)

### 23. Lists with Rich Text Items

**Input**: `"<ul><li>Item with <b>bold</b> and <a href=\"...\">a link</a></li><li>Plain item</li></ul>"`
**Expected**: Two bullet paragraphs. First has mixed formatting runs. Second is plain.
**Strategy**: `<li>` creates a new paragraph with bullet properties. Content within `<li>` is processed normally with style accumulation.

### 24. Nested Lists

**Input**: `"<ul><li>Level 0<ul><li>Level 1<ul><li>Level 2</li></ul></li></ul></li></ul>"`
**Expected**: Three bullet paragraphs with increasing `p.level` (0, 1, 2).
**Strategy**: Track nesting depth of `<ul>`/`<ol>`. Each nesting level increments `p.level` for contained `<li>` elements.

---

## Implementation Order

### Phase 1: Skeleton + Plain Text

1. Set up project structure (`pyproject.toml`, `src/`, `tests/`)
2. Define JSON schema in `schema.py` (validation with clear error messages)
3. Implement `generator.py` — create presentation, slides, text boxes from JSON
4. Plain text only (no HTML parsing yet) — `content` string goes directly into a single run
5. Unit parsing (`"1in"` → `Inches(1)`, `"18pt"` → `Pt(18)`)
6. Tests for schema validation + basic generation

### Phase 2: HTML Parser — Tier 1 Tags

1. Implement `html_parser.py` with recursive style accumulation
2. Tags: `<b>`, `<strong>`, `<i>`, `<em>`, `<u>`, `<br>`, `<p>`, plain text passthrough
3. Handle the TextFrame first-paragraph gotcha
4. Tests for each tag individually + combinations + edge cases 1-10

### Phase 3: Tier 2 Tags

1. Add: `<s>`/`<del>`, `<code>`, `<a>`, `<span style>`, `<sup>`, `<sub>`
2. Implement CSS inline style parser in `styles.py`
3. Color parser (hex, rgb, named)
4. Font size parser (pt, px with conversion)
5. Tests for each + edge cases 11-22

### Phase 4: Tier 3 Tags + Polish

1. Add: `<ul>`, `<ol>`, `<li>`, `<h1>`–`<h6>`
2. List nesting logic
3. Tests for edge cases 23-24
4. End-to-end examples
5. Error handling (malformed JSON, invalid units, bad color values)

---

## Existing Related Work

### In This Knowledge Base
- [[powerpoint-automation-ideas]] — earlier exploration of PowerPoint automation with Python

### External Libraries Surveyed
- **python-pptx** (scanny) — the foundation library we'll build on. Mature, well-documented, v1.0+
- **json-to-ppt** (jsonforge) — declarative JSON-to-PPTX. Interesting schema ideas but different approach (no HTML)
- **pptx-template** (m3dev) — template-based generation from JSON. Good for fill-in-the-blanks, not for custom content
- **html2pptx** (maximecaruchet) — CSS-selector-based extraction from HTML pages. Different goal (web scraping → slides)
- **Aspose.Slides** — commercial, has `add_from_html()` but $2000+/year

### Key Technical References
- python-pptx text API: runs, paragraphs, fonts, hyperlinks
- BeautifulSoup: `.children` iterator, NavigableString vs Tag, entity decoding
- OOXML: `<a:rPr>` for run properties, `<a:pPr>` for paragraph properties

---

## Verified Technical Facts

These were confirmed against documentation (not assumed):

- `run.font.bold/italic/underline` — ✅ work, accept `True`/`False`/`None` (None = inherit)
- `run.font.color.rgb = RGBColor(r,g,b)` — ✅ must use `RGBColor` class, not tuples
- `run.font.size = Pt(n)` — ✅ must use `Pt()`, not bare integers
- `run.font.name = 'str'` — ✅ string font family name
- `run.hyperlink.address = url` — ✅ set via property, no `add_hyperlink()` method
- Strikethrough — ❌ no public API. Use `run.font._element.attrib['strike'] = 'sngStrike'`
- Superscript/subscript — ❌ no public API. Use `run.font._element.set('baseline', '30000'/'-25000')`
- `text_frame.paragraphs[0]` — always exists (default empty paragraph)
- `text_frame.add_paragraph()` — appends after existing content
- `\v` (vertical tab) — produces soft return (line break within paragraph) in PPTX
- BeautifulSoup `.children` — yields direct children only (NavigableString + Tag)
- BeautifulSoup auto-decodes all HTML entities (`&amp;` → `&`, `&nbsp;` → `\xa0`)
- BeautifulSoup handles malformed HTML (auto-closes tags, etc.)

---

## Open Questions / Decisions for Implementation Time

1. **Slide layout handling**: Use built-in layouts from the default template, or always use blank + manual positioning? Leaning toward both: support `"blank"` for full control and named layouts for convenience.
2. **Default font inheritance**: When no explicit font is set in JSON or HTML, what does python-pptx default to? Need to test whether the theme font comes through or if we need to set an explicit default.
3. **Error handling strategy**: Strict (throw on invalid HTML/JSON) vs permissive (best-effort, log warnings)? Leaning toward permissive for HTML (follow BeautifulSoup's lead) and strict for JSON schema.
4. **Image/table support**: Explicitly out of scope for v1, but the JSON schema should be extensible so `"type": "image"` and `"type": "table"` elements can be added later without breaking changes.

---

## Next Actions

### Immediate
- [ ] Create project directory and `pyproject.toml`
- [ ] Implement Phase 1 (skeleton + plain text + schema)
- [ ] Write first tests

### Then
- [ ] Phase 2: Tier 1 HTML tags
- [ ] Phase 3: Tier 2 HTML tags + style parsing
- [ ] Phase 4: Tier 3 tags + polish

---

## Notes

- The `\v` soft return is the key insight for `<br>` handling — `\n` in python-pptx text creates a new paragraph, not a line break within one
- Strikethrough and super/subscript require reaching into the OOXML XML layer — these are the only tags that need private API access
- The recursive style accumulation pattern is well-suited to how BeautifulSoup represents the DOM — `.children` gives you exactly the interleaved text+tag nodes you need
- CSS `background-color` is explicitly unsupported because PPTX text runs genuinely don't support character-level highlighting (unlike Word's `<w:highlight>`)
