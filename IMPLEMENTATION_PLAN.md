# Implementation Plan

## Baseline Delivered

- Git repository initialized.
- Phase 1 starter package scaffolded.
- JSON schema validation implemented for presentation/slides/textbox elements.
- Dimension parsing implemented for `in`, `cm`, `pt`, `emu`.
- Plain text generation path implemented (`json -> pptx`).
- CLI entrypoint added.
- Baseline tests added for schema, unit parsing, and end-to-end plain text generation.

## Next Implementation Steps

1. Phase 2: Tier 1 HTML parser
- Add `html_parser.py` with recursive style accumulation.
- Implement `<b>`, `<strong>`, `<i>`, `<em>`, `<u>`, `<br>`, `<p>`.
- Integrate parser into `generator.py`.
- Add focused tests for tag behavior and paragraph/soft-return semantics.

2. Phase 3: Tier 2 formatting
- Add inline style parser for `style=""` attributes.
- Implement color parsing: hex, short hex, rgb, named colors.
- Add `<s>/<del>/<strike>`, `<code>`, `<a>`, `<span>`, `<sup>`, `<sub>`.
- Add tests for private OOXML workarounds and hyperlink behavior.

3. Phase 4: Tier 3 block elements
- Implement `<ul>`, `<ol>`, `<li>`, `<h1>`-`<h6>`.
- Add nested list depth handling and numbering fallback.
- Add tests for nested/mixed content and block transitions.

4. Hardening
- Improve error messages and strict/permissive boundaries.
- Add examples for rich text and edge cases.
- Add CI job for lint + tests.
