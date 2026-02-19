# Publishing Guide

This repo supports publishing to both GitHub and PyPI.

## 1) Verify local state

```bash
. .venv/bin/activate
pytest -q
python -m pip install -e ".[release]"
python -m build
python -m twine check dist/*
```

## 2) Commit and tag for Git

```bash
git add .
git commit -m "Release v0.1.0"
git tag v0.1.0
git push origin <branch>
git push origin v0.1.0
```

## 3) Publish to PyPI

Set credentials via token:

```bash
export TWINE_USERNAME="__token__"
export TWINE_PASSWORD="<pypi-token>"
```

Upload:

```bash
python -m twine upload dist/*
```

## 4) Verify install from PyPI

```bash
python -m venv /tmp/pptx-html-generator-smoke
. /tmp/pptx-html-generator-smoke/bin/activate
python -m pip install --upgrade pip
python -m pip install pptx-html-generator
python -c "from pptx_html_generator import render_html_to_text_frame; print('ok')"
```
