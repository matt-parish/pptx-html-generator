"""Command-line entrypoint."""

from __future__ import annotations

import argparse
import json

from .generator import generate_pptx_from_file
from .reverse import extract_shape_html, list_slide_elements


def main() -> int:
    parser = argparse.ArgumentParser(description="Generate PPTX from JSON spec.")
    subparsers = parser.add_subparsers(dest="command")

    generate_parser = subparsers.add_parser("generate", help="Generate PPTX from JSON")
    generate_parser.add_argument("json_in", help="Path to input JSON file")
    generate_parser.add_argument("pptx_out", help="Path to output PPTX file")

    list_parser = subparsers.add_parser(
        "list-elements", help="List slide elements using Selection Pane names"
    )
    list_parser.add_argument("pptx_in", help="Path to input PPTX file")
    list_parser.add_argument("--slide", type=int, required=True, help="1-based slide number")

    extract_parser = subparsers.add_parser(
        "extract-html", help="Extract HTML from a selected shape"
    )
    extract_parser.add_argument("pptx_in", help="Path to input PPTX file")
    extract_parser.add_argument("--slide", type=int, required=True, help="1-based slide number")
    extract_group = extract_parser.add_mutually_exclusive_group(required=True)
    extract_group.add_argument(
        "--shape-name", help="Selection Pane shape name (matches PowerPoint selection pane)"
    )
    extract_group.add_argument("--shape-id", type=int, help="Shape id")
    extract_parser.add_argument("--out", help="Optional output HTML file path")

    parser.add_argument("legacy", nargs="*", help=argparse.SUPPRESS)
    args = parser.parse_args()

    # Backward compatibility: `pptx-html-generator input.json out.pptx`
    if args.command is None and len(args.legacy) == 2:
        generate_pptx_from_file(args.legacy[0], args.legacy[1])
        return 0

    if args.command == "generate":
        generate_pptx_from_file(args.json_in, args.pptx_out)
        return 0
    if args.command == "list-elements":
        items = list_slide_elements(args.pptx_in, args.slide)
        print(json.dumps(items, indent=2))
        return 0
    if args.command == "extract-html":
        html = extract_shape_html(
            args.pptx_in,
            args.slide,
            shape_name=args.shape_name,
            shape_id=args.shape_id,
        )
        if args.out:
            with open(args.out, "w", encoding="utf-8") as handle:
                handle.write(html)
        else:
            print(html)
        return 0

    parser.print_help()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
