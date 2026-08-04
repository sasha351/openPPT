#!/usr/bin/env python3
"""Convert an HTML outline to a .pptx directly, no Open WebUI involved.

For when the "Export to PowerPoint" button can't reliably read the chat (e.g.
a locked-down Open WebUI deployment). Copy the model's outline out of the chat
into a text file, then run this against it.

Usage:
    python outline_to_pptx.py outline.html
    python outline_to_pptx.py outline.html -o deck.pptx -t brand.pptx
    pbpaste | python outline_to_pptx.py -    # read the outline from stdin

See README.md's "Outline format" section for the HTML grammar.
"""
import argparse
import sys

from openppt_action import build_pptx, parse_outline


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("outline", help="path to an HTML outline file, or - to read stdin")
    ap.add_argument("-o", "--output", help="output .pptx path (default: <outline>.pptx)")
    ap.add_argument("-t", "--template", help=".pptx/.potx to inherit theme, fonts, and layouts from")
    args = ap.parse_args()

    if args.outline == "-":
        text = sys.stdin.read()
        out = args.output or "deck.pptx"
    else:
        with open(args.outline, encoding="utf-8") as f:
            text = f.read()
        out = args.output or args.outline.rsplit(".", 1)[0] + ".pptx"

    slides = parse_outline(text)
    if not slides:
        sys.exit(
            "no slide-shaped content found — expected a <slide title=\"...\"> "
            "element per slide, with <li> bullets inside it"
        )

    with open(out, "wb") as f:
        f.write(build_pptx(slides, template=args.template))
    print(f"wrote {out} ({len(slides)} slides)")


if __name__ == "__main__":
    main()
