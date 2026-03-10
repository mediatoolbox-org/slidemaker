"""
title: Command-line helpers for slidemaker.
"""

from __future__ import annotations

import argparse
from pathlib import Path
from typing import Optional, Sequence

from slidemaker.cli import SlideBuilder


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="slidemaker")
    subparsers = parser.add_subparsers(dest="command", required=True)

    gen = subparsers.add_parser(
        "generate-anchor-map",
        help="Generate an editable template anchor map YAML file.",
    )
    gen.add_argument(
        "--template",
        type=Path,
        required=True,
        help="Path to template .pptx file.",
    )
    gen.add_argument(
        "--out",
        type=Path,
        default=Path("template_anchor_map.yaml"),
        help="Output path for anchor map file.",
    )
    gen.add_argument(
        "--default-template-page",
        type=int,
        default=5,
        help="Default template page for generic content slides.",
    )
    gen.add_argument(
        "--no-shape-catalog",
        action="store_true",
        help="Do not include template shape catalog in output.",
    )
    return parser


def main(argv: Optional[Sequence[str]] = None) -> int:
    parser = _build_parser()
    args = parser.parse_args(argv)

    if args.command == "generate-anchor-map":
        out_path = SlideBuilder.generate_anchor_map_file(
            out=args.out,
            template=args.template,
            default_template_page=args.default_template_page,
            include_shape_catalog=not args.no_shape_catalog,
        )
        print(f"Saved anchor map -> {out_path}")
        return 0

    parser.error(f"Unknown command: {args.command}")
    return 2


if __name__ == "__main__":
    raise SystemExit(main())
