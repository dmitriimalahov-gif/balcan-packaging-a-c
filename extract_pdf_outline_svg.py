#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""CLI: извлечь контур (обводки по цвету/толщине) из PDF → SVG."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from pdf_outline_to_svg import (
    DEFAULT_KNIFE_COLOR_TOLERANCE,
    DEFAULT_KNIFE_MIN_WIDTH_PT,
    extract_outline_svg_from_pdf,
)


def main() -> int:
    ap = argparse.ArgumentParser(
        description="Извлечь из PDF векторные обводки (как эталон Corel: магента #E61081 и др.) в SVG.",
    )
    ap.add_argument("input_pdf", type=Path, help="Входной PDF")
    ap.add_argument("output_svg", type=Path, help="Выходной .svg")
    ap.add_argument("--page", type=int, default=0, help="Индекс страницы (с нуля)")
    ap.add_argument(
        "--color",
        action="append",
        dest="colors",
        metavar="HEX",
        help="Целевой цвет обводки без #, можно повторять (по умолчанию E61081 DD0031 E02020)",
    )
    ap.add_argument(
        "--color-tolerance",
        type=float,
        default=DEFAULT_KNIFE_COLOR_TOLERANCE,
        help=f"Порог близости RGB 0..√3 (по умолчанию {DEFAULT_KNIFE_COLOR_TOLERANCE:g})",
    )
    ap.add_argument(
        "--min-width-pt",
        type=float,
        default=DEFAULT_KNIFE_MIN_WIDTH_PT,
        help=f"Мин. толщина линии в pt (по умолчанию {DEFAULT_KNIFE_MIN_WIDTH_PT:g})",
    )
    ap.add_argument("--max-width-pt", type=float, default=None, help="Макс. толщина линии в pt (опц.)")
    ap.add_argument(
        "--no-exclude-gray",
        action="store_true",
        help="Не отсекать тёмно-серые вспомогательные линии (~#34302F)",
    )
    ap.add_argument(
        "--viewbox",
        choices=("page", "content"),
        default="page",
        help="Холст: вся страница или обрезка по bbox отобранных rect",
    )
    ap.add_argument(
        "--output-stroke",
        default="E61081",
        help="Цвет обводки в SVG (hex без #)",
    )
    args = ap.parse_args()

    inp = args.input_pdf.expanduser().resolve()
    if not inp.is_file():
        print(f"Файл не найден: {inp}", file=sys.stderr)
        return 2

    svg = extract_outline_svg_from_pdf(
        str(inp),
        args.page,
        target_hex_colors=args.colors if args.colors else None,
        color_tolerance=args.color_tolerance,
        min_width_pt=args.min_width_pt,
        max_width_pt=args.max_width_pt,
        exclude_gray_auxiliary=not args.no_exclude_gray,
        viewbox=args.viewbox,
        output_stroke_hex=args.output_stroke,
    )
    if not (svg or "").strip():
        print(
            "Не удалось извлечь обводки: проверьте цвета, толщину или что в PDF есть векторные линии.",
            file=sys.stderr,
        )
        return 1

    out = args.output_svg.expanduser().resolve()
    try:
        out.parent.mkdir(parents=True, exist_ok=True)
        out.write_text(svg, encoding="utf-8")
    except OSError as e:
        print(f"Ошибка записи: {e}", file=sys.stderr)
        return 2

    print(f"Записано: {out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
