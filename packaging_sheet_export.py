# -*- coding: utf-8 -*-
"""Экспорт схемы печатного листа: растры слотов и сборка PDF (PyMuPDF)."""

from __future__ import annotations

import base64
from pathlib import Path
from typing import Any

import fitz

import packaging_pdf_sheet_preview as ppsp
import pdf_outline_to_svg as pdf_outline
from packaging_print_planning import PlacedRect, SheetParams

MM_TO_PT = 72.0 / 25.4


def _resolve_unicode_font() -> Path | None:
    """Системный шрифт с кириллицей для insert_text (если есть)."""
    candidates = [
        Path("/System/Library/Fonts/Supplemental/Arial Unicode.ttf"),
        Path("/Library/Fonts/Arial Unicode.ttf"),
        Path("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"),
        Path("/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf"),
        Path("C:/Windows/Fonts/arial.ttf"),
        Path("C:/Windows/Fonts/calibri.ttf"),
    ]
    for p in candidates:
        if p.is_file():
            return p
    return None


def build_slot_png_and_outline(
    *,
    pl_active: list[PlacedRect],
    slot_er_list: list[int | None],
    rows_by_er: dict[int, dict[str, Any]],
    pdf_root: Path,
    dpi: float,
    knife_raster: bool,
    transparent_png: bool,
    show_outline: bool,
    knives_by_er: dict[int, dict[str, Any]] | None = None,
) -> tuple[list[str | None], list[str | None], list[bytes | None]]:
    """
    Рендер PNG и контур для каждого слота (без ограничения «первые N»).
    Возвращает (base64_png, inner_outline_svg, raw_png_bytes).

    Если передан ``knives_by_er`` (как из ``knife_cache``): для каждого слота **сначала** берётся
    растр и контур из **PDF этой строки** (если файл есть), иначе — из сохранённого SVG в кэше.
    """
    dpi_eff = max(36.0, min(float(dpi), 240.0))
    b64_list: list[str | None] = []
    outlines: list[str | None] = []
    raw_list: list[bytes | None] = []
    kmap = knives_by_er or {}

    for idx, p_rect in enumerate(pl_active):
        er = slot_er_list[idx] if idx < len(slot_er_list) else None
        pb: bytes | None = None
        ol: str | None = None
        if er is not None:
            row = rows_by_er.get(int(er))
            rel = (row.get("file") or "").strip() if row else ""
            p_pdf = ppsp.resolve_pdf_path(pdf_root, rel) if rel else None
            kn = kmap.get(int(er))
            kn_svg = (kn.get("svg_full") or "").strip() if kn else ""
            kn_w = float(kn.get("width_mm") or 0) if kn else 0.0
            kn_h = float(kn.get("height_mm") or 0) if kn else 0.0
            has_db_knife = bool(kn_svg and kn_w > 0 and kn_h > 0)
            if knife_raster:
                if p_pdf and p_pdf.is_file():
                    pb = ppsp.render_knife_bbox_fit_to_mm(
                        str(p_pdf),
                        float(p_rect.w),
                        float(p_rect.h),
                        dpi=dpi_eff,
                        transparent_bg=transparent_png,
                    )
                if pb is None and has_db_knife:
                    pb = ppsp.render_cached_svg_knife_fit_to_mm(
                        kn_svg,
                        float(p_rect.w),
                        float(p_rect.h),
                        dpi=dpi_eff,
                        transparent_bg=transparent_png,
                    )
                if pb is None and p_pdf and p_pdf.is_file():
                    pb = ppsp.render_first_page_fit_to_mm(
                        str(p_pdf),
                        float(p_rect.w),
                        float(p_rect.h),
                        dpi=dpi_eff,
                        transparent_bg=transparent_png,
                    )
            elif p_pdf and p_pdf.is_file():
                pb = ppsp.render_first_page_fit_to_mm(
                    str(p_pdf),
                    float(p_rect.w),
                    float(p_rect.h),
                    dpi=dpi_eff,
                    transparent_bg=transparent_png,
                )
            if show_outline:
                ol = None
                if p_pdf and p_pdf.is_file():
                    ol = (
                        pdf_outline.extract_outline_svg_inner_for_slot(
                            str(p_pdf),
                            float(p_rect.w),
                            float(p_rect.h),
                            0,
                        )
                        or None
                    )
                if (ol is None or ol == "") and has_db_knife:
                    ol = pdf_outline.inner_outline_from_stored_knife_svg(
                        kn_svg,
                        float(p_rect.w),
                        float(p_rect.h),
                        kn_w,
                        kn_h,
                    ) or None
        raw_list.append(pb)
        b64_list.append(base64.b64encode(pb).decode("ascii") if pb else None)
        outlines.append(ol)

    return b64_list, outlines, raw_list


def sheet_layout_to_pdf_bytes(
    sheet_params: SheetParams,
    pl_active: list[PlacedRect],
    slot_png_bytes: list[bytes | None],
    stats_lines: list[str],
    *,
    title_line: str = "",
) -> bytes:
    """
    Страница 1: визуальный лист с вставленными PNG по слотам и рамками.
    Страница 2: текстовая сводка (кириллица — при наличии системного TTF).
    """
    doc = fitz.open()
    wp = float(sheet_params.width_mm) * MM_TO_PT
    hp = float(sheet_params.height_mm) * MM_TO_PT
    page = doc.new_page(width=wp, height=hp)
    page.draw_rect(page.rect, color=(1, 1, 1), fill=(1, 1, 1))

    mpt = float(sheet_params.margin_mm) * MM_TO_PT
    iw = float(sheet_params.width_mm) * MM_TO_PT - 2 * mpt
    ih = float(sheet_params.height_mm) * MM_TO_PT - 2 * mpt
    inner_r = fitz.Rect(mpt, mpt, mpt + iw, mpt + ih)
    page.draw_rect(inner_r, color=(0.55, 0.55, 0.55), width=0.6, dashes="[3 3] 0")

    for i, pr in enumerate(pl_active):
        x0 = mpt + float(pr.x) * MM_TO_PT
        y0 = mpt + float(pr.y) * MM_TO_PT
        x1 = x0 + float(pr.w) * MM_TO_PT
        y1 = y0 + float(pr.h) * MM_TO_PT
        rect = fitz.Rect(x0, y0, x1, y1)
        stream = slot_png_bytes[i] if i < len(slot_png_bytes) else None
        if stream:
            try:
                page.insert_image(rect, stream=stream, keep_proportion=True)
            except Exception:
                pass
        page.draw_rect(rect, color=(0.09, 0.29, 0.62), width=0.9)

    a4w, a4h = 595.0, 842.0
    p2 = doc.new_page(width=a4w, height=a4h)
    font_path = _resolve_unicode_font()
    fontname = "helv"
    if font_path:
        try:
            p2.insert_font("exf", fontfile=str(font_path))
            fontname = "exf"
        except Exception:
            fontname = "helv"

    left, top0 = 48.0, 52.0
    y = top0
    lh_title = 14.0
    lh_body = 10.5
    fs_title = 11.0
    fs_body = 9.0

    if title_line:
        p2.insert_text((left, y), title_line, fontname=fontname, fontsize=fs_title)
        y += lh_title
    y += 4.0

    for line in stats_lines:
        if y > a4h - 48.0:
            p2 = doc.new_page(width=a4w, height=a4h)
            if font_path and fontname == "exf":
                try:
                    p2.insert_font("exf", fontfile=str(font_path))
                except Exception:
                    pass
            y = top0
        p2.insert_text((left, y), line[:500], fontname=fontname, fontsize=fs_body)
        y += lh_body

    out = doc.tobytes(deflate=True, garbage=4, clean=True)
    doc.close()
    return out
