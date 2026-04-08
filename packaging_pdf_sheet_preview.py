#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Превью PDF для раскладки на лист (мм): вписывание первой страницы в слот как PNG."""

from __future__ import annotations

from pathlib import Path
from typing import Any

import fitz

from pdf_outline_to_svg import (
    DEFAULT_KNIFE_COLOR_TOLERANCE,
    DEFAULT_KNIFE_MIN_WIDTH_PT,
    knife_bbox_union_pdf_points,
    open_pdf_document,
)


def resolve_pdf_path(pdf_root: Path, file_value: str | None) -> Path | None:
    """Путь к PDF: абсолютный файл или pdf_root / относительное имя."""
    if not file_value or not str(file_value).strip():
        return None
    raw = str(file_value).strip()
    p = Path(raw).expanduser()
    if p.is_file():
        return p.resolve()
    cand = (pdf_root / raw).resolve()
    if cand.is_file():
        return cand
    name_only = Path(raw).name
    if name_only and name_only != raw:
        c2 = (pdf_root / name_only).resolve()
        if c2.is_file():
            return c2
    return None


def render_first_page_fit_to_mm(
    path_str: str,
    slot_w_mm: float,
    slot_h_mm: float,
    *,
    dpi: float = 144.0,
    transparent_bg: bool = False,
) -> bytes | None:
    """
    Первая страница PDF → PNG, вписанная в прямоугольник slot_w_mm × slot_h_mm (как contain).

    При ``transparent_bg=True`` pixmap с альфа-каналом (RGBA): прозрачны незакрашенные области PDF.
    Сплошная белая подложка в макете остаётся непрозрачной (это не удаление фона как в графическом редакторе).
    """
    if slot_w_mm <= 0 or slot_h_mm <= 0:
        return None
    tw = max(48, int(slot_w_mm / 25.4 * dpi))
    th = max(48, int(slot_h_mm / 25.4 * dpi))
    doc = open_pdf_document(path_str)
    if doc is None:
        return None
    try:
        page = doc.load_page(0)
        r = page.rect
        pw = max(r.width, 0.01)
        ph = max(r.height, 0.01)
        s = min(tw / pw, th / ph)
        s = max(0.02, min(s, 8.0))
        mat = fitz.Matrix(s, s) * page.derotation_matrix
        use_alpha = bool(transparent_bg)
        try:
            pix = page.get_pixmap(
                matrix=mat, alpha=use_alpha, colorspace=fitz.csRGB, annots=True
            )
        except Exception:
            try:
                pix = page.get_pixmap(matrix=mat, alpha=use_alpha, colorspace=fitz.csRGB)
            except Exception:
                pix = page.get_pixmap(matrix=mat, alpha=False, colorspace=fitz.csRGB)
        if pix.width <= 0 or pix.height <= 0:
            return None
        return pix.tobytes("png")
    except Exception:
        return None
    finally:
        try:
            doc.close()
        except Exception:
            pass


def render_knife_bbox_fit_to_mm(
    path_str: str,
    slot_w_mm: float,
    slot_h_mm: float,
    *,
    dpi: float = 144.0,
    page_index: int = 0,
    transparent_bg: bool = False,
    **outline_kwargs: Any,
) -> bytes | None:
    """
    Первая страница PDF → PNG: только область union bbox контура ножа (те же фильтры, что у контура),
    вписанная в slot_w_mm × slot_h_mm (contain). Если контура нет — ``None``.

    ``transparent_bg`` — как в :func:`render_first_page_fit_to_mm`.
    """
    if slot_w_mm <= 0 or slot_h_mm <= 0:
        return None
    tw = max(48, int(slot_w_mm / 25.4 * dpi))
    th = max(48, int(slot_h_mm / 25.4 * dpi))
    doc = open_pdf_document(path_str)
    if doc is None:
        return None
    try:
        if doc.page_count < 1 or page_index < 0 or page_index >= doc.page_count:
            return None
        page = doc.load_page(page_index)
        u = knife_bbox_union_pdf_points(
            page,
            target_hex_colors=outline_kwargs.get("target_hex_colors"),
            color_tolerance=float(outline_kwargs.get("color_tolerance", DEFAULT_KNIFE_COLOR_TOLERANCE)),
            min_width_pt=float(outline_kwargs.get("min_width_pt", DEFAULT_KNIFE_MIN_WIDTH_PT)),
            max_width_pt=outline_kwargs.get("max_width_pt"),
            exclude_gray_auxiliary=bool(outline_kwargs.get("exclude_gray_auxiliary", True)),
            gray_exclude_hex=str(outline_kwargs.get("gray_exclude_hex", "34302F")),
        )
        if u is None:
            return None
        x0, y0, x1, y1 = u
        bbox_w = max(x1 - x0, 0.01)
        bbox_h = max(y1 - y0, 0.01)
        clip = fitz.Rect(x0, y0, x1, y1)
        try:
            clip = clip & page.rect
        except Exception:
            pass
        if clip.is_empty:
            return None
        s = min(tw / bbox_w, th / bbox_h)
        s = max(0.02, min(s, 8.0))
        mat = fitz.Matrix(s, s) * page.derotation_matrix
        use_alpha = bool(transparent_bg)
        try:
            pix = page.get_pixmap(
                matrix=mat,
                alpha=use_alpha,
                colorspace=fitz.csRGB,
                annots=True,
                clip=clip,
            )
        except Exception:
            try:
                pix = page.get_pixmap(
                    matrix=mat, alpha=use_alpha, colorspace=fitz.csRGB, clip=clip
                )
            except Exception:
                pix = page.get_pixmap(
                    matrix=mat, alpha=False, colorspace=fitz.csRGB, clip=clip
                )
        if pix.width <= 0 or pix.height <= 0:
            return None
        return pix.tobytes("png")
    except Exception:
        return None
    finally:
        try:
            doc.close()
        except Exception:
            pass


def render_cached_svg_knife_fit_to_mm(
    svg_full: str,
    slot_w_mm: float,
    slot_h_mm: float,
    *,
    dpi: float = 144.0,
    transparent_bg: bool = False,
) -> bytes | None:
    """
    Сохранённый SVG ножа (из ``knife_cache``) → PNG, вписанный в слот (contain), как у PDF-ножа.

    PyMuPDF открывает SVG как документ; при ``transparent_bg=True`` — pixmap с альфой.
    """
    if slot_w_mm <= 0 or slot_h_mm <= 0:
        return None
    raw = (svg_full or "").strip()
    if not raw:
        return None
    tw = max(48, int(slot_w_mm / 25.4 * dpi))
    th = max(48, int(slot_h_mm / 25.4 * dpi))
    try:
        doc = fitz.open(stream=raw.encode("utf-8"), filetype="svg")
    except Exception:
        return None
    try:
        if doc.page_count < 1:
            return None
        page = doc.load_page(0)
        r = page.rect
        pw = max(r.width, 0.01)
        ph = max(r.height, 0.01)
        s = min(tw / pw, th / ph)
        s = max(0.02, min(s, 8.0))
        mat = fitz.Matrix(s, s)
        use_alpha = bool(transparent_bg)
        try:
            pix = page.get_pixmap(
                matrix=mat, alpha=use_alpha, colorspace=fitz.csRGB
            )
        except Exception:
            pix = page.get_pixmap(matrix=mat, alpha=False, colorspace=fitz.csRGB)
        if pix.width <= 0 or pix.height <= 0:
            return None
        return pix.tobytes("png")
    except Exception:
        return None
    finally:
        try:
            doc.close()
        except Exception:
            pass
