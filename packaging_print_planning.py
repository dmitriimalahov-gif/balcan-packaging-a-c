#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Раскладка оттисков на печатный лист (мм), заявки и приоритеты печати."""

from __future__ import annotations

import io
import math
import re
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any

import pandas as pd

import import_cutii_forecast as icf
from packaging_sizes import (
    canonicalize_size_mm,
    extract_gabarit_mm_values,
    size_key_from_string,
)

# --- Геометрия листа ---


@dataclass(frozen=True)
class SheetParams:
    """
    Параметры печатного листа (мм).

    ``gap_mm`` — зазор **по X** (между оттисками в одном ряду).
    ``gap_y_mm`` — зазор **по Y** (между рядами).
    Допускаются отрицательные значения (сближение / нахлёст).
    """

    width_mm: float
    height_mm: float
    margin_mm: float = 5.0
    gap_mm: float = 2.0
    gap_y_mm: float = 2.0


@dataclass
class PlacedRect:
    x: float
    y: float
    w: float
    h: float
    rotated: bool


def size_key_display(key_str: str) -> str:
    """Человекочитаемая подпись ключа габаритов (как в основном UI)."""
    if key_str == "__empty__":
        return "Без размера"
    parts = [int(x) for x in key_str.split("|")]
    while parts and parts[-1] == 0:
        parts.pop()
    if not parts:
        return "Без размера"
    return " × ".join(str(p) for p in parts) + " mm"


def _sort_key_for_size_key(sk: str) -> tuple:
    """Порядок списка габаритов: по мм (лексикографически по убывающему кортежу из ключа), не по числу позиций."""
    if sk == "__empty__":
        return (1, ())
    return (0, tuple(int(x) for x in sk.split("|")))


def layout_size_str_from_db_row(row: dict[str, Any]) -> str:
    """
    Габариты для раскладки только из поля size строки (как в SQLite packaging_items).
    Без PDF и без подстановок из других источников.
    """
    return canonicalize_size_mm(row.get("size") or "")


def collect_box_size_groups(box_rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """
    Группы коробок с одинаковым ключом габаритов (перестановки мм совпадают).
    Размер — строго из поля size каждой строки (ожидаются данные из БД).
    Каждая группа: size_key, rows, sample_size_str (строка для геометрии).
    """
    groups_map: dict[str, list[dict[str, Any]]] = {}
    for r in box_rows:
        eff = layout_size_str_from_db_row(r)
        sk = size_key_from_string(eff)
        if sk == "__empty__":
            continue
        groups_map.setdefault(sk, []).append(r)
    out: list[dict[str, Any]] = []
    for sk, rows in groups_map.items():
        rows.sort(key=lambda x: int(x["excel_row"]))
        sample = layout_size_str_from_db_row(rows[0])
        out.append(
            {
                "size_key": sk,
                "rows": rows,
                "sample_size_str": sample,
            }
        )
    out.sort(key=lambda g: _sort_key_for_size_key(g["size_key"]))
    return out


def footprint_mm_from_size(size_str: str) -> tuple[float, float] | None:
    """
    Две наибольшие ненулевые величины из строки размера как прямоугольник оттиска (эвристика).
    Сначала канонизация (как в макетах), чтобы лишние числа из Excel не ломали разбор.
    """
    raw = (size_str or "").strip()
    s = canonicalize_size_mm(raw) or raw
    dims = extract_gabarit_mm_values(s)
    if len(dims) < 2:
        return None
    sd = sorted([d for d in dims if d > 0.01], reverse=True)
    if len(sd) < 2:
        return None
    w, h = float(sd[0]), float(sd[1])
    if w <= 0 or h <= 0:
        return None
    return w, h


def inner_sheet_size(params: SheetParams) -> tuple[float, float]:
    m = max(0.0, params.margin_mm)
    iw = max(0.0, params.width_mm - 2 * m)
    ih = max(0.0, params.height_mm - 2 * m)
    return iw, ih


def pack_shelf_single_item(
    params: SheetParams,
    item_w: float,
    item_h: float,
) -> tuple[int, list[PlacedRect], float]:
    """
    Жадная укладка одинаковых прямоугольников по полкам с поворотом на 90°.
    Возвращает: (количество, размещения в координатах внутренней области), % заполнения площади.
    """
    gap_x = float(params.gap_mm)
    gap_y = float(params.gap_y_mm)
    inner_w, inner_h = inner_sheet_size(params)
    if inner_w <= 0 or inner_h <= 0 or item_w <= 0 or item_h <= 0:
        return 0, [], 0.0

    placements: list[PlacedRect] = []
    x_cursor = 0.0
    y_cursor = 0.0
    row_h = 0.0

    def try_orientations(cx: float, cy: float) -> tuple[float, float, bool] | None:
        for rot in (False, True):
            w, h = (item_h, item_w) if rot else (item_w, item_h)
            if cx + w <= inner_w + 1e-6 and cy + h <= inner_h + 1e-6:
                return (w, h, rot)
        return None

    while y_cursor < inner_h - 1e-9:
        fit = try_orientations(x_cursor, y_cursor)
        if fit is not None:
            w, h, rot = fit
            placements.append(PlacedRect(x_cursor, y_cursor, w, h, rot))
            x_cursor += w + gap_x
            row_h = max(row_h, h)
            continue
        if abs(x_cursor) > 1e-9:
            x_cursor = 0.0
            dy = row_h + gap_y
            if dy < 1e-6:
                dy = 1e-6
            y_cursor += dy
            row_h = 0.0
            continue
        break

    used = sum(p.w * p.h for p in placements)
    area = inner_w * inner_h
    fill_pct = 100.0 * used / area if area > 0 else 0.0
    return len(placements), placements, fill_pct


def _svg_knife_content_transform(
    cell_w: float,
    cell_h: float,
    *,
    rotate_deg: int = 0,
    flip_h: bool = False,
    flip_v: bool = False,
) -> str:
    """
    Атрибут ``transform`` для группы: поворот и зеркало **вокруг центра ячейки** (нож в слоте),
    лист и сетка слотов остаются без изменений.
    """
    r = int(rotate_deg) % 360
    if r not in (0, 90, 180, 270):
        r = 0
    sx = -1.0 if flip_h else 1.0
    sy = -1.0 if flip_v else 1.0
    if r == 0 and sx == 1.0 and sy == 1.0:
        return ""
    cx = float(cell_w) / 2.0
    cy = float(cell_h) / 2.0
    return (
        f"translate({cx:.6f},{cy:.6f}) rotate({r}) scale({sx:.6f},{sy:.6f}) "
        f"translate({-cx:.6f},{-cy:.6f})"
    )


def imposition_preview_svg_mm(
    params: SheetParams,
    placements: list[PlacedRect],
    *,
    slot_image_b64: str | None = None,
    knife_rotate_deg: int = 0,
    knife_flip_h: bool = False,
    knife_flip_v: bool = False,
) -> str:
    """
    Компактная SVG-схема листа (мм): поле, внутренняя область и ячейки по размещениям.
    Координаты placements — как в ``pack_shelf_single_item`` (внутренняя область).
    Если задан ``slot_image_b64`` (PNG base64), в каждой ячейке — то же изображение (импозиция одного ножа).
    ``knife_*`` — поворот/зеркало **содержимого** каждой ячейки (ножа), не всего листа.
    Разные макеты по слотам — в ``sheet_layout_svg``. В Streamlit-viewer превью листа — только ``sheet_layout_svg``;
    эта функция оставлена для скриптов и внешнего использования.
    """
    m = max(0.0, params.margin_mm)
    w = max(0.01, params.width_mm)
    h = max(0.01, params.height_mm)
    iw, ih = inner_sheet_size(params)
    parts: list[str] = [
        f'<svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" '
        f'width="{w:.2f}mm" height="{h:.2f}mm" '
        f'viewBox="0 0 {w:.4f} {h:.4f}" style="shape-rendering:geometricPrecision">',
        f"<title>imposition</title>",
        f'<rect x="0" y="0" width="{w:.4f}" height="{h:.4f}" fill="#f8f8f8" stroke="#444" stroke-width="0.6"/>',
        f'<rect x="{m:.4f}" y="{m:.4f}" width="{iw:.4f}" height="{ih:.4f}" '
        f'fill="none" stroke="#888" stroke-dasharray="3 3" stroke-width="0.45"/>',
    ]
    for i, p in enumerate(placements):
        x = m + p.x
        y = m + p.y
        if slot_image_b64:
            cid = f"impos_clip_{i}"
            parts.append(f'<g transform="translate({x:.4f},{y:.4f})">')
            parts.append("<defs>")
            parts.append(
                f'<clipPath id="{cid}"><rect x="0" y="0" width="{p.w:.4f}" height="{p.h:.4f}"/></clipPath>'
            )
            parts.append("</defs>")
            parts.append(f'<g clip-path="url(#{cid})">')
            _kt = _svg_knife_content_transform(
                float(p.w),
                float(p.h),
                rotate_deg=int(knife_rotate_deg),
                flip_h=bool(knife_flip_h),
                flip_v=bool(knife_flip_v),
            )
            if _kt:
                parts.append(f'<g transform="{_kt}">')
            parts.append(
                f'<image href="data:image/png;base64,{slot_image_b64}" x="0" y="0" '
                f'width="{p.w:.4f}" height="{p.h:.4f}" preserveAspectRatio="xMidYMid meet"/>'
            )
            if _kt:
                parts.append("</g>")
            parts.append("</g>")
            parts.append(
                f'<rect x="0" y="0" width="{p.w:.4f}" height="{p.h:.4f}" fill="none" '
                f'stroke="#E61081" stroke-width="0.45"/>'
            )
            parts.append("</g>")
        else:
            parts.append(
                f'<rect x="{x:.4f}" y="{y:.4f}" width="{p.w:.4f}" height="{p.h:.4f}" '
                f'fill="rgba(230,16,129,0.1)" stroke="#E61081" stroke-width="0.5"/>'
            )
    parts.append("</svg>")
    return '<?xml version="1.0" encoding="UTF-8"?>\n' + "\n".join(parts)


def parse_qty_per_sheet(raw: str | None) -> int | None:
    if raw is None:
        return None
    s = str(raw).strip().replace(",", ".")
    if not s:
        return None
    m = re.search(r"\d+", s)
    if not m:
        return None
    try:
        n = int(m.group(0))
    except ValueError:
        return None
    return n if n > 0 else None


def geometry_vs_db_qty(
    params: SheetParams,
    size_str: str,
    qty_db: str | None,
    *,
    mismatch_warn_ratio: float = 0.25,
) -> dict[str, Any]:
    """Сравнение расчёта по геометрии с «Кол-во на листе» из БД."""
    fp = footprint_mm_from_size(size_str)
    db_n = parse_qty_per_sheet(qty_db)
    if fp is None:
        return {
            "footprint_ok": False,
            "geom_count": None,
            "db_count": db_n,
            "fill_pct": None,
            "placements": [],
            "mismatch": False,
            "note": "Нет двух габаритов в поле «Размер (мм)»",
        }
    w, h = fp
    n, pl, fill = pack_shelf_single_item(params, w, h)
    mismatch = False
    if db_n is not None and n > 0:
        if abs(n - db_n) / max(n, db_n) > mismatch_warn_ratio:
            mismatch = True
    return {
        "footprint_ok": True,
        "footprint_w": w,
        "footprint_h": h,
        "geom_count": n,
        "db_count": db_n,
        "fill_pct": fill,
        "placements": pl,
        "mismatch": mismatch,
        "note": "",
    }


def _svg_escape(s: str) -> str:
    return (
        s.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


def _wrap_slot_label_lines(text: str, chars_per_line: int) -> list[str]:
    """Разбивка подписи слота по строкам (по словам; очень длинные слова режутся)."""
    text = (text or "").strip()
    if not text:
        return []
    cpl = max(3, int(chars_per_line))
    words = text.split()
    lines: list[str] = []
    current = ""
    for word in words:
        trial = word if not current else f"{current} {word}"
        if len(trial) <= cpl:
            current = trial
            continue
        if current:
            lines.append(current)
            current = ""
        if len(word) <= cpl:
            current = word
        else:
            rest = word
            while len(rest) > cpl:
                lines.append(rest[:cpl])
                rest = rest[cpl:]
            current = rest
    if current:
        lines.append(current)
    return lines


def _truncate_slot_label_lines(lines: list[str], max_lines: int, cpl: int) -> list[str]:
    """Не больше max_lines; хвост сливается в последнюю строку с … при переполнении."""
    if max_lines < 1:
        return []
    if len(lines) <= max_lines:
        return lines
    head = lines[: max_lines - 1]
    tail = " ".join(lines[max_lines - 1 :])
    cpl = max(3, int(cpl))
    if len(tail) <= cpl:
        head.append(tail)
    else:
        head.append(tail[: max(1, cpl - 1)] + "…")
    return head


def _svg_caption_multiline(
    lines: list[str],
    *,
    center_x: float,
    first_baseline_y: float,
    line_height: float,
    font_size: float,
    fill: str = "#0d47a1",
) -> str:
    """Один <text> с <tspan> на строку (перенос внутри ячейки)."""
    if not lines:
        return ""
    tspans: list[str] = []
    for i, line in enumerate(lines):
        dy = 0.0 if i == 0 else float(line_height)
        tspans.append(
            f'<tspan x="{center_x}" dy="{dy}">{_svg_escape(line)}</tspan>'
        )
    return (
        f'<text xml:space="preserve" x="{center_x}" y="{first_baseline_y}" '
        f'font-size="{font_size}" text-anchor="middle" dominant-baseline="alphabetic" '
        f'fill="{fill}">{"".join(tspans)}</text>'
    )


def sheet_layout_svg(
    params: SheetParams,
    placements: list[PlacedRect],
    *,
    title: str = "",
    slot_labels: list[str | None] | None = None,
    slot_images_b64: list[str | None] | None = None,
    slot_outline_svg_inner: list[str | None] | None = None,
    highlight_slot_index: int | None = None,
    slot_image_gray_matte: bool = True,
    knife_rotate_deg: int = 0,
    knife_flip_h: bool = False,
    knife_flip_v: bool = False,
) -> str:
    """
    SVG листа: слоты с опциональными PNG (превью PDF) и опциональным контуром (фрагменты inner SVG).
    Координаты слотов — внутренняя область + margin.
    Номер слота и подпись коробки привязаны к **низу** ячейки (полоса с названием внизу, номер над ней).
    ``slot_image_gray_matte`` — серая подложка под PNG в ячейке; выключайте при PNG с альфой, чтобы был виден белый фон листа.
    ``knife_*`` — поворот и зеркало **макета внутри каждой ячейки** (нож), рамка листа и позиции слотов не меняются.
    """
    m = max(0.0, params.margin_mm)
    sw, sh = params.width_mm, params.height_mm
    iw, ih = inner_sheet_size(params)
    stroke = 0.4
    parts = [
        f'<svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" '
        f'viewBox="0 0 {sw} {sh}" width="100%" height="auto">',
        f'<rect x="0" y="0" width="{sw}" height="{sh}" fill="#f8f8f8" stroke="#333" stroke-width="{stroke}"/>',
        f'<rect x="{m}" y="{m}" width="{iw}" height="{ih}" fill="#fff" stroke="#999" stroke-dasharray="2,2" stroke-width="{stroke}"/>',
    ]
    for i, p in enumerate(placements):
        x = m + p.x
        y = m + p.y
        raw_l = ""
        if slot_labels is not None and i < len(slot_labels):
            raw_l = (slot_labels[i] or "").strip()
        img_b64: str | None = None
        if slot_images_b64 is not None and i < len(slot_images_b64):
            v = slot_images_b64[i]
            img_b64 = v if v else None
        outline_part = ""
        if slot_outline_svg_inner is not None and i < len(slot_outline_svg_inner):
            outline_part = (slot_outline_svg_inner[i] or "").strip()

        parts.append(f'<g transform="translate({x},{y})">')
        if raw_l:
            parts.append(f'<title>{_svg_escape(raw_l)}</title>')
        cid = f"pp_cp_{i}"
        parts.append("<defs>")
        parts.append(
            f'<clipPath id="{cid}"><rect x="0" y="0" width="{p.w}" height="{p.h}"/></clipPath>'
        )
        parts.append("</defs>")
        # Растр → контур поверх (линия реза видна на макете). Подписи и номер — к низу ячейки, столбиком вниз.
        parts.append(f'<g clip-path="url(#{cid})">')
        has_knife = bool(img_b64 or outline_part)
        _kt = ""
        if has_knife:
            _kt = _svg_knife_content_transform(
                float(p.w),
                float(p.h),
                rotate_deg=int(knife_rotate_deg),
                flip_h=bool(knife_flip_h),
                flip_v=bool(knife_flip_v),
            )
        if _kt:
            parts.append(f'<g transform="{_kt}">')
        if img_b64:
            if slot_image_gray_matte:
                parts.append(f'<rect x="0" y="0" width="{p.w}" height="{p.h}" fill="#eeeeee"/>')
            parts.append(
                f'<image href="data:image/png;base64,{img_b64}" x="0" y="0" '
                f'width="{p.w}" height="{p.h}" preserveAspectRatio="xMidYMid meet"/>'
            )
        else:
            parts.append(
                f'<rect x="0" y="0" width="{p.w}" height="{p.h}" fill="#cfe8ff" opacity="0.92"/>'
            )
        if outline_part:
            parts.append(
                f'<g opacity="0.92" pointer-events="none">{outline_part}</g>'
            )
        if _kt:
            parts.append("</g>")
        parts.append("</g>")
        hi = highlight_slot_index is not None and i == highlight_slot_index
        if hi:
            parts.append(
                f'<rect x="0" y="0" width="{p.w}" height="{p.h}" fill="#ff9800" fill-opacity="0.18"/>'
            )
        sw_st = 1.85 if hi else stroke
        col_st = "#e65100" if hi else "#1565c0"
        parts.append(
            f'<rect x="0" y="0" width="{p.w}" height="{p.h}" fill="none" '
            f'stroke="{col_st}" stroke-width="{sw_st}"/>'
        )
        idx_fs = max(2.0, min(p.w, p.h) * 0.11)
        tfs = max(2.0, min(p.w, p.h) * 0.065)
        pad_x = 0.4
        usable_w = max(0.5, float(p.w) - 2 * pad_x)
        avg_char_w = tfs * 0.52
        cpl = max(3, int(usable_w / avg_char_w))
        lh_cap = tfs * 1.2
        reserve_idx = idx_fs * 1.2
        max_h_caption = max(lh_cap, float(p.h) * 0.48 - reserve_idx)
        max_lines_cap = max(1, min(8, int(max_h_caption / lh_cap)))
        caption_lines: list[str] = []
        if raw_l:
            caption_lines = _truncate_slot_label_lines(
                _wrap_slot_label_lines(raw_l, cpl),
                max_lines_cap,
                cpl,
            )
        # Номер и подпись — к низу ячейки: полоса с названием (несколько строк), выше — номер слота.
        if img_b64 and caption_lines:
            n_ln = len(caption_lines)
            bar_h = min(float(p.h) * 0.55, n_ln * lh_cap + tfs * 0.35)
            idx_y = max(idx_fs * 0.9, p.h - bar_h - idx_fs * 0.95)
            y_last = p.h - tfs * 0.22
            y_first = y_last - (n_ln - 1) * lh_cap
            parts.append(
                f'<rect x="0" y="{p.h - bar_h}" width="{p.w}" height="{bar_h}" '
                f'fill="#ffffff" fill-opacity="0.88"/>'
            )
            parts.append(
                _svg_caption_multiline(
                    caption_lines,
                    center_x=float(p.w) / 2.0,
                    first_baseline_y=y_first,
                    line_height=lh_cap,
                    font_size=tfs,
                )
            )
            parts.append(
                f'<text x="{idx_fs * 0.55}" y="{idx_y}" font-size="{idx_fs}" '
                f'fill="{col_st}" font-weight="600">{i + 1}</text>'
            )
        elif not img_b64 and raw_l:
            fs = max(2.5, min(p.w, p.h) * 0.09)
            avg_w2 = fs * 0.52
            cpl2 = max(3, int(usable_w / avg_w2))
            lh2 = fs * 1.2
            max_h2 = max(lh2, float(p.h) * 0.42 - reserve_idx)
            max_ln2 = max(1, min(8, int(max_h2 / lh2)))
            lines2 = _truncate_slot_label_lines(
                _wrap_slot_label_lines(raw_l, cpl2),
                max_ln2,
                cpl2,
            )
            n2 = len(lines2)
            idx_y = max(idx_fs * 0.9, p.h - n2 * lh2 - fs * 0.5 - idx_fs * 0.95)
            parts.append(
                f'<text x="{idx_fs * 0.55}" y="{idx_y}" font-size="{idx_fs}" '
                f'fill="{col_st}" font-weight="600">{i + 1}</text>'
            )
            y_last2 = p.h - fs * 0.28
            y_first2 = y_last2 - (n2 - 1) * lh2
            parts.append(
                _svg_caption_multiline(
                    lines2,
                    center_x=float(p.w) / 2.0,
                    first_baseline_y=y_first2,
                    line_height=lh2,
                    font_size=fs,
                )
            )
        else:
            idx_y = max(idx_fs * 0.9, p.h - idx_fs * 0.45)
            parts.append(
                f'<text x="{idx_fs * 0.55}" y="{idx_y}" font-size="{idx_fs}" '
                f'fill="{col_st}" font-weight="600">{i + 1}</text>'
            )
        parts.append("</g>")
    if title:
        parts.append(
            f'<text x="{sw/2}" y="{min(8, sh * 0.03)}" font-size="{min(6, sw * 0.02)}" '
            f'text-anchor="middle" fill="#333">{_svg_escape(title)}</text>'
        )
    parts.append("</svg>")
    return "\n".join(parts)


# --- Заявки ---


def read_orders_file(content: bytes, filename: str) -> pd.DataFrame:
    name = (filename or "").lower()
    bio = io.BytesIO(content)
    if name.endswith(".xlsx") or name.endswith(".xls"):
        return pd.read_excel(bio)
    return pd.read_csv(bio)


def _to_float_qty(v: Any) -> float:
    if v is None or (isinstance(v, float) and math.isnan(v)):
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(",", ".")
    s = re.sub(r"[^\d.\-]", "", s)
    if not s:
        return 0.0
    try:
        return float(s)
    except ValueError:
        return 0.0


def _parse_month_year(v: Any, default_year: int) -> tuple[int, int] | None:
    """Вернуть (year, month) или None."""
    if v is None or (isinstance(v, float) and math.isnan(v)):
        return None
    if hasattr(v, "year") and hasattr(v, "month"):
        try:
            return int(v.year), int(v.month)
        except (TypeError, ValueError):
            pass
    s = str(v).strip()
    if not s:
        return None
    for fmt in ("%Y-%m-%d", "%d.%m.%Y", "%d/%m/%Y", "%Y-%m", "%m/%Y"):
        try:
            from datetime import datetime as _dt

            d = _dt.strptime(s[:32], fmt)
            return d.year, d.month
        except ValueError:
            continue
    m = re.match(r"^(\d{1,2})\s*$", s)
    if m:
        mo = int(m.group(1))
        if 1 <= mo <= 12:
            return default_year, mo
    m = re.match(r"^(\d{4})-(\d{1,2})$", s)
    if m:
        return int(m.group(1)), int(m.group(2))
    return None


def build_order_records(
    df: pd.DataFrame,
    col_name: str,
    col_qty: str,
    col_month: str,
    default_year: int,
) -> list[dict[str, Any]]:
    out: list[dict[str, Any]] = []
    for _, row in df.iterrows():
        raw_name = row.get(col_name)
        if raw_name is None or (isinstance(raw_name, float) and math.isnan(raw_name)):
            continue
        name = str(raw_name).strip()
        if not name:
            continue
        qty = _to_float_qty(row.get(col_qty))
        if qty <= 0:
            continue
        my = _parse_month_year(row.get(col_month), default_year)
        if my is None:
            continue
        y, mo = my
        out.append(
            {
                "raw_name": name,
                "qty": qty,
                "year": y,
                "month": mo,
            }
        )
    return out


def month_horizon_slices(
    start_year: int,
    start_month: int,
    horizon_months: int,
) -> list[tuple[int, int]]:
    """Включительно: первый месяц — start, далее horizon_months-1."""
    out: list[tuple[int, int]] = []
    y, m = start_year, start_month
    for _ in range(horizon_months):
        out.append((y, m))
        m += 1
        if m > 12:
            m = 1
            y += 1
    return out


def filter_orders_in_horizon(
    orders: list[dict[str, Any]],
    start_y: int,
    start_m: int,
    horizon: int,
) -> list[dict[str, Any]]:
    allowed = set(month_horizon_slices(start_y, start_m, horizon))
    return [o for o in orders if (o["year"], o["month"]) in allowed]


def aggregate_demand_by_excel_row(
    orders: list[dict[str, Any]],
    line_to_er: dict[int, int],
) -> dict[int, float]:
    """line_to_er: индекс строки заявки -> excel_row."""
    acc: dict[int, float] = {}
    for i, o in enumerate(orders):
        er = line_to_er.get(i)
        if er is None:
            continue
        acc[er] = acc.get(er, 0.0) + float(o["qty"])
    return acc


# --- Матчинг к коробкам (как cutii: имя + PDF) ---


def box_rows_only(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    return [r for r in rows if icf.is_packaging_box(r.get("kind") or "")]


def printable_rows_only(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Коробки + блистеры + пакеты — всё, что раскладывается на печатные листы."""
    return [r for r in rows if icf.is_printable_packaging(r.get("kind") or "")]


def sheet_layout_candidate_rows(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """
    Строки для раскладки на лист: коробка, блистер, пакет **и** явная этикетка
    (см. ``import_cutii_forecast.is_sheet_layout_candidate_kind``).
    """
    return [
        r
        for r in rows
        if icf.is_sheet_layout_candidate_kind(str(r.get("kind") or ""))
    ]


def match_order_name(
    raw: str,
    box_rows: list[dict[str, Any]],
    *,
    min_score: int = 50,
    ambiguous_gap: int = 5,
    fallback_pdf: bool = True,
) -> tuple[str, dict[str, Any] | None, int, str, list[dict[str, Any]]]:
    """
    Статус ok|ambiguous|no_match, лучшая строка, счёт, деталь, топ кандидатов для UI.
    """
    tops_raw = icf.top_combined_candidates(raw, box_rows, k=8)
    tops_ui: list[dict[str, Any]] = []
    for best, sn, sf, br in tops_raw:
        tops_ui.append(
            {
                "excel_row": int(br["excel_row"]),
                "score": best,
                "score_name": sn,
                "score_file": sf,
                "name": (br.get("name") or "")[:200],
                "pdf": (br.get("file") or "")[:120],
            }
        )
    st, row, sc, det = icf.pick_match(
        raw,
        box_rows,
        min_score=min_score,
        gap=ambiguous_gap,
        fallback_pdf=fallback_pdf,
    )
    return st, row, sc, det, tops_ui


def auto_match_all_orders(
    orders: list[dict[str, Any]],
    box_rows: list[dict[str, Any]],
    *,
    min_score: int,
    ambiguous_gap: int,
    fallback_pdf: bool,
) -> tuple[list[dict[str, Any]], dict[int, int]]:
    """
    Для каждой заявки — статус и опционально excel_row.
    Возвращает список аннотаций и карту индекс_строки -> excel_row для однозначных ok.
    """
    annotations: list[dict[str, Any]] = []
    auto_map: dict[int, int] = {}
    for i, o in enumerate(orders):
        st, row, sc, det, tops = match_order_name(
            o["raw_name"],
            box_rows,
            min_score=min_score,
            ambiguous_gap=ambiguous_gap,
            fallback_pdf=fallback_pdf,
        )
        er: int | None = int(row["excel_row"]) if st == "ok" and row else None
        if er is not None:
            auto_map[i] = er
        annotations.append(
            {
                "idx": i,
                "raw_name": o["raw_name"],
                "qty": o["qty"],
                "year": o["year"],
                "month": o["month"],
                "status": st,
                "excel_row": er,
                "score": sc,
                "match_via": det,
                "tops": tops,
            }
        )
    return annotations, auto_map


def build_priority_rows(
    demand_by_er: dict[int, float],
    rows_by_er: dict[int, dict[str, Any]],
    params: SheetParams,
    db_rows_by_er: dict[int, dict[str, Any]],
) -> list[dict[str, Any]]:
    """Строки таблицы приоритетов: спрос, листы, заполнение листа. Габариты листа — только size из БД."""
    out: list[dict[str, Any]] = []
    for er, qty in demand_by_er.items():
        item = rows_by_er.get(er)
        if not item:
            continue
        qps = parse_qty_per_sheet(item.get("qty_per_sheet"))
        sheets = math.ceil(qty / qps) if qps and qps > 0 else None
        db_item = db_rows_by_er.get(er)
        eff_sz = layout_size_str_from_db_row(db_item) if db_item else ""
        ginfo = geometry_vs_db_qty(
            params,
            eff_sz,
            db_item.get("qty_per_sheet") if db_item else None,
        )
        out.append(
            {
                "excel_row": er,
                "name": (item.get("name") or "")[:120],
                "demand_qty": qty,
                "qty_per_sheet_db": qps,
                "sheets_estimate": sheets,
                "geom_per_sheet": ginfo.get("geom_count"),
                "fill_pct_sheet": ginfo.get("fill_pct"),
                "geom_db_mismatch": ginfo.get("mismatch"),
            }
        )
    return out


# --- Планировщик оптимизации печати ---


@dataclass
class SlotAllocation:
    excel_row: int
    slots: int
    demand: float
    actual_printed: float
    overprint: float
    overprint_pct: float


@dataclass
class SheetPlan:
    size_key: str
    n_slots: int
    n_sheets: int
    allocations: list[SlotAllocation]
    empty_slots: int
    cost: float


@dataclass
class Scenario:
    horizon_months: int
    plans: list[SheetPlan]
    total_sheets: int
    total_cost: float
    total_units: float
    cost_per_unit: float
    sheets_if_separate: int
    cost_if_separate: float
    savings_abs: float
    savings_pct: float


def merge_demand_sources(
    monthly_db: list[dict[str, Any]],
    orders: list[dict[str, Any]],
    line_to_er: dict[int, int],
    start_y: int,
    start_m: int,
    horizon: int,
    *,
    merge_mode: str = "max",
) -> dict[int, float]:
    """
    Объединённый спрос из двух источников за горизонт.

    merge_mode:
        "max" — для каждого (excel_row, year, month) берём максимум из двух источников;
        "sum" — суммируем;
        "cutii" — только cutii;
        "orders" — только заявки.
    """
    allowed = set(month_horizon_slices(start_y, start_m, horizon))

    cutii: dict[int, dict[tuple[int, int], float]] = {}
    for row in monthly_db:
        ym = (int(row["year"]), int(row["month"]))
        if ym not in allowed:
            continue
        er = int(row["excel_row"])
        cutii.setdefault(er, {})[ym] = float(row["qty"])

    ord_demand: dict[int, dict[tuple[int, int], float]] = {}
    for i, o in enumerate(orders):
        ym = (int(o["year"]), int(o["month"]))
        if ym not in allowed:
            continue
        er = line_to_er.get(i)
        if er is None:
            continue
        ord_demand.setdefault(er, {}).setdefault(ym, 0.0)
        ord_demand[er][ym] += float(o["qty"])

    all_ers = set(cutii.keys()) | set(ord_demand.keys())
    result: dict[int, float] = {}
    for er in all_ers:
        total = 0.0
        ym_keys = set()
        if er in cutii:
            ym_keys |= set(cutii[er].keys())
        if er in ord_demand:
            ym_keys |= set(ord_demand[er].keys())
        for ym in ym_keys:
            vc = cutii.get(er, {}).get(ym, 0.0)
            vo = ord_demand.get(er, {}).get(ym, 0.0)
            if merge_mode == "max":
                total += max(vc, vo)
            elif merge_mode == "sum":
                total += vc + vo
            elif merge_mode == "cutii":
                total += vc
            elif merge_mode == "orders":
                total += vo
            else:
                total += max(vc, vo)
        if total > 0:
            result[er] = total
    return result


def optimize_sheet_allocation(
    demand_by_er: dict[int, float],
    size_groups: list[dict[str, Any]],
    sheet_params: SheetParams,
    tariffs: list[dict[str, Any]],
    overprint_pct: float = 0.05,
) -> list[SheetPlan]:
    """
    Оптимальная раскладка коробок одного размера на листы.
    Возвращает SheetPlan для каждой группы размера, где есть спрос.
    """
    import packaging_db as _pdb

    plans: list[SheetPlan] = []
    for group in size_groups:
        sk = group["size_key"]
        fp = footprint_mm_from_size(group["sample_size_str"])
        if fp is None:
            continue
        fw, fh = fp
        n_total, placements, _ = pack_shelf_single_item(sheet_params, fw, fh)
        if n_total <= 0:
            continue

        ers_with_demand = [
            (int(r["excel_row"]), demand_by_er.get(int(r["excel_row"]), 0.0))
            for r in group["rows"]
            if demand_by_er.get(int(r["excel_row"]), 0.0) > 0
        ]
        if not ers_with_demand:
            continue

        ers_with_demand.sort(key=lambda x: x[1], reverse=True)
        total_demand = sum(d for _, d in ers_with_demand)

        n_kinds = len(ers_with_demand)
        if n_kinds > n_total:
            ers_with_demand = ers_with_demand[:n_total]
            n_kinds = n_total
            total_demand = sum(d for _, d in ers_with_demand)

        raw_slots = [max(1, round(d / total_demand * n_total)) for _, d in ers_with_demand]

        diff = sum(raw_slots) - n_total
        while diff != 0:
            if diff > 0:
                idx = max(range(n_kinds), key=lambda i: raw_slots[i])
                raw_slots[idx] -= 1
            else:
                idx = max(range(n_kinds), key=lambda i: ers_with_demand[i][1] / max(raw_slots[i], 1))
                raw_slots[idx] += 1
            diff = sum(raw_slots) - n_total

        for it in range(15):
            n_sheets = max(1, math.ceil(max(
                ers_with_demand[i][1] / max(raw_slots[i], 1) for i in range(n_kinds)
            )))
            worst_over_idx = -1
            worst_over = -1.0
            for i in range(n_kinds):
                actual = raw_slots[i] * n_sheets
                demand_i = ers_with_demand[i][1]
                over = actual - demand_i
                over_pct = over / max(demand_i, 1.0)
                if over_pct > overprint_pct and over_pct > worst_over:
                    worst_over = over_pct
                    worst_over_idx = i
            if worst_over_idx < 0 or worst_over <= overprint_pct:
                break
            if raw_slots[worst_over_idx] > 1:
                raw_slots[worst_over_idx] -= 1
                best_idx = max(range(n_kinds), key=lambda j: ers_with_demand[j][1] / max(raw_slots[j], 1) if j != worst_over_idx else -1)
                raw_slots[best_idx] += 1

        n_sheets = max(1, math.ceil(max(
            ers_with_demand[i][1] / max(raw_slots[i], 1) for i in range(n_kinds)
        )))

        allocations: list[SlotAllocation] = []
        for i in range(n_kinds):
            er, dem = ers_with_demand[i]
            actual = raw_slots[i] * n_sheets
            over = actual - dem
            allocations.append(SlotAllocation(
                excel_row=er,
                slots=raw_slots[i],
                demand=dem,
                actual_printed=actual,
                overprint=over,
                overprint_pct=(over / max(dem, 1.0)) * 100.0,
            ))

        used_slots = sum(raw_slots)
        empty = n_total - used_slots
        price = _pdb.sheet_price(n_sheets, tariffs)
        cost = price * n_sheets

        plans.append(SheetPlan(
            size_key=sk,
            n_slots=n_total,
            n_sheets=n_sheets,
            allocations=allocations,
            empty_slots=empty,
            cost=cost,
        ))
    return plans


def _separate_sheets_cost(
    demand_by_er: dict[int, float],
    size_groups: list[dict[str, Any]],
    sheet_params: SheetParams,
    tariffs: list[dict[str, Any]],
) -> tuple[int, float]:
    """Если каждый вид печатать отдельно (один на весь лист): листы и стоимость."""
    import packaging_db as _pdb

    total_sheets = 0
    total_cost = 0.0
    for group in size_groups:
        fp = footprint_mm_from_size(group["sample_size_str"])
        if fp is None:
            continue
        fw, fh = fp
        n_total, _, _ = pack_shelf_single_item(sheet_params, fw, fh)
        if n_total <= 0:
            continue
        for r in group["rows"]:
            er = int(r["excel_row"])
            dem = demand_by_er.get(er, 0.0)
            if dem <= 0:
                continue
            sh = math.ceil(dem / n_total)
            total_sheets += sh
            total_cost += _pdb.sheet_price(sh, tariffs) * sh
    return total_sheets, total_cost


def build_planning_scenarios(
    monthly_db: list[dict[str, Any]],
    orders: list[dict[str, Any]],
    line_to_er: dict[int, int],
    size_groups: list[dict[str, Any]],
    sheet_params: SheetParams,
    tariffs: list[dict[str, Any]],
    start_y: int,
    start_m: int,
    overprint_pct: float = 0.05,
    horizons: list[int] | None = None,
    merge_mode: str = "max",
) -> list[Scenario]:
    """Сценарии для нескольких горизонтов."""
    if horizons is None:
        horizons = [1, 2, 3, 6, 12]

    scenarios: list[Scenario] = []
    for h in horizons:
        demand = merge_demand_sources(
            monthly_db, orders, line_to_er,
            start_y, start_m, h,
            merge_mode=merge_mode,
        )
        plans = optimize_sheet_allocation(
            demand, size_groups, sheet_params, tariffs, overprint_pct,
        )
        t_sheets = sum(p.n_sheets for p in plans)
        t_cost = sum(p.cost for p in plans)
        t_units = sum(a.demand for p in plans for a in p.allocations)
        cpu = t_cost / max(t_units, 1.0)

        sep_sheets, sep_cost = _separate_sheets_cost(demand, size_groups, sheet_params, tariffs)
        sav = sep_cost - t_cost
        sav_pct = (sav / max(sep_cost, 0.01)) * 100.0 if sep_cost > 0 else 0.0

        scenarios.append(Scenario(
            horizon_months=h,
            plans=plans,
            total_sheets=t_sheets,
            total_cost=t_cost,
            total_units=t_units,
            cost_per_unit=cpu,
            sheets_if_separate=sep_sheets,
            cost_if_separate=sep_cost,
            savings_abs=sav,
            savings_pct=sav_pct,
        ))
    return scenarios


def run_self_check() -> None:
    """Минимальные проверки без pytest."""
    sp = SheetParams(320, 450, margin_mm=5, gap_mm=2, gap_y_mm=2)
    iw, ih = inner_sheet_size(sp)
    assert abs(iw - 310) < 1e-6 and abs(ih - 440) < 1e-6
    n, pl, fill = pack_shelf_single_item(sp, 50, 30)
    assert n >= 1 and len(pl) == n and 0 <= fill <= 100
    fp = footprint_mm_from_size("84 × 62 × 15 mm")
    assert fp is not None and fp[0] >= fp[1]
    assert parse_qty_per_sheet("12") == 12
    assert parse_qty_per_sheet("") is None
    m = month_horizon_slices(2025, 11, 3)
    assert m == [(2025, 11), (2025, 12), (2026, 1)]
    print("packaging_print_planning self-check OK")


if __name__ == "__main__":
    run_self_check()
