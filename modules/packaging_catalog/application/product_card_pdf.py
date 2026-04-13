# -*- coding: utf-8 -*-
"""
Сборка PDF-карточки группы препарата (A4, PyMuPDF).

Стр. 1: макеты — коробка крупно слева (~60 %), блистер / этикетка / пакет столбиком справа (~40 %).
Стр. 2: информационный блок — остатки упаковки, остатки субстанции, прогноз продаж, рекомендации.

Исходные PDF-макеты вставляются **векторно** через ``show_pdf_page`` — без
растеризации, с максимальным качеством при любом масштабе печати.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

import fitz

import packaging_pdf_sheet_preview as ppsp
from packaging_db import extract_gmp_code

from .product_card_data import ProductCardData

_A4_W = 595.0
_A4_H = 842.0
_MARGIN = 36.0
_HEADER_H = 48.0
_GAP = 10.0

_KIND_LABELS: dict[str, str] = {
    "box": "Коробка",
    "blister": "Блистер",
    "label": "Этикетка",
    "pack": "Пакет",
}

_ACCENT = (0.15, 0.35, 0.65)
_DARK = (0.12, 0.12, 0.15)
_GRAY = (0.4, 0.4, 0.45)
_LIGHT = (0.6, 0.6, 0.6)
_TABLE_BORDER = (0.78, 0.78, 0.82)
_TABLE_HEADER_BG = (0.92, 0.94, 0.97)
_GREEN = (0.1, 0.55, 0.2)
_RED = (0.75, 0.15, 0.15)


def _resolve_unicode_font() -> Path | None:
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


def _insert_font(page: fitz.Page, font_path: Path | None) -> str:
    if font_path and font_path.is_file():
        try:
            page.insert_font(fontname="uni", fontfile=str(font_path))
            return "uni"
        except Exception:
            pass
    return "helv"


def _resolve_row_pdf(row: dict[str, Any] | None, pdf_root: Path) -> Path | None:
    if row is None:
        return None
    file_val = (row.get("file") or "").strip()
    p = ppsp.resolve_pdf_path(pdf_root, file_val)
    if p is not None and p.is_file():
        return p
    return None


def _show_pdf_vector(
    target_page: fitz.Page,
    rect: fitz.Rect,
    src_path: Path,
) -> bool:
    try:
        src_doc = fitz.open(str(src_path))
    except Exception:
        return False
    try:
        if src_doc.page_count <= 0:
            return False
        target_page.show_pdf_page(rect, src_doc, 0)
        return True
    except Exception:
        return False
    finally:
        try:
            src_doc.close()
        except Exception:
            pass


def _slot_label(row: dict[str, Any] | None, kind_key: str) -> str:
    label = _KIND_LABELS.get(kind_key, kind_key)
    if row is None:
        return f"{label}: —"
    size = (row.get("size") or "").strip()
    name = (row.get("name") or "").strip()[:50]
    parts = [label]
    if size:
        parts.append(size)
    if name:
        parts.append(name)
    return " · ".join(parts)


# ── Helpers for page 2 (info tables) ──────────────────────────────────


def _draw_table_header(
    page: fitz.Page, fn: str, y: float, cols: list[tuple[float, float, str]],
) -> float:
    """Рисует строку заголовка таблицы. Возвращает y после строки."""
    row_h = 20.0
    for x, w, text in cols:
        rect = fitz.Rect(x, y, x + w, y + row_h)
        page.draw_rect(rect, color=_TABLE_BORDER, fill=_TABLE_HEADER_BG, width=0.4)
        page.insert_text((x + 4, y + 14), text, fontname=fn, fontsize=8, color=_DARK)
    return y + row_h


def _draw_table_row(
    page: fitz.Page, fn: str, y: float,
    cols: list[tuple[float, float, str]], *,
    color: tuple[float, float, float] = _DARK,
) -> float:
    row_h = 18.0
    for x, w, text in cols:
        rect = fitz.Rect(x, y, x + w, y + row_h)
        page.draw_rect(rect, color=_TABLE_BORDER, width=0.3)
        page.insert_text((x + 4, y + 12), text, fontname=fn, fontsize=7.5, color=color)
    return y + row_h


def _fmt_qty(qty: float) -> str:
    if qty == 0:
        return "—"
    if qty == int(qty):
        return f"{int(qty):,}".replace(",", " ")
    return f"{qty:,.1f}".replace(",", " ")


# ── Page 2: info block ────────────────────────────────────────────────


def _render_info_page(
    doc: fitz.Document,
    font_path: Path | None,
    title: str,
    card_data: ProductCardData | None,
) -> None:
    page = doc.new_page(width=_A4_W, height=_A4_H)
    fn = _insert_font(page, font_path)

    page.insert_text(
        (_MARGIN, _MARGIN + 14),
        title,
        fontname=fn, fontsize=11, color=_DARK,
    )
    page.insert_text(
        (_MARGIN, _MARGIN + 28),
        "Складские остатки · Прогноз · Рекомендации",
        fontname=fn, fontsize=8.5, color=_GRAY,
    )

    y = _MARGIN + 50.0
    content_w = _A4_W - 2 * _MARGIN

    if card_data is None:
        page.insert_text(
            (_MARGIN, y + 20),
            "Данные недоступны (нет подключения к БД или GMP-код не определён).",
            fontname=fn, fontsize=9, color=_LIGHT,
        )
        return

    # ── 1. Остатки упаковки ──
    page.insert_text((_MARGIN, y), "Остатки упаковки на складах", fontname=fn, fontsize=10, color=_ACCENT)
    y += 16.0

    col_kind_w = content_w * 0.4
    col_qty_w = content_w * 0.3
    col_status_w = content_w * 0.3
    hdr = [
        (_MARGIN, col_kind_w, "Вид упаковки"),
        (_MARGIN + col_kind_w, col_qty_w, "Остаток (шт.)"),
        (_MARGIN + col_kind_w + col_qty_w, col_status_w, "Статус"),
    ]
    y = _draw_table_header(page, fn, y, hdr)

    avg_m = card_data.forecast.avg_monthly if card_data.forecast else 0
    for ps in card_data.packaging_stock:
        status = "—"
        row_color = _DARK
        if avg_m > 0 and ps.qty > 0:
            months = ps.qty / avg_m
            if months > 3:
                status = f"~{months:.0f} мес."
                row_color = _GREEN
            elif months > 1:
                status = f"~{months:.1f} мес."
                row_color = _DARK
            else:
                status = "Мало!"
                row_color = _RED
        elif ps.qty <= 0 and avg_m > 0:
            status = "Нет на складе"
            row_color = _RED

        y = _draw_table_row(page, fn, y, [
            (_MARGIN, col_kind_w, ps.kind_label),
            (_MARGIN + col_kind_w, col_qty_w, _fmt_qty(ps.qty)),
            (_MARGIN + col_kind_w + col_qty_w, col_status_w, status),
        ], color=row_color)

    y += 20.0

    # ── 2. Остатки субстанции ──
    page.insert_text((_MARGIN, y), "Остатки субстанции (препарата)", fontname=fn, fontsize=10, color=_ACCENT)
    y += 16.0

    sub = card_data.substance
    sub_cols = [
        (_MARGIN, content_w * 0.5, "Количество"),
        (_MARGIN + content_w * 0.5, content_w * 0.5, "Единица измерения"),
    ]
    y = _draw_table_header(page, fn, y, sub_cols)
    y = _draw_table_row(page, fn, y, [
        (_MARGIN, content_w * 0.5, _fmt_qty(sub.qty)),
        (_MARGIN + content_w * 0.5, content_w * 0.5, sub.unit),
    ])

    y += 20.0

    # ── 3. Прогноз продаж ──
    page.insert_text((_MARGIN, y), "Прогноз продаж", fontname=fn, fontsize=10, color=_ACCENT)
    y += 16.0

    fc = card_data.forecast
    kv_data = [
        ("Среднемесячные продажи (шт.)", _fmt_qty(fc.avg_monthly)),
        ("Продажи за последние 12 мес. (шт.)", _fmt_qty(fc.last_12m_total)),
        ("Запас коробок (мес.)", _fmt_qty(fc.months_of_stock) if fc.months_of_stock is not None else "—"),
    ]
    kv_label_w = content_w * 0.6
    kv_val_w = content_w * 0.4
    y = _draw_table_header(page, fn, y, [
        (_MARGIN, kv_label_w, "Показатель"),
        (_MARGIN + kv_label_w, kv_val_w, "Значение"),
    ])
    for label_text, val_text in kv_data:
        y = _draw_table_row(page, fn, y, [
            (_MARGIN, kv_label_w, label_text),
            (_MARGIN + kv_label_w, kv_val_w, val_text),
        ])

    y += 20.0

    # ── 4. Рекомендации по заказу ──
    page.insert_text((_MARGIN, y), "Рекомендации по заказу", fontname=fn, fontsize=10, color=_ACCENT)
    y += 16.0

    rec_data = [
        ("Когда заказывать", fc.recommended_order_date or "—"),
        ("Рекомендуемый объём заказа (шт.)", _fmt_qty(fc.recommended_order_qty)),
    ]
    y = _draw_table_header(page, fn, y, [
        (_MARGIN, kv_label_w, "Параметр"),
        (_MARGIN + kv_label_w, kv_val_w, "Значение"),
    ])
    for label_text, val_text in rec_data:
        color = _RED if "сейчас" in val_text.lower() else _DARK
        y = _draw_table_row(page, fn, y, [
            (_MARGIN, kv_label_w, label_text),
            (_MARGIN + kv_label_w, kv_val_w, val_text),
        ], color=color)

    y += 24.0

    # ── Футер ──
    page.insert_text(
        (_MARGIN, y),
        f"GMP: {card_data.gmp_code}" if card_data.gmp_code else "",
        fontname=fn, fontsize=7, color=_LIGHT,
    )


# ── Main entry point ──────────────────────────────────────────────────


def build_product_card_pdf(
    box_row: dict[str, Any],
    related_rows: dict[str, dict[str, Any] | None],
    pdf_root: Path,
    *,
    card_data: ProductCardData | None = None,
) -> bytes:
    """
    Собирает PDF-карточку группы препарата (2 страницы).

    Стр. 1 — векторные макеты упаковки.
    Стр. 2 — остатки, прогноз, рекомендации (если ``card_data`` передан).

    ``related_rows``: ``{"blister": row|None, "label": row|None, "pack": row|None}``
    """
    doc = fitz.open()

    # ── Страница 1: макеты ──
    page = doc.new_page(width=_A4_W, height=_A4_H)
    font_path = _resolve_unicode_font()
    fn = _insert_font(page, font_path)

    gmp = (box_row.get("gmp_code") or "").strip()
    if not gmp:
        gmp = extract_gmp_code(box_row.get("name") or "", box_row.get("file") or "")
    title_name = (box_row.get("name") or "").strip()[:80]
    title_size = (box_row.get("size") or "").strip()
    title_parts = []
    if title_name:
        title_parts.append(title_name)
    if gmp:
        title_parts.append(f"({gmp})")
    title = " ".join(title_parts) or "Карточка препарата"

    page.insert_text(
        (_MARGIN, _MARGIN + 14),
        title,
        fontname=fn, fontsize=11, color=_DARK,
    )
    if title_size:
        page.insert_text(
            (_MARGIN, _MARGIN + 28),
            f"Размер коробки: {title_size}",
            fontname=fn, fontsize=8.5, color=_GRAY,
        )

    body_top = _MARGIN + _HEADER_H
    body_bottom = _A4_H - _MARGIN
    body_h = body_bottom - body_top

    left_w = (_A4_W - 2 * _MARGIN - _GAP) * 0.6
    right_w = (_A4_W - 2 * _MARGIN - _GAP) * 0.4
    right_x = _MARGIN + left_w + _GAP

    left_rect = fitz.Rect(_MARGIN, body_top, _MARGIN + left_w, body_bottom)
    page.draw_rect(left_rect, color=(0.85, 0.85, 0.88), width=0.5)

    box_pdf = _resolve_row_pdf(box_row, pdf_root)
    if box_pdf:
        _show_pdf_vector(page, left_rect, box_pdf)
    else:
        page.insert_text(
            (_MARGIN + 10, body_top + body_h / 2),
            "PDF коробки не найден",
            fontname=fn, fontsize=10, color=_LIGHT,
        )

    side_kinds = ["blister", "label", "pack"]
    n_slots = len(side_kinds)
    slot_label_h = 18.0
    slot_gap = 6.0
    total_gaps = slot_gap * (n_slots - 1) + slot_label_h * n_slots
    slot_img_h = (body_h - total_gaps) / n_slots
    slot_img_h = max(slot_img_h, 40.0)

    y_cursor = body_top
    for kind_key in side_kinds:
        row = related_rows.get(kind_key)
        label = _slot_label(row, kind_key)

        page.insert_text(
            (right_x + 2, y_cursor + 12),
            label,
            fontname=fn, fontsize=8, color=(0.2, 0.2, 0.25),
        )
        y_cursor += slot_label_h

        img_rect = fitz.Rect(right_x, y_cursor, right_x + right_w, y_cursor + slot_img_h)
        page.draw_rect(img_rect, color=(0.88, 0.88, 0.9), width=0.4)

        slot_pdf = _resolve_row_pdf(row, pdf_root)
        if slot_pdf:
            _show_pdf_vector(page, img_rect, slot_pdf)
        else:
            cx = right_x + right_w / 2 - 20
            cy = y_cursor + slot_img_h / 2
            page.insert_text(
                (cx, cy), "нет",
                fontname=fn, fontsize=9, color=(0.65, 0.65, 0.65),
            )

        y_cursor += slot_img_h + slot_gap

    page.draw_line(
        fitz.Point(_MARGIN + left_w + _GAP / 2, body_top),
        fitz.Point(_MARGIN + left_w + _GAP / 2, body_bottom),
        color=(0.82, 0.82, 0.85), width=0.5,
    )

    # ── Страница 2: информационный блок ──
    _render_info_page(doc, font_path, title, card_data)

    data = doc.tobytes(deflate=True, garbage=3)
    doc.close()
    return data
