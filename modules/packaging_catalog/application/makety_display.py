# -*- coding: utf-8 -*-
"""Чистые функции отображения полей макетов (без Streamlit)."""

from __future__ import annotations

import math


def parse_qty_int_for_cg(val: str) -> int:
    if not val:
        return 0
    cleaned = val.replace(" ", "").replace("\u00a0", "").replace(",", ".")
    try:
        return int(float(cleaned))
    except (ValueError, TypeError):
        return 0


def format_qty_year_caption(raw: str | None) -> str:
    """Одна строка для UI: годовой объём заказа коробок (столбец qty_per_year)."""
    if raw is None or not str(raw).strip():
        return "Заказ/год: —"
    cleaned = str(raw).strip().replace(" ", "").replace("\u00a0", "").replace(",", ".")
    try:
        v = float(cleaned)
    except (ValueError, TypeError):
        t = str(raw).strip()
        return f"Заказ/год: {t}" if len(t) <= 28 else f"Заказ/год: {t[:25]}…"
    if not math.isfinite(v):
        return "Заказ/год: —"
    if abs(v - round(v)) < 1e-9:
        n = int(round(v))
        s = str(abs(n))
        chunks = [s[max(0, i - 3) : i] for i in range(len(s), 0, -3)][::-1]
        grouped = " ".join(chunks)
        if n < 0:
            grouped = "-" + grouped
        return f"Заказ/год: {grouped} шт"
    return f"Заказ/год: {v:g} шт"
