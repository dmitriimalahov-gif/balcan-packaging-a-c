# -*- coding: utf-8 -*-
"""
Фильтрация и поиск строк каталога макетов (без Streamlit).
"""

from __future__ import annotations

from typing import Any

from modules.packaging_catalog.domain.kind_bucket import kind_bucket
from packaging_sizes import row_size_key


def item_matches_text_query(item: dict[str, Any], q_lower: str) -> bool:
    """Подстрока без учёта регистра по полям строки, CG, PDF, № Excel."""
    if not q_lower:
        return True
    parts = [
        item.get("name") or "",
        item.get("file") or "",
        item.get("kind") or "",
        item.get("size") or "",
        item.get("price") or "",
        item.get("price_new") or "",
        item.get("qty_per_sheet") or "",
        item.get("qty_per_year") or "",
        item.get("_cg_knife_name") or "",
        item.get("_cg_category") or "",
        item.get("_cg_lacquers") or "",
        item.get("_cg_cutit_no") or "",
        item.get("_knife_size_mm") or "",
        str(item.get("excel_row") or ""),
    ]
    return q_lower in " ".join(parts).lower()


def item_matches_bucket(item: dict[str, Any], bucket: str) -> bool:
    if bucket == "all":
        return True
    return kind_bucket(item) == bucket


def item_matches_size_key(item: dict[str, Any], key_str: str | None) -> bool:
    if key_str is None:
        return True
    return size_key_str(item) == key_str


def size_key_str(item: dict[str, Any]) -> str:
    """Ключ группы габаритов; перестановки тех же мм совпадают."""
    return row_size_key(item)


def format_size_key_label(key_str: str) -> str:
    """Подпись для кнопки: «80 × 57 mm» или «Без размера»."""
    if key_str == "__empty__":
        return "Без размера"
    parts = [int(x) for x in key_str.split("|")]
    while parts and parts[-1] == 0:
        parts.pop()
    if not parts:
        return "Без размера"
    return " × ".join(str(p) for p in parts) + " mm"
