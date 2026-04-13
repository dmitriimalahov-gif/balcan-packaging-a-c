# -*- coding: utf-8 -*-
"""
Заголовки и карта столбцов листа Excel «Макеты» (без openpyxl / Streamlit).
"""

from __future__ import annotations

import re
from typing import Any

from modules.packaging_catalog.domain.makety_excel_config import HEADERS


def excel_cell_str(row: tuple[Any, ...] | None, idx: int) -> str:
    if row is None or len(row) <= idx:
        return ""
    v = row[idx]
    return "" if v is None else str(v).strip()


def normalize_excel_header_title(val: Any) -> str:
    """Сравнение заголовков: без лишних пробелов, без различия регистра."""
    if val is None:
        return ""
    s = str(val).replace("\u00a0", " ").strip()
    s = re.sub(r"\s+", " ", s)
    return s.casefold()


def excel_header_is_makety_v3(hdr_t: tuple[Any, ...]) -> bool:
    """Столбец H (индекс 7) — «Нож CG» в актуальном макете Excel."""
    if len(hdr_t) <= 7:
        return False
    h = excel_cell_str(hdr_t, 7).lower().replace(" ", "")
    return "нож" in h and "cg" in h


def build_makety_column_index_map(hdr_t: tuple[Any, ...]) -> list[int] | None:
    """
    Для каждого из 13 эталонных заголовков HEADERS — индекс столбца в файле (0-based).
    Порядок столбцов в файле может отличаться от эталона, имена должны совпадать попарно.
    """
    norm_expected = [normalize_excel_header_title(h) for h in HEADERS]
    norm_actual = [normalize_excel_header_title(x) for x in hdr_t]
    used: set[int] = set()
    out: list[int] = []
    for exp in norm_expected:
        found: int | None = None
        for j, act in enumerate(norm_actual):
            if j in used or act != exp:
                continue
            found = j
            break
        if found is None:
            return None
        used.add(found)
        out.append(found)
    return out


def excel_row_dict_from_column_map(
    excel_row: int,
    row_t: tuple[Any, ...],
    col_map: list[int],
) -> dict[str, Any]:
    """Одна строка данных по карте столбцов (индексы в порядке HEADERS)."""
    g = col_map
    return {
        "excel_row": excel_row,
        "name": excel_cell_str(row_t, g[12]),
        "size": excel_cell_str(row_t, g[3]),
        "kind": excel_cell_str(row_t, g[4]),
        "file": excel_cell_str(row_t, g[5]),
        "price": excel_cell_str(row_t, g[8]),
        "price_new": excel_cell_str(row_t, g[9]),
        "qty_per_sheet": excel_cell_str(row_t, g[10]),
        "qty_per_year": excel_cell_str(row_t, g[11]),
    }
