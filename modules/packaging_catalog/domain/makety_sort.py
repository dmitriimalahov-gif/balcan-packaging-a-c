# -*- coding: utf-8 -*-
"""
Сортировка строк каталога макетов (без Streamlit).
"""

from __future__ import annotations

from typing import Any

from packaging_sizes import parse_box_dimensions_mm


def sort_rows(
    items: list[dict[str, Any]],
    by: str,
    reverse: bool,
) -> list[dict[str, Any]]:
    if by == "По виду":
        return sorted(
            items,
            key=lambda r: (r.get("kind") or "").lower(),
            reverse=reverse,
        )
    if by in ("По размеру (габариты мм)", "По размеру"):

        def size_key(r: dict[str, Any]) -> tuple[float, ...]:
            return parse_box_dimensions_mm(r.get("size") or "")

        return sorted(items, key=size_key, reverse=reverse)
    if by == "По названию":
        return sorted(
            items,
            key=lambda r: (
                (r.get("_cg_knife_name") or r.get("name") or "").lower(),
            ),
            reverse=reverse,
        )
    if by == "По PDF":
        return sorted(
            items,
            key=lambda r: (r.get("file") or "").lower(),
            reverse=reverse,
        )
    if by == "По строке Excel":
        return sorted(
            items,
            key=lambda r: int(r.get("excel_row") or 0),
            reverse=reverse,
        )
    return list(items)
