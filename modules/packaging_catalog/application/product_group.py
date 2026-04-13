# -*- coding: utf-8 -*-
"""
Поиск связанных позиций каталога для одного препарата (без Streamlit).

По выбранной коробке находит блистеры, этикетки и пакеты с тем же GMP-кодом
и/или похожим названием.
"""

from __future__ import annotations

import re
from typing import Any

from modules.packaging_catalog.domain.kind_bucket import kind_bucket
from packaging_db import extract_gmp_code


def _normalize_name(raw: str) -> str:
    """Нормализация для fuzzy-сравнения: lowercase, без пробелов/дефисов/скобок."""
    s = (raw or "").strip().lower()
    s = re.sub(r"[\s\-_()]+", "", s)
    return s


def _gmp_for_row(row: dict[str, Any]) -> str:
    gmp = (row.get("gmp_code") or "").strip().upper()
    if gmp:
        return gmp
    return extract_gmp_code(
        row.get("name") or "",
        row.get("file") or "",
    )


def find_related_items(
    anchor_row: dict[str, Any],
    all_rows: list[dict[str, Any]],
) -> dict[str, list[dict[str, Any]]]:
    """
    По выбранной позиции (обычно коробке) находит кандидатов для каждого вида.

    Возвращает ``{"blister": [...], "label": [...], "pack": [...]}``.
    Каждый список отсортирован по ``excel_row``.
    """
    anchor_er = int(anchor_row.get("excel_row") or 0)
    anchor_gmp = _gmp_for_row(anchor_row)
    anchor_name = _normalize_name(anchor_row.get("name") or "")
    min_prefix = min(15, max(6, len(anchor_name)))

    result: dict[str, list[dict[str, Any]]] = {
        "blister": [],
        "label": [],
        "pack": [],
    }

    for row in all_rows:
        er = int(row.get("excel_row") or 0)
        if er == anchor_er:
            continue
        bucket = kind_bucket(row)
        if bucket not in result:
            continue

        matched = False
        if anchor_gmp:
            row_gmp = _gmp_for_row(row)
            if row_gmp and row_gmp == anchor_gmp:
                matched = True

        if not matched and anchor_name and len(anchor_name) >= 6:
            row_name = _normalize_name(row.get("name") or "")
            if row_name and len(row_name) >= 6:
                prefix = anchor_name[:min_prefix]
                if row_name.startswith(prefix) or anchor_name.startswith(row_name[:min_prefix]):
                    matched = True

        if matched:
            result[bucket].append(row)

    for bucket in result:
        result[bucket].sort(key=lambda r: int(r.get("excel_row") or 0))

    return result


def format_row_label(row: dict[str, Any]) -> str:
    """Человекочитаемая подпись строки для selectbox."""
    er = row.get("excel_row", "?")
    name = (row.get("name") or "").strip()[:60]
    size = (row.get("size") or "").strip()
    kind = (row.get("kind") or "").strip()
    parts = [f"Row {er}"]
    if name:
        parts.append(name)
    if kind:
        parts.append(kind)
    if size:
        parts.append(size)
    return " · ".join(parts)
