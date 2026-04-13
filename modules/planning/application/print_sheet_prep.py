# -*- coding: utf-8 -*-
"""
Подготовка строк для раскладки печатного листа.

Делегирует [`packaging_print_planning`](../../../packaging_print_planning.py), чтобы UI/API
не импортировали тяжёлый модуль напрямую из точек входа Streamlit.
"""

from __future__ import annotations

from typing import Any

from packaging_print_planning import (
    printable_rows_only,
    sheet_layout_candidate_rows,
)

__all__ = ["printable_rows_only", "sheet_layout_candidate_rows", "prepare_sheet_candidates"]


def prepare_sheet_candidates(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Фильтр строк, пригодных для построения кандидатов раскладки."""
    return sheet_layout_candidate_rows(printable_rows_only(rows))
