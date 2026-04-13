# -*- coding: utf-8 -*-
"""
Автосопоставление ножей Cutting Group с продуктами каталога (без Streamlit).
"""

from __future__ import annotations

import re
from typing import Any


def auto_match_cg(
    cg_knives: list[dict[str, Any]],
    rows_by_er: dict[int, dict[str, Any]],
    box_rows: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    """Возвращает список ``{excel_row, cutit_no, confirmed: 0}``."""
    result: list[dict[str, Any]] = []
    matched_ers: set[int] = set()
    for k in cg_knives:
        cg_name = (k.get("name") or "").lower().replace("-", "").replace(" ", "")
        m = re.search(r"\(([^)]+)\)", k.get("name") or "")
        hint = m.group(1).strip().lower().replace("-", "").replace(" ", "") if m else cg_name[:20]
        if len(hint) < 3:
            continue
        best_er: int | None = None
        best_score = 0
        for r in box_rows:
            er = int(r["excel_row"])
            if er in matched_ers:
                continue
            full = rows_by_er.get(er) or r
            db_name = (full.get("name") or "").lower().replace("-", "").replace(" ", "")
            if hint[:8] in db_name or db_name[:10] in hint:
                score = len(set(hint) & set(db_name))
                if score > best_score:
                    best_score = score
                    best_er = er
        if best_er is not None and best_score >= 4:
            result.append(
                {
                    "excel_row": best_er,
                    "cutit_no": k["cutit_no"],
                    "confirmed": 0,
                }
            )
            matched_ers.add(best_er)
    return result
