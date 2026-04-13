# -*- coding: utf-8 -*-
"""
Подстановка поля «Вид» из SQLite после чтения Excel (без Streamlit и без записи файла).
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

import packaging_db as pkg_db


def merge_kind_values_from_sqlite(
    rows: list[dict[str, Any]],
    db_path: Path,
    *,
    overwrite_nonempty_excel: bool = False,
) -> tuple[bool, set[int]]:
    """
    Для совпадающих ``excel_row`` подставляет ``kind`` из БД.

    Строки с ``kind_locked = 1`` в БД **всегда** получают вид из БД
    (закреплённый вид имеет приоритет над Excel).

    Возвращает ``(были_изменения, множество_excel_row_где_подставили_вид)``.
    """
    if not db_path.is_file():
        return False, set()
    try:
        conn = pkg_db.connect(db_path)
        try:
            pkg_db.init_db(conn)
            if pkg_db.row_count(conn) == 0:
                return False, set()
            db_rows = pkg_db.load_all(conn)
        finally:
            conn.close()
    except Exception:
        return False, set()

    by_er = {int(r["excel_row"]): r for r in db_rows}
    any_kind_change = False
    kind_fixed_rows: set[int] = set()
    for item in rows:
        br = by_er.get(int(item["excel_row"]))
        if br is None:
            continue
        k = (br.get("kind") or "").strip()
        if not k:
            continue
        is_locked = bool(br.get("kind_locked"))
        prev = (item.get("kind") or "").strip()

        if is_locked:
            if k != prev:
                item["kind"] = k
                item["kind_locked"] = 1
                any_kind_change = True
                kind_fixed_rows.add(int(item["excel_row"]))
            else:
                item["kind_locked"] = 1
            continue

        if not overwrite_nonempty_excel:
            if prev:
                continue
        elif k == prev:
            continue
        item["kind"] = k
        any_kind_change = True
        kind_fixed_rows.add(int(item["excel_row"]))
    return any_kind_change, kind_fixed_rows
