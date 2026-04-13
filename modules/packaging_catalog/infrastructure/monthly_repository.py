# -*- coding: utf-8 -*-
"""Тонкий слой над `packaging_db` для помесячных объёмов (`packaging_monthly_qty`)."""

from __future__ import annotations

import sqlite3
from typing import Any

import packaging_db as pkg_db


def load_for_excel_rows(
    conn: sqlite3.Connection,
    excel_rows: list[int],
) -> list[dict[str, Any]]:
    return pkg_db.load_monthly_for_rows(conn, excel_rows)


def upsert_batch(
    conn: sqlite3.Connection,
    rows: list[dict[str, Any]],
    *,
    default_source: str = "",
) -> None:
    pkg_db.upsert_monthly_batch(conn, rows, default_source=default_source)
