# -*- coding: utf-8 -*-
"""Тонкий слой над `packaging_db` для таблицы позиций каталога (SQLite)."""

from __future__ import annotations

import sqlite3
from typing import Any

import packaging_db as pkg_db


def init_schema(conn: sqlite3.Connection) -> None:
    pkg_db.init_db(conn)


def fetch_all(conn: sqlite3.Connection) -> list[dict[str, Any]]:
    init_schema(conn)
    return pkg_db.load_all(conn)


def count_rows(conn: sqlite3.Connection) -> int:
    init_schema(conn)
    return pkg_db.row_count(conn)


def upsert_many(conn: sqlite3.Connection, rows: list[dict[str, Any]]) -> None:
    init_schema(conn)
    pkg_db.upsert_all(conn, rows)
