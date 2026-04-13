# -*- coding: utf-8 -*-
"""Чтение каталога из SQLite (packaging_db)."""

from __future__ import annotations

from pathlib import Path
from typing import Any

import packaging_db as pkg_db

from modules.packaging_catalog.infrastructure import items_repository


def load_packaging_catalog(db_path: Path) -> list[dict[str, Any]]:
    """Все строки каталога (инициализация схемы при необходимости)."""
    if not db_path.is_file():
        return []
    conn = pkg_db.connect(db_path)
    try:
        return items_repository.fetch_all(conn)
    finally:
        conn.close()


def count_packaging_rows(db_path: Path) -> int:
    if not db_path.is_file():
        return 0
    conn = pkg_db.connect(db_path)
    try:
        return items_repository.count_rows(conn)
    finally:
        conn.close()
