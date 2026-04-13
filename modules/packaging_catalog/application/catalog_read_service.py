# -*- coding: utf-8 -*-
"""Сценарий чтения каталога: выбор SQLite или Postgres по окружению."""

from __future__ import annotations

import os
from pathlib import Path
from typing import Any

import packaging_db as pkg_db

from modules.packaging_catalog.infrastructure import catalog_postgres, catalog_sqlite


def _use_postgres() -> bool:
    url = (os.environ.get("PACKAGING_DATABASE_URL") or "").strip().lower()
    return url.startswith("postgresql")


def get_catalog_items() -> list[dict[str, Any]]:
    if _use_postgres():
        return catalog_postgres.list_packaging_items()
    path = Path(os.environ.get("PACKAGING_DB_PATH", str(pkg_db.DEFAULT_DB_PATH)))
    return catalog_sqlite.load_packaging_catalog(path)


def get_catalog_count() -> int:
    if _use_postgres():
        return catalog_postgres.count_packaging_items()
    path = Path(os.environ.get("PACKAGING_DB_PATH", str(pkg_db.DEFAULT_DB_PATH)))
    return catalog_sqlite.count_packaging_rows(path)
