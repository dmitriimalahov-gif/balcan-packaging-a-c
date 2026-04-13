#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Обновить gmp_code в SQLite по именам PDF (packaging_items.pdf_file)."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

import packaging_db as pkg_db

ROOT = Path(__file__).resolve().parent


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument(
        "--db",
        type=Path,
        default=pkg_db.DEFAULT_DB_PATH,
        help=f"Путь к packaging_data.db (по умолчанию: {pkg_db.DEFAULT_DB_PATH})",
    )
    args = ap.parse_args()
    db_path = args.db.expanduser().resolve()
    if not db_path.is_file():
        print(f"БД не найдена: {db_path}", file=sys.stderr)
        return 1
    conn = pkg_db.connect(db_path)
    try:
        pkg_db.init_db(conn)
        if pkg_db.row_count(conn) == 0:
            print("В БД нет строк.")
            return 0
        u, same, skip = pkg_db.sync_gmp_from_pdf_filenames(conn)
        print(
            f"Обновлено записей: {u} · уже совпадало с именем PDF: {same} · "
            f"без распознаваемого кода в имени файла: {skip}"
        )
    finally:
        conn.close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
