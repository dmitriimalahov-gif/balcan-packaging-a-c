#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Одноразовый перенос данных SQLite (packaging_db) → PostgreSQL после alembic upgrade.

Пример:
  export PACKAGING_DATABASE_URL=postgresql+psycopg://user:pass@localhost/packaging
  python scripts/migrate_sqlite_to_postgres.py /path/to/packaging_data.db

Требования: применены миграции Alembic; целевая БД может быть пустой по данным
(скрипт делает INSERT; при повторном запуске возможны конфликты уникальных ключей).
"""

from __future__ import annotations

import argparse
import sqlite3
import sys
from datetime import datetime
from pathlib import Path


def _parse_ts(raw: str | None) -> datetime:
    if not raw:
        return datetime.now().astimezone()
    raw = raw.strip()
    if raw.endswith("Z"):
        raw = raw[:-1] + "+00:00"
    try:
        return datetime.fromisoformat(raw)
    except ValueError:
        return datetime.now().astimezone()


def main() -> int:
    parser = argparse.ArgumentParser(description="SQLite → Postgres для каталога упаковки")
    parser.add_argument("sqlite_path", type=Path, help="Путь к .db (например packaging_data.db)")
    args = parser.parse_args()
    if not args.sqlite_path.is_file():
        print("Файл SQLite не найден:", args.sqlite_path, file=sys.stderr)
        return 1

    try:
        from sqlalchemy import create_engine, text
    except ImportError:
        print("Установите зависимости: pip install -r requirements-api.txt", file=sys.stderr)
        return 1

    import os

    url = (os.environ.get("PACKAGING_DATABASE_URL") or "").strip()
    if not url.startswith("postgresql"):
        print("Задайте PACKAGING_DATABASE_URL=postgresql+psycopg://...", file=sys.stderr)
        return 1

    engine = create_engine(url, pool_pre_ping=True)
    src = sqlite3.connect(str(args.sqlite_path))
    src.row_factory = sqlite3.Row

    with engine.begin() as conn:
        # 1) Каталог
        cur = src.execute(
            """
            SELECT excel_row, name, size, kind, pdf_file, price, price_new,
                   qty_per_sheet, qty_per_year, gmp_code, updated_at
            FROM packaging_items
            """
        )
        for row in cur.fetchall():
            conn.execute(
                text(
                    """
                    INSERT INTO packaging_items (
                        excel_row, name, size, kind, pdf_file, price, price_new,
                        qty_per_sheet, qty_per_year, gmp_code, updated_at
                    ) VALUES (
                        :excel_row, :name, :size, :kind, :pdf_file, :price, :price_new,
                        :qty_per_sheet, :qty_per_year, :gmp_code, :updated_at
                    )
                    ON CONFLICT (excel_row) DO UPDATE SET
                        name = EXCLUDED.name,
                        size = EXCLUDED.size,
                        kind = EXCLUDED.kind,
                        pdf_file = EXCLUDED.pdf_file,
                        price = EXCLUDED.price,
                        price_new = EXCLUDED.price_new,
                        qty_per_sheet = EXCLUDED.qty_per_sheet,
                        qty_per_year = EXCLUDED.qty_per_year,
                        gmp_code = EXCLUDED.gmp_code,
                        updated_at = EXCLUDED.updated_at
                    """
                ),
                {
                    "excel_row": int(row["excel_row"]),
                    "name": row["name"],
                    "size": row["size"],
                    "kind": row["kind"],
                    "pdf_file": row["pdf_file"],
                    "price": row["price"],
                    "price_new": row["price_new"],
                    "qty_per_sheet": row["qty_per_sheet"],
                    "qty_per_year": row["qty_per_year"],
                    "gmp_code": row["gmp_code"],
                    "updated_at": _parse_ts(row["updated_at"]),
                },
            )

        def _table_exists(name: str) -> bool:
            r = src.execute(
                "SELECT 1 FROM sqlite_master WHERE type='table' AND name=?",
                (name,),
            )
            return r.fetchone() is not None

        # 2) Помесячные объёмы
        if _table_exists("packaging_monthly_qty"):
            cur = src.execute(
                "SELECT excel_row, year, month, qty, source, updated_at FROM packaging_monthly_qty"
            )
            for row in cur.fetchall():
                conn.execute(
                    text(
                        """
                        INSERT INTO packaging_monthly_qty (
                            excel_row, year, month, qty, source, updated_at
                        ) VALUES (
                            :excel_row, :year, :month, :qty, :source, :updated_at
                        )
                        ON CONFLICT (excel_row, year, month) DO UPDATE SET
                            qty = EXCLUDED.qty,
                            source = COALESCE(EXCLUDED.source, packaging_monthly_qty.source),
                            updated_at = EXCLUDED.updated_at
                        """
                    ),
                    {
                        "excel_row": int(row["excel_row"]),
                        "year": int(row["year"]),
                        "month": int(row["month"]),
                        "qty": float(row["qty"]),
                        "source": row["source"],
                        "updated_at": _parse_ts(row["updated_at"]),
                    },
                )

        # 3) Остальные таблицы (если есть в SQLite)
        def copy_simple(
            table: str,
            cols: list[str],
            conflict: str | None = None,
        ) -> None:
            if not _table_exists(table):
                return
            cur2 = src.execute(f"SELECT {', '.join(cols)} FROM {table}")
            for row in cur2.fetchall():
                placeholders = ", ".join(f":{c}" for c in cols)
                q = f"INSERT INTO {table} ({', '.join(cols)}) VALUES ({placeholders})"
                if conflict:
                    q += f" {conflict}"
                params = {c: row[c] for c in cols}
                if "updated_at" in params and isinstance(params["updated_at"], str):
                    params["updated_at"] = _parse_ts(params["updated_at"])
                if "confirmed" in params:
                    params["confirmed"] = bool(int(params["confirmed"] or 0))
                conn.execute(text(q), params)

        copy_simple(
            "cutii_confirmations",
            ["cutii_sheet_row", "confirmed_excel_row", "cutii_name", "updated_at"],
            "ON CONFLICT (cutii_sheet_row) DO NOTHING",
        )
        copy_simple(
            "print_tariffs",
            ["min_sheets", "max_sheets", "price_per_sheet", "updated_at"],
        )
        copy_simple(
            "print_finish_extras",
            ["code", "label", "extra_per_sheet", "updated_at"],
            "ON CONFLICT (code) DO NOTHING",
        )
        copy_simple(
            "knife_cache",
            ["excel_row", "svg_full", "width_mm", "height_mm", "pdf_file", "updated_at"],
            "ON CONFLICT (excel_row) DO UPDATE SET svg_full = EXCLUDED.svg_full, "
            "width_mm = EXCLUDED.width_mm, height_mm = EXCLUDED.height_mm, "
            "pdf_file = EXCLUDED.pdf_file, updated_at = EXCLUDED.updated_at",
        )
        copy_simple(
            "stock_on_hand",
            ["gmp_code", "qty", "source", "updated_at"],
            "ON CONFLICT (gmp_code) DO UPDATE SET qty = EXCLUDED.qty, "
            "source = EXCLUDED.source, updated_at = EXCLUDED.updated_at",
        )
        copy_simple(
            "cg_knives",
            ["cutit_no", "name", "category", "cardboard", "updated_at"],
            "ON CONFLICT (cutit_no) DO UPDATE SET name = EXCLUDED.name, "
            "category = EXCLUDED.category, cardboard = EXCLUDED.cardboard, "
            "updated_at = EXCLUDED.updated_at",
        )
        copy_simple(
            "cg_prices",
            [
                "cutit_no",
                "finish_type",
                "min_qty",
                "max_qty",
                "price_per_1000",
                "price_old_per_1000",
                "updated_at",
            ],
            "ON CONFLICT (cutit_no, finish_type, min_qty) DO UPDATE SET "
            "max_qty = EXCLUDED.max_qty, price_per_1000 = EXCLUDED.price_per_1000, "
            "price_old_per_1000 = EXCLUDED.price_old_per_1000, "
            "updated_at = EXCLUDED.updated_at",
        )
        copy_simple(
            "cg_mapping",
            ["excel_row", "cutit_no", "confirmed", "updated_at"],
            "ON CONFLICT (excel_row) DO UPDATE SET cutit_no = EXCLUDED.cutit_no, "
            "confirmed = EXCLUDED.confirmed, updated_at = EXCLUDED.updated_at",
        )

    src.close()
    print("Миграция данных завершена.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
