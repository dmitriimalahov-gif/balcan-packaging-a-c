#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Генерирует XLSX-шаблоны импорта в templates/import/ (нужен openpyxl)."""

from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
OUT_DIR = ROOT / "templates" / "import"


def main() -> int:
    try:
        from openpyxl import Workbook
    except ImportError:
        print("Установите openpyxl: pip install openpyxl", file=sys.stderr)
        return 1

    OUT_DIR.mkdir(parents=True, exist_ok=True)

    wb = Workbook()
    ws = wb.active
    ws.title = "makety"
    headers = [
        "excel_row",
        "name",
        "size",
        "kind",
        "file",
        "price",
        "price_new",
        "qty_per_sheet",
        "qty_per_year",
        "gmp_code",
    ]
    ws.append(headers)
    ws.append(
        [
            2,
            "Пример (удалить строку)",
            "100x50x20",
            "Коробка",
            "example.pdf",
            "",
            "",
            "4",
            "1000",
            "",
        ]
    )

    path = OUT_DIR / "makety_v1.xlsx"
    wb.save(path)
    print("Записано:", path)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
