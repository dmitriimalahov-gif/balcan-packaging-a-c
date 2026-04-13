# -*- coding: utf-8 -*-
"""Тесты для product_card_data: сбор остатков, прогноз, рекомендации."""

from __future__ import annotations

import sqlite3
import tempfile
from pathlib import Path

import pytest

import packaging_db as pkg_db
from modules.packaging_catalog.application.product_card_data import (
    ProductCardData,
    SubstanceStockInfo,
    _calc_avg_monthly,
    _calc_forecast,
    collect_product_card_data,
)


@pytest.fixture()
def tmp_db(tmp_path: Path) -> Path:
    db_file = tmp_path / "test.db"
    conn = pkg_db.connect(db_file)
    pkg_db.init_db(conn)
    conn.close()
    return db_file


def test_calc_avg_monthly_empty():
    assert _calc_avg_monthly([]) == 0.0


def test_calc_avg_monthly_with_data():
    from datetime import date

    today = date.today()
    rows = [
        {"year": today.year, "month": today.month, "qty": 100},
        {"year": today.year, "month": max(today.month - 1, 1), "qty": 200},
    ]
    avg = _calc_avg_monthly(rows)
    assert avg > 0


def test_calc_forecast_no_sales():
    fc = _calc_forecast(0, 100, 0)
    assert fc.avg_monthly == 0
    assert fc.recommended_order_qty == 0


def test_calc_forecast_with_sales():
    fc = _calc_forecast(100, 500, 1200)
    assert fc.avg_monthly == 100
    assert fc.months_of_stock == 5.0
    assert fc.recommended_order_qty == 300


def test_calc_forecast_order_now():
    fc = _calc_forecast(100, 50, 1200)
    assert "сейчас" in fc.recommended_order_date.lower()


def test_collect_no_db():
    box_row = {"name": "Test", "gmp_code": "GMP-001", "size": "100x50"}
    data = collect_product_card_data(None, "GMP-001", box_row, {}, [])
    assert isinstance(data, ProductCardData)
    assert len(data.packaging_stock) == 4
    assert data.substance.qty == 0.0


def test_collect_with_db(tmp_db: Path):
    conn = pkg_db.connect(tmp_db)
    pkg_db.init_db(conn)
    pkg_db.upsert_packaging_stock(conn, "GMP-001", "box", 500, source="test")
    pkg_db.upsert_packaging_stock(conn, "GMP-001", "label", 200, source="test")
    pkg_db.upsert_substance_stock(conn, "GMP-001", 25.5, "кг", source="test")
    conn.close()

    box_row = {"name": "Test", "gmp_code": "GMP-001", "size": "100x50", "excel_row": 1}
    data = collect_product_card_data(
        str(tmp_db), "GMP-001", box_row, {}, [box_row],
    )
    assert data.gmp_code == "GMP-001"
    box_stock = next(ps for ps in data.packaging_stock if ps.kind == "box")
    assert box_stock.qty == 500.0
    label_stock = next(ps for ps in data.packaging_stock if ps.kind == "label")
    assert label_stock.qty == 200.0
    assert data.substance.qty == 25.5
    assert data.substance.unit == "кг"


def test_packaging_stock_crud(tmp_db: Path):
    conn = pkg_db.connect(tmp_db)
    pkg_db.init_db(conn)

    pkg_db.upsert_packaging_stock(conn, "GMP-002", "box", 100)
    pkg_db.upsert_packaging_stock(conn, "GMP-002", "blister", 300)

    stock = pkg_db.load_packaging_stock_for_gmp(conn, "GMP-002")
    assert stock["box"] == 100.0
    assert stock["blister"] == 300.0

    pkg_db.upsert_packaging_stock(conn, "GMP-002", "box", 150)
    stock2 = pkg_db.load_packaging_stock_for_gmp(conn, "GMP-002")
    assert stock2["box"] == 150.0

    conn.close()


def test_substance_stock_crud(tmp_db: Path):
    conn = pkg_db.connect(tmp_db)
    pkg_db.init_db(conn)

    pkg_db.upsert_substance_stock(conn, "GMP-003", 10.5, "мл")
    sub = pkg_db.load_substance_stock(conn, "GMP-003")
    assert sub["GMP-003"]["qty"] == 10.5
    assert sub["GMP-003"]["unit"] == "мл"

    pkg_db.upsert_substance_stock(conn, "GMP-003", 20.0, "л")
    sub2 = pkg_db.load_substance_stock(conn, "GMP-003")
    assert sub2["GMP-003"]["qty"] == 20.0
    assert sub2["GMP-003"]["unit"] == "л"

    conn.close()


def test_substance_stock_invalid_unit(tmp_db: Path):
    conn = pkg_db.connect(tmp_db)
    pkg_db.init_db(conn)

    pkg_db.upsert_substance_stock(conn, "GMP-004", 5.0, "фунты")
    sub = pkg_db.load_substance_stock(conn, "GMP-004")
    assert sub["GMP-004"]["unit"] == "кг"

    conn.close()
