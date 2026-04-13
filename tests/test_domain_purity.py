# -*- coding: utf-8 -*-
"""Тесты чистого домена (без Streamlit)."""

from __future__ import annotations

import tempfile
from pathlib import Path

import pytest

import packaging_db as pkg_db
from packaging_color_analysis import canonical_color_analytics_bucket
from packaging_facade import count_packaging_rows, load_packaging_catalog
from packaging_print_planning import parse_qty_per_sheet


def test_extract_gmp_from_parentheses() -> None:
    assert pkg_db.extract_gmp_code("Prod (ВУМ-169-01)", "") == "ВУМ-169-01"


def test_extract_gmp_from_filename() -> None:
    assert pkg_db.extract_gmp_code("", "foo_ВУМ-170-02_bar.pdf") == "ВУМ-170-02"


def test_parse_qty_per_sheet() -> None:
    assert parse_qty_per_sheet("12") == 12
    assert parse_qty_per_sheet("") is None


@pytest.mark.parametrize(
    "kind,expected",
    [
        ("Блистер", "blister"),
        ("Коробка", "box"),
        ("Пакет", "pack"),
        ("Этикетка", "label"),
    ],
)
def test_canonical_color_bucket(kind: str, expected: str) -> None:
    assert canonical_color_analytics_bucket({"kind": kind}) == expected


def test_facade_empty_db_path() -> None:
    assert load_packaging_catalog(Path("/nonexistent/packaging.db")) == []
    assert count_packaging_rows(Path("/nonexistent/packaging.db")) == 0


def test_facade_minimal_db() -> None:
    with tempfile.TemporaryDirectory() as td:
        p = Path(td) / "t.db"
        conn = pkg_db.connect(p)
        try:
            pkg_db.init_db(conn)
            pkg_db.upsert_all(
                conn,
                [
                    {
                        "excel_row": 5,
                        "name": "Test",
                        "size": "100x50x20",
                        "kind": "Коробка",
                        "file": "x.pdf",
                        "price": "",
                        "price_new": "",
                        "qty_per_sheet": "4",
                        "qty_per_year": "1000",
                    },
                ],
            )
        finally:
            conn.close()
        rows = load_packaging_catalog(p)
        assert len(rows) == 1
        assert int(rows[0]["excel_row"]) == 5
        assert count_packaging_rows(p) == 1
