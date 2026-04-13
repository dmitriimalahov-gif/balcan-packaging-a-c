# -*- coding: utf-8 -*-
from modules.packaging_catalog.application.makety_display import (
    format_qty_year_caption,
    parse_qty_int_for_cg,
)


def test_parse_qty_int_for_cg() -> None:
    assert parse_qty_int_for_cg("") == 0
    assert parse_qty_int_for_cg("1 000") == 1000
    assert parse_qty_int_for_cg("12,5") == 12


def test_format_qty_year_caption() -> None:
    assert "—" in format_qty_year_caption(None)
    assert "шт" in format_qty_year_caption("1000")
    assert format_qty_year_caption("abc") == "Заказ/год: abc"
