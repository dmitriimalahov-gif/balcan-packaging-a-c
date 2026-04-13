# -*- coding: utf-8 -*-
from modules.packaging_catalog.application.cg_pret_import import (
    clean_price_token,
    parse_cg_price,
)


def test_clean_price_token_basic() -> None:
    assert clean_price_token("12.50") == 12.5
    assert clean_price_token("(old) 15.00 / 1000") == 15.0
    assert clean_price_token("abc") is None
    assert clean_price_token("0.1") is None


def test_parse_cg_price_none() -> None:
    assert parse_cg_price(None) == (None, None)


def test_parse_cg_price_numeric() -> None:
    assert parse_cg_price(25.0) == (25.0, 25.0)
    assert parse_cg_price(0) == (None, None)


def test_parse_cg_price_two_lines() -> None:
    old, new = parse_cg_price("18.50\n22.00")
    assert old == 18.5
    assert new == 22.0


def test_parse_cg_price_single() -> None:
    old, new = parse_cg_price("30.00")
    assert old == 30.0
    assert new == 30.0
