# -*- coding: utf-8 -*-
from modules.packaging_catalog.domain.makety_filters import (
    format_size_key_label,
    item_matches_text_query,
    item_matches_bucket,
)


def test_item_matches_text_query_empty() -> None:
    assert item_matches_text_query({"name": "x"}, "") is True


def test_item_matches_text_query_found() -> None:
    item = {"name": "Cutii ABC", "file": "a.pdf", "kind": "Коробка", "size": "80x50",
            "price": "", "price_new": "", "qty_per_sheet": "", "qty_per_year": "",
            "excel_row": 5}
    assert item_matches_text_query(item, "abc") is True
    assert item_matches_text_query(item, "xyz") is False


def test_item_matches_bucket() -> None:
    item = {"kind": "Коробка"}
    assert item_matches_bucket(item, "all") is True
    assert item_matches_bucket(item, "box") is True
    assert item_matches_bucket(item, "blister") is False


def test_format_size_key_label() -> None:
    assert format_size_key_label("80|50|0") == "80 × 50 mm"
    assert format_size_key_label("__empty__") == "Без размера"
