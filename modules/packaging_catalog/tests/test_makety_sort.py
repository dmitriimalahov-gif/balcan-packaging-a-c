# -*- coding: utf-8 -*-
from modules.packaging_catalog.domain.makety_sort import sort_rows


def test_sort_by_kind() -> None:
    rows = [{"kind": "Пакет"}, {"kind": "Blister"}, {"kind": "Коробка"}]
    result = sort_rows(rows, "По виду", False)
    assert [r["kind"] for r in result] == ["Blister", "Коробка", "Пакет"]


def test_sort_by_excel_row() -> None:
    rows = [{"excel_row": 5}, {"excel_row": 2}, {"excel_row": 10}]
    result = sort_rows(rows, "По строке Excel", False)
    assert [r["excel_row"] for r in result] == [2, 5, 10]


def test_sort_unknown_key_returns_copy() -> None:
    rows = [{"kind": "A"}, {"kind": "B"}]
    result = sort_rows(rows, "Неизвестная", False)
    assert result == rows
    assert result is not rows
