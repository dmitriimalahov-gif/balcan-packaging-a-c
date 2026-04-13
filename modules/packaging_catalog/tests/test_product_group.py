# -*- coding: utf-8 -*-
from modules.packaging_catalog.application.product_group import (
    find_related_items,
    format_row_label,
)


def _row(er: int, kind: str, name: str = "", gmp: str = "", file: str = "") -> dict:
    return {
        "excel_row": er,
        "kind": kind,
        "name": name,
        "file": file,
        "size": "80x50",
        "gmp_code": gmp,
    }


def test_find_by_gmp() -> None:
    box = _row(1, "Коробка", name="Drug A", gmp="ВУМ-100-01")
    blister = _row(2, "Blister", name="Drug A blister", gmp="ВУМ-100-01")
    label = _row(3, "Этикетка", name="Drug A label", gmp="ВУМ-100-01")
    other_box = _row(4, "Коробка", name="Drug B", gmp="ВУМ-200-02")
    all_rows = [box, blister, label, other_box]

    result = find_related_items(box, all_rows)
    assert len(result["blister"]) == 1
    assert result["blister"][0]["excel_row"] == 2
    assert len(result["label"]) == 1
    assert result["label"][0]["excel_row"] == 3
    assert result["pack"] == []


def test_find_by_name_prefix() -> None:
    box = _row(10, "Коробка", name="Paracetamol 500mg tablets")
    blister = _row(11, "Blister", name="Paracetamol 500mg blister")
    unrelated = _row(12, "Blister", name="Ibuprofen 200mg blister")
    all_rows = [box, blister, unrelated]

    result = find_related_items(box, all_rows)
    assert len(result["blister"]) == 1
    assert result["blister"][0]["excel_row"] == 11


def test_no_self_match() -> None:
    box = _row(1, "Коробка", name="Drug X", gmp="ВУМ-100-01")
    all_rows = [box]
    result = find_related_items(box, all_rows)
    assert all(len(v) == 0 for v in result.values())


def test_find_by_gmp_from_filename() -> None:
    box = _row(1, "Коробка", name="Some drug", file="макет (ВУМ-150-03).pdf")
    label = _row(2, "Этикетка", name="Some drug label", file="этикетка (ВУМ-150-03).pdf")
    all_rows = [box, label]

    result = find_related_items(box, all_rows)
    assert len(result["label"]) == 1


def test_format_row_label() -> None:
    row = _row(5, "Коробка", name="Drug Z", gmp="ВУМ-100-01")
    label = format_row_label(row)
    assert "Row 5" in label
    assert "Drug Z" in label
    assert "Коробка" in label
