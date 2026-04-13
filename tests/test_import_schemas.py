# -*- coding: utf-8 -*-
from packaging_schemas import MaketyImportRow, validate_makety_rows


def test_makety_row_gmp_autofill():
    r = MaketyImportRow(
        excel_row=3,
        name="X (ВУМ-999-01)",
        file="",
    )
    d = r.to_upsert_dict()
    assert d["gmp_code"] == "ВУМ-999-01"


def test_validate_makety_rows_mixed():
    ok, err = validate_makety_rows(
        [
            {"excel_row": 1, "name": "A"},
            {"excel_row": 0, "name": "bad"},
        ]
    )
    assert len(ok) == 1
    assert ok[0]["excel_row"] == 1
    assert len(err) >= 1
    assert any(e.index == 1 for e in err)
