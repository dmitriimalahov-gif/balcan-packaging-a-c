# -*- coding: utf-8 -*-
from modules.packaging_catalog.application.excel_headers import (
    build_makety_column_index_map,
    excel_header_is_makety_v3,
    excel_row_dict_from_column_map,
    normalize_excel_header_title,
)
from modules.packaging_catalog.domain.makety_excel_config import HEADERS


def test_normalize_excel_header_title() -> None:
    assert normalize_excel_header_title("  Размер  (мм)  ") == normalize_excel_header_title(
        "размер (мм)"
    )


def test_build_makety_column_index_map_permutation() -> None:
    """Эталонные заголовки в обратном порядке — карта всё рава находится."""
    hdr = tuple(reversed(HEADERS))
    m = build_makety_column_index_map(hdr)
    assert m is not None
    assert len(m) == 13
    assert sorted(m) == list(range(13))


def test_build_makety_column_index_map_incomplete() -> None:
    assert build_makety_column_index_map(tuple(HEADERS[:5])) is None


def test_excel_row_dict_from_column_map() -> None:
    col_map = list(range(13))
    row = tuple(f"v{i}" for i in range(13))
    d = excel_row_dict_from_column_map(5, row, col_map)
    assert d["excel_row"] == 5
    assert d["name"] == "v12"
    assert d["size"] == "v3"


def test_excel_header_is_makety_v3() -> None:
    h = [""] * 8
    h[7] = "Нож CG"
    assert excel_header_is_makety_v3(tuple(h)) is True
    assert excel_header_is_makety_v3(tuple()) is False
