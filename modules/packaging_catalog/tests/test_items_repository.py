# -*- coding: utf-8 -*-
from pathlib import Path

import packaging_db as pkg_db
from modules.packaging_catalog.infrastructure import items_repository, monthly_repository


def test_items_and_monthly_repositories(tmp_path: Path) -> None:
    db = tmp_path / "t.db"
    conn = pkg_db.connect(db)
    items_repository.upsert_many(
        conn,
        [{"excel_row": 1, "name": "A", "size": "", "kind": "", "file": "x.pdf"}],
    )
    assert items_repository.count_rows(conn) == 1
    rows = items_repository.fetch_all(conn)
    assert rows[0]["excel_row"] == 1

    monthly_repository.upsert_batch(
        conn,
        [{"excel_row": 1, "year": 2025, "month": 1, "qty": 10.0}],
    )
    m = monthly_repository.load_for_excel_rows(conn, [1])
    assert len(m) == 1
    assert m[0]["qty"] == 10.0
    conn.close()
