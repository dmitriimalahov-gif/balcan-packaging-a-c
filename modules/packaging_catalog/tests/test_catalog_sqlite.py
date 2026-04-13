# -*- coding: utf-8 -*-
from pathlib import Path

import packaging_db as pkg_db
from modules.packaging_catalog.infrastructure.catalog_sqlite import (
    count_packaging_rows,
    load_packaging_catalog,
)


def test_sqlite_catalog_roundtrip(tmp_path: Path) -> None:
    db = tmp_path / "x.db"
    conn = pkg_db.connect(db)
    pkg_db.init_db(conn)
    pkg_db.upsert_all(
        conn,
        [{"excel_row": 1, "name": "A", "size": "", "kind": "", "file": "a.pdf"}],
    )
    conn.close()

    rows = load_packaging_catalog(db)
    assert len(rows) == 1
    assert rows[0]["excel_row"] == 1
    assert count_packaging_rows(db) == 1
