# -*- coding: utf-8 -*-
from pathlib import Path

import packaging_db as pkg_db
from modules.packaging_catalog.application.makety_kind_merge import merge_kind_values_from_sqlite


def test_merge_kind_missing_db(tmp_path: Path) -> None:
    rows = [{"excel_row": 1, "kind": ""}]
    ch, fixed = merge_kind_values_from_sqlite(rows, tmp_path / "nope.db")
    assert ch is False
    assert fixed == set()


def test_merge_kind_fills_empty_kind(tmp_path: Path) -> None:
    db = tmp_path / "d.db"
    conn = pkg_db.connect(db)
    pkg_db.init_db(conn)
    pkg_db.upsert_all(
        conn,
        [
            {
                "excel_row": 2,
                "name": "N",
                "size": "",
                "kind": "Коробка",
                "file": "x.pdf",
            },
        ],
    )
    conn.close()
    rows = [{"excel_row": 2, "kind": ""}]
    ch, fixed = merge_kind_values_from_sqlite(rows, db)
    assert ch is True
    assert fixed == {2}
    assert rows[0]["kind"] == "Коробка"


def test_kind_locked_not_overwritten(tmp_path: Path) -> None:
    """Закреплённый вид из БД имеет приоритет над Excel."""
    db = tmp_path / "d.db"
    conn = pkg_db.connect(db)
    pkg_db.init_db(conn)
    pkg_db.upsert_all(
        conn,
        [{"excel_row": 3, "name": "N", "size": "", "kind": "Этикетка", "file": "x.pdf"}],
    )
    pkg_db.set_kind_locked(conn, 3, "Этикетка", locked=True)
    conn.close()

    rows = [{"excel_row": 3, "kind": "Коробка"}]
    ch, fixed = merge_kind_values_from_sqlite(rows, db)
    assert ch is True
    assert rows[0]["kind"] == "Этикетка"
    assert rows[0].get("kind_locked") == 1


def test_kind_locked_same_value_no_change(tmp_path: Path) -> None:
    """Закреплённый вид совпадает с Excel — нет изменений, но kind_locked ставится."""
    db = tmp_path / "d.db"
    conn = pkg_db.connect(db)
    pkg_db.init_db(conn)
    pkg_db.upsert_all(
        conn,
        [{"excel_row": 4, "name": "N", "size": "", "kind": "Блистер", "file": "x.pdf"}],
    )
    pkg_db.set_kind_locked(conn, 4, "Блистер", locked=True)
    conn.close()

    rows = [{"excel_row": 4, "kind": "Блистер"}]
    ch, fixed = merge_kind_values_from_sqlite(rows, db)
    assert ch is False
    assert rows[0]["kind"] == "Блистер"
    assert rows[0].get("kind_locked") == 1


def test_kind_locked_upsert_preserves_locked(tmp_path: Path) -> None:
    """upsert_all с kind_locked=0 не сбрасывает закреплённый вид."""
    db = tmp_path / "d.db"
    conn = pkg_db.connect(db)
    pkg_db.init_db(conn)
    pkg_db.upsert_all(
        conn,
        [{"excel_row": 5, "name": "N", "size": "", "kind": "Пакет", "file": "x.pdf", "kind_locked": 1}],
    )
    pkg_db.upsert_all(
        conn,
        [{"excel_row": 5, "name": "N2", "size": "", "kind": "Коробка", "file": "x.pdf"}],
    )
    loaded = pkg_db.load_all(conn)
    conn.close()
    item = next(r for r in loaded if r["excel_row"] == 5)
    assert item["kind"] == "Пакет"
    assert item["kind_locked"] == 1


def test_set_kind_locked_and_load(tmp_path: Path) -> None:
    """set_kind_locked записывает, load_kind_locked_set читает."""
    db = tmp_path / "d.db"
    conn = pkg_db.connect(db)
    pkg_db.init_db(conn)
    pkg_db.upsert_all(
        conn,
        [
            {"excel_row": 10, "name": "A", "size": "", "kind": "Коробка", "file": "a.pdf"},
            {"excel_row": 11, "name": "B", "size": "", "kind": "Этикетка", "file": "b.pdf"},
        ],
    )
    pkg_db.set_kind_locked(conn, 10, "Коробка", locked=True)
    locked = pkg_db.load_kind_locked_set(conn)
    assert 10 in locked
    assert 11 not in locked

    pkg_db.set_kind_locked(conn, 10, "Коробка", locked=False)
    locked2 = pkg_db.load_kind_locked_set(conn)
    assert 10 not in locked2
    conn.close()
