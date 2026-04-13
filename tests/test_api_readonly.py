# -*- coding: utf-8 -*-
from pathlib import Path

import pytest
from fastapi.testclient import TestClient

import packaging_db as pkg_db


@pytest.fixture()
def api_client(tmp_path: Path, monkeypatch: pytest.MonkeyPatch):
    db = tmp_path / "t.db"
    monkeypatch.setenv("PACKAGING_DB_PATH", str(db))
    monkeypatch.delenv("PACKAGING_DATABASE_URL", raising=False)
    conn = pkg_db.connect(db)
    pkg_db.init_db(conn)
    pkg_db.upsert_all(
        conn,
        [
            {
                "excel_row": 2,
                "name": "N",
                "size": "1x1x1",
                "kind": "K",
                "file": "f.pdf",
                "gmp_code": "ВУМ-001-01",
            },
        ],
    )
    conn.close()

    from api.main import app

    return TestClient(app)


def test_health():
    from api.main import app

    c = TestClient(app)
    assert c.get("/health").json() == {"status": "ok"}


def test_items_sqlite(api_client: TestClient):
    r = api_client.get("/api/v1/items")
    assert r.status_code == 200
    data = r.json()
    assert data["total"] == 1
    assert len(data["items"]) == 1
    row = data["items"][0]
    assert row["excel_row"] == 2
    assert row["gmp_code"] == "ВУМ-001-01"
    assert row["file"] == "f.pdf"
    assert row["id"] is None
