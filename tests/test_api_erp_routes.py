# -*- coding: utf-8 -*-
import pytest
from fastapi.testclient import TestClient


@pytest.fixture()
def client(monkeypatch: pytest.MonkeyPatch) -> TestClient:
    monkeypatch.delenv("PACKAGING_DATABASE_URL", raising=False)
    from api.main import app

    return TestClient(app)


def test_clients_empty_without_postgres(client: TestClient) -> None:
    r = client.get("/api/v1/clients")
    assert r.status_code == 200
    assert r.json() == {"items": []}


def test_events_empty_without_postgres(client: TestClient) -> None:
    r = client.get("/api/v1/events/recent")
    assert r.status_code == 200
    assert r.json() == {"items": []}
