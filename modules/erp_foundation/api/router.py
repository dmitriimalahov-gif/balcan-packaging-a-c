# -*- coding: utf-8 -*-
from __future__ import annotations

import os

from fastapi import APIRouter

from modules.erp_foundation.api.schemas import (
    ClientOut,
    ClientsListResponse,
    DomainEventOut,
    DomainEventsResponse,
)
from modules.erp_foundation.application.clients_service import list_clients
from modules.erp_foundation.application.events_service import list_recent_events

router = APIRouter(prefix="/api/v1", tags=["erp"])


def _postgres_configured() -> bool:
    return (os.environ.get("PACKAGING_DATABASE_URL") or "").strip().lower().startswith(
        "postgresql"
    )


@router.get("/clients", response_model=ClientsListResponse)
def http_list_clients() -> dict:
    if not _postgres_configured():
        return {"items": []}
    raw = list_clients()
    return {
        "items": [
            ClientOut(
                id=r["id"],
                code=r["code"],
                name=r["name"],
                created_at=r["created_at"],
                updated_at=r["updated_at"],
            )
            for r in raw
        ]
    }


@router.get("/events/recent", response_model=DomainEventsResponse)
def http_recent_events(limit: int = 50) -> dict:
    if not _postgres_configured():
        return {"items": []}
    raw = list_recent_events(limit=limit)
    return {
        "items": [
            DomainEventOut(
                id=r["id"],
                event_type=r["event_type"],
                payload=r["payload"],
                created_at=r["created_at"],
            )
            for r in raw
        ]
    }
