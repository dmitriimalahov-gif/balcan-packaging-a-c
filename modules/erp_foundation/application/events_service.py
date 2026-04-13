# -*- coding: utf-8 -*-
from __future__ import annotations

from datetime import datetime, timezone
from typing import Any

from sqlalchemy import select

from db.models import DomainEvent
from db.session import session_scope


def list_recent_events(limit: int = 50) -> list[dict[str, Any]]:
    lim = max(1, min(int(limit), 500))
    with session_scope() as session:
        rows = session.scalars(
            select(DomainEvent)
            .order_by(DomainEvent.created_at.desc())
            .limit(lim)
        ).all()
        return [
            {
                "id": r.id,
                "event_type": r.event_type,
                "payload": r.payload,
                "created_at": r.created_at.isoformat() if r.created_at else "",
            }
            for r in rows
        ]


def record_domain_event(event_type: str, payload: dict[str, Any] | None = None) -> int:
    """Записать событие; вернуть id (для воркеров и application-слоёв)."""
    now = datetime.now(timezone.utc)
    with session_scope() as session:
        ev = DomainEvent(event_type=event_type, payload=payload, created_at=now)
        session.add(ev)
        session.flush()
        eid = int(ev.id)
    return eid
