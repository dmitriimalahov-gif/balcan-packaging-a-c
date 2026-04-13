# -*- coding: utf-8 -*-
from __future__ import annotations

from typing import Any

from sqlalchemy import select

from db.models import Client
from db.session import session_scope


def list_clients() -> list[dict[str, Any]]:
    with session_scope() as session:
        rows = session.scalars(select(Client).order_by(Client.code)).all()
        return [
            {
                "id": r.id,
                "code": r.code,
                "name": r.name,
                "created_at": r.created_at.isoformat() if r.created_at else "",
                "updated_at": r.updated_at.isoformat() if r.updated_at else "",
            }
            for r in rows
        ]
