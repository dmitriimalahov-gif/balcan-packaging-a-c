# -*- coding: utf-8 -*-
"""Чтение каталога из PostgreSQL (SQLAlchemy)."""

from __future__ import annotations

from typing import Any

from sqlalchemy import func, select

from db.models import PackagingItem
from db.session import session_scope


def list_packaging_items() -> list[dict[str, Any]]:
    with session_scope() as session:
        rows = session.scalars(
            select(PackagingItem).order_by(PackagingItem.excel_row)
        ).all()
        out: list[dict[str, Any]] = []
        for r in rows:
            out.append(
                {
                    "id": r.id,
                    "excel_row": r.excel_row,
                    "name": r.name or "",
                    "size": r.size or "",
                    "kind": r.kind or "",
                    "file": r.pdf_file or "",
                    "price": r.price or "",
                    "price_new": r.price_new or "",
                    "qty_per_sheet": r.qty_per_sheet or "",
                    "qty_per_year": r.qty_per_year or "",
                    "gmp_code": r.gmp_code or "",
                    "updated_at": r.updated_at.isoformat() if r.updated_at else "",
                }
            )
        return out


def count_packaging_items() -> int:
    with session_scope() as session:
        n = session.scalar(select(func.count()).select_from(PackagingItem))
        return int(n or 0)
