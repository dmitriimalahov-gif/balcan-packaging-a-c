# -*- coding: utf-8 -*-
from __future__ import annotations

from typing import Any

from fastapi import APIRouter

from modules.packaging_catalog.api.schemas import CatalogItemOut, CatalogResponse
from modules.packaging_catalog.application.catalog_read_service import (
    get_catalog_count,
    get_catalog_items,
)

router = APIRouter(prefix="/api/v1", tags=["catalog"])


@router.get("/items", response_model=CatalogResponse)
def list_items() -> dict[str, Any]:
    raw = get_catalog_items()
    total = get_catalog_count()
    items = [
        CatalogItemOut(
            id=row.get("id"),
            excel_row=int(row["excel_row"]),
            name=str(row.get("name") or ""),
            size=str(row.get("size") or ""),
            kind=str(row.get("kind") or ""),
            file=str(row.get("file") or ""),
            price=str(row.get("price") or ""),
            price_new=str(row.get("price_new") or ""),
            qty_per_sheet=str(row.get("qty_per_sheet") or ""),
            qty_per_year=str(row.get("qty_per_year") or ""),
            gmp_code=str(row.get("gmp_code") or ""),
            updated_at=str(row.get("updated_at") or ""),
        )
        for row in raw
    ]
    return {"items": items, "total": total}
