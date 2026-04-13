# -*- coding: utf-8 -*-
from pydantic import BaseModel, Field


class CatalogItemOut(BaseModel):
    id: int | None = None
    excel_row: int
    name: str = ""
    size: str = ""
    kind: str = ""
    file: str = Field(default="", description="Имя/путь PDF (как в SQLite: pdf_file)")
    price: str = ""
    price_new: str = ""
    qty_per_sheet: str = ""
    qty_per_year: str = ""
    gmp_code: str = ""
    updated_at: str = ""


class CatalogResponse(BaseModel):
    items: list[CatalogItemOut]
    total: int
