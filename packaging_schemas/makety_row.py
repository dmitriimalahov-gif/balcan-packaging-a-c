# -*- coding: utf-8 -*-
"""Строка импорта каталога макетов (совместимо с полями upsert_all в packaging_db)."""

from __future__ import annotations

from typing import Any

from pydantic import BaseModel, ConfigDict, Field

import packaging_db as pkg_db


class MaketyImportRow(BaseModel):
    """Ожидаемые колонки после нормализации заголовков (как в Excel-импорте)."""

    model_config = ConfigDict(str_strip_whitespace=True)

    excel_row: int = Field(..., ge=1, description="Номер строки в мастер-Excel")
    name: str = Field(default="", max_length=2000)
    size: str = Field(default="", max_length=500)
    kind: str = Field(default="", max_length=500)
    file: str = Field(default="", max_length=2000, description="PDF: имя или путь")
    price: str = Field(default="", max_length=200)
    price_new: str = Field(default="", max_length=200)
    qty_per_sheet: str = Field(default="", max_length=200)
    qty_per_year: str = Field(default="", max_length=200)
    gmp_code: str = Field(default="", max_length=64)

    def to_upsert_dict(self) -> dict[str, Any]:
        d = self.model_dump()
        if not (d.get("gmp_code") or "").strip():
            d["gmp_code"] = pkg_db.extract_gmp_code(d.get("name") or "", d.get("file") or "")
        return d

    @classmethod
    def makety_json_schema(cls) -> dict[str, Any]:
        return cls.model_json_schema()
