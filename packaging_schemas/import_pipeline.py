# -*- coding: utf-8 -*-
"""Пайплайн: сырые строки → валидация → отчёт об ошибках (без записи в БД)."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any

from pydantic import ValidationError

from packaging_schemas.makety_row import MaketyImportRow


@dataclass
class ImportRowError:
    index: int
    message: str
    field: str | None = None


def validate_makety_rows(
    raw_rows: list[dict[str, Any]],
) -> tuple[list[dict[str, Any]], list[ImportRowError]]:
    """
    Возвращает (валидные payload для upsert_all, ошибки по индексу входного списка).
    """
    ok: list[dict[str, Any]] = []
    errors: list[ImportRowError] = []
    for i, row in enumerate(raw_rows):
        try:
            m = MaketyImportRow.model_validate(row)
            ok.append(m.to_upsert_dict())
        except ValidationError as e:
            for err in e.errors():
                loc = err.get("loc", ())
                field = str(loc[0]) if loc else None
                errors.append(
                    ImportRowError(
                        index=i,
                        message=err.get("msg", "validation error"),
                        field=field,
                    )
                )
    return ok, errors
