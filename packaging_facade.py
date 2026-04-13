# -*- coding: utf-8 -*-
"""
Тонкий фасад над БД и доменом без Streamlit — для API, тестов и будущего SPA.

Реализация: `modules.packaging_catalog.infrastructure.catalog_sqlite`.
"""

from __future__ import annotations

from modules.packaging_catalog.infrastructure.catalog_sqlite import (
    count_packaging_rows,
    load_packaging_catalog,
)

__all__ = ["load_packaging_catalog", "count_packaging_rows"]
