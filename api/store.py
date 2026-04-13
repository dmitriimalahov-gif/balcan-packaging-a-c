# -*- coding: utf-8 -*-
"""Обратная совместимость: делегирование в application-слой каталога."""

from __future__ import annotations

from modules.packaging_catalog.application.catalog_read_service import (
    get_catalog_count,
    get_catalog_items,
)

__all__ = ["get_catalog_items", "get_catalog_count"]
