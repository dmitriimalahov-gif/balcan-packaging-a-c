# -*- coding: utf-8 -*-
"""
Классификация строк каталога по виду упаковки и агрегация статистики.
"""

from __future__ import annotations

from typing import Any


def kind_bucket(item: dict[str, Any]) -> str:
    """Категории фильтра: box | blister | pack | label."""
    raw = (item.get("kind") or "").strip()
    k = raw.lower()
    if "без вторичной" in k or "fara secundar" in k or "fara cutie" in k:
        return "label"
    if raw == "Коробка" or "короб" in k:
        return "box"
    if "блистер" in k or "blister" in k:
        return "blister"
    if raw == "Пакет" or "пакет" in k:
        return "pack"
    return "label"


def kind_stats(rows: list[dict[str, Any]]) -> dict[str, int]:
    boxes = blisters = packs = labels = 0
    for r in rows:
        b = kind_bucket(r)
        if b == "box":
            boxes += 1
        elif b == "blister":
            blisters += 1
        elif b == "pack":
            packs += 1
        else:
            labels += 1
    return {
        "Коробки": boxes,
        "Блистеры": blisters,
        "Пакеты": packs,
        "Этикетки": labels,
    }


DEFAULT_KIND_OPTIONS = (
    "Коробка",
    "Blister",
    "Пакет",
    "Этикетка",
)


def build_kind_options(rows: list[dict[str, Any]]) -> list[str]:
    from_file = {r["kind"].strip() for r in rows if r.get("kind")}
    merged = set(DEFAULT_KIND_OPTIONS) | from_file
    return sorted(merged, key=lambda x: (x.lower(), x))
