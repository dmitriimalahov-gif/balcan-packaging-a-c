#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Габариты в мм: перестановки одних и тех же чисел (84×15×62 ≡ 15×62×84) считаются одним размером.
Канонический вид — по убыванию: «84 × 62 × 15 mm».
"""

from __future__ import annotations

import re
from typing import Any


def normalize_size(s: str) -> str:
    """Единый вид размеров: пробелы вокруг ×, «mm» отделено пробелом."""
    t = (s or "").strip()
    if not t:
        return ""
    t = t.replace("х", "×").replace("Х", "×")
    t = re.sub(r"\s*[x×*]\s*", " × ", t, flags=re.IGNORECASE)
    t = re.sub(r"(?i)\s*mm\s*", " mm", t)
    t = re.sub(r"\s+", " ", t).strip()
    return t


def extract_gabarit_mm_values(size_str: str) -> list[float]:
    """Числа в диапазоне типичных габаритов макета (мм), по порядку в строке, до 6 шт."""
    if not size_str or not str(size_str).strip():
        return []
    s = str(size_str).replace(",", ".")
    raw = re.findall(r"\d+(?:\.\d+)?", s)
    dims: list[float] = []
    for x in raw:
        try:
            v = float(x)
        except ValueError:
            continue
        if not (4.0 <= v <= 950.0):
            continue
        dims.append(v)
        if len(dims) >= 6:
            break
    return dims


def parse_box_dimensions_mm(size_str: str) -> tuple[float, ...]:
    """Кортеж габаритов для сортировки: всегда по убыванию, дополнен нулями до 6."""
    dims = extract_gabarit_mm_values(size_str)
    if not dims:
        return (99999.0, 0.0, 0.0)
    sd = sorted(dims, reverse=True)
    while len(sd) < 6:
        sd.append(0.0)
    return tuple(sd[:6])


def canonicalize_size_mm(size_str: str) -> str:
    """Для 2–3 габаритов — канон по убыванию (перестановки совпадают). Для 4+ чисел — только normalize_size."""
    raw = (size_str or "").strip()
    if not raw:
        return ""
    dims = extract_gabarit_mm_values(raw)
    if not dims:
        return normalize_size(raw)
    if len(dims) >= 4:
        return normalize_size(raw)

    def fmt(v: float) -> str:
        if abs(v - round(v)) < 0.05:
            return str(int(round(v)))
        s = f"{v:.1f}".rstrip("0").rstrip(".")
        return s if s else str(int(round(v)))

    inner = " × ".join(fmt(v) for v in sorted(dims, reverse=True))
    return normalize_size(f"{inner} mm")


def size_key_from_string(size_str: str) -> str:
    """Ключ группировки по габаритам (перестановки дают тот же ключ)."""
    t = parse_box_dimensions_mm(size_str)
    if t[0] >= 99998:
        return "__empty__"
    parts: list[int] = []
    for x in t[:6]:
        if x > 0.01:
            parts.append(int(round(x)))
        elif parts:
            break
    return "|".join(str(p) for p in parts) if parts else "__empty__"


def row_size_key(item: dict[str, Any]) -> str:
    return size_key_from_string(item.get("size") or "")
