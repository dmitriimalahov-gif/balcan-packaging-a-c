# -*- coding: utf-8 -*-
"""
Чтение / запись эталонных метаданных каталога макетов (JSON-файлы в корне проекта).
"""

from __future__ import annotations

import json
import os
import tempfile
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parent.parent.parent.parent

MAKETY_CATALOG_REF_PATH = ROOT / "makety_catalog_ref.json"
MAKETY_PATHS_REF_PATH = ROOT / "makety_paths_ref.json"

REF_CATALOG_TOTAL_ROWS = 852
REF_CATALOG_KIND_STATS: dict[str, int] = {
    "Коробки": 473,
    "Блистеры": 238,
    "Пакеты": 49,
    "Этикетки": 92,
}


def load_makety_paths_ref() -> tuple[Path | None, Path | None]:
    """Пути к эталонному Excel и БД из makety_paths_ref.json (если файл есть)."""
    p = MAKETY_PATHS_REF_PATH
    if not p.is_file():
        return None, None
    try:
        data = json.loads(p.read_text(encoding="utf-8"))
        ex = (data.get("excel_path") or "").strip()
        db = (data.get("db_path") or "").strip()

        def _resolve(rel: str) -> Path:
            q = Path(rel)
            return q.expanduser().resolve() if q.is_absolute() else (ROOT / rel).expanduser().resolve()

        return (
            _resolve(ex) if ex else None,
            _resolve(db) if db else None,
        )
    except Exception:
        return None, None


def load_makety_catalog_ref() -> tuple[int, dict[str, int]]:
    """Эталон для сверки: из JSON или встроенные константы."""
    p = MAKETY_CATALOG_REF_PATH
    if not p.is_file():
        return REF_CATALOG_TOTAL_ROWS, dict(REF_CATALOG_KIND_STATS)
    try:
        data = json.loads(p.read_text(encoding="utf-8"))
        total = int(data["total_rows"])
        raw_stats = data.get("kind_stats") or {}
        stats: dict[str, int] = {}
        for lbl in ("Коробки", "Блистеры", "Пакеты", "Этикетки"):
            stats[lbl] = int(raw_stats.get(lbl, 0))
        return total, stats
    except Exception:
        return REF_CATALOG_TOTAL_ROWS, dict(REF_CATALOG_KIND_STATS)


def save_makety_catalog_ref(total: int, stats: dict[str, int]) -> None:
    """Записать эталон (число строк и разбивка по видам) в makety_catalog_ref.json."""
    payload = {
        "total_rows": int(total),
        "kind_stats": {
            "Коробки": int(stats.get("Коробки", 0)),
            "Блистеры": int(stats.get("Блистеры", 0)),
            "Пакеты": int(stats.get("Пакеты", 0)),
            "Этикетки": int(stats.get("Этикетки", 0)),
        },
    }
    target = MAKETY_CATALOG_REF_PATH.expanduser().resolve()
    target.parent.mkdir(parents=True, exist_ok=True)
    fd, raw = tempfile.mkstemp(suffix=".json", dir=str(target.parent))
    os.close(fd)
    tmp = Path(raw)
    try:
        tmp.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        os.replace(tmp, target)
    except Exception:
        if tmp.exists():
            try:
                tmp.unlink()
            except OSError:
                pass
        raise
