#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""SQLite-хранилище строк упаковки (синхронно с Excel)."""

from __future__ import annotations

import sqlite3
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

DEFAULT_DB_PATH = Path(__file__).resolve().parent / "packaging_data.db"

_EXTRA_COLUMNS: tuple[tuple[str, str], ...] = (
    ("price", "TEXT"),
    ("price_new", "TEXT"),
    ("qty_per_sheet", "TEXT"),
    ("qty_per_year", "TEXT"),
    ("gmp_code", "TEXT"),
)


def connect(db_path: Path) -> sqlite3.Connection:
    conn = sqlite3.connect(str(db_path), check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn


def _ensure_extra_columns(conn: sqlite3.Connection) -> None:
    cur = conn.execute("PRAGMA table_info(packaging_items)")
    existing = {row[1] for row in cur.fetchall()}
    for col_name, col_type in _EXTRA_COLUMNS:
        if col_name not in existing:
            conn.execute(
                f"ALTER TABLE packaging_items ADD COLUMN {col_name} {col_type}"
            )
    conn.commit()


def _ensure_packaging_monthly_qty(conn: sqlite3.Connection) -> None:
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS packaging_monthly_qty (
            excel_row INTEGER NOT NULL,
            year INTEGER NOT NULL,
            month INTEGER NOT NULL,
            qty REAL NOT NULL,
            source TEXT,
            updated_at TEXT NOT NULL,
            PRIMARY KEY (excel_row, year, month)
        )
        """
    )
    conn.commit()


def _ensure_cutii_confirmations(conn: sqlite3.Connection) -> None:
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS cutii_confirmations (
            cutii_sheet_row INTEGER PRIMARY KEY,
            confirmed_excel_row INTEGER NOT NULL,
            cutii_name TEXT,
            updated_at TEXT NOT NULL
        )
        """
    )
    conn.commit()


def _ensure_print_tariffs(conn: sqlite3.Connection) -> None:
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS print_tariffs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            min_sheets INTEGER NOT NULL,
            max_sheets INTEGER,
            price_per_sheet REAL NOT NULL,
            updated_at TEXT NOT NULL
        )
        """
    )
    conn.commit()


def _ensure_stock_on_hand(conn: sqlite3.Connection) -> None:
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS stock_on_hand (
            gmp_code TEXT PRIMARY KEY,
            qty REAL NOT NULL,
            source TEXT,
            updated_at TEXT NOT NULL
        )
        """
    )
    conn.commit()


def _ensure_cg_knives(conn: sqlite3.Connection) -> None:
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS cg_knives (
            cutit_no TEXT PRIMARY KEY,
            name TEXT,
            category TEXT,
            cardboard TEXT,
            updated_at TEXT NOT NULL
        )
        """
    )
    conn.commit()


def _ensure_cg_prices(conn: sqlite3.Connection) -> None:
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS cg_prices (
            cutit_no TEXT NOT NULL,
            finish_type TEXT NOT NULL,
            min_qty INTEGER NOT NULL,
            max_qty INTEGER,
            price_per_1000 REAL NOT NULL,
            price_old_per_1000 REAL,
            updated_at TEXT NOT NULL,
            PRIMARY KEY (cutit_no, finish_type, min_qty)
        )
        """
    )
    conn.commit()
    try:
        conn.execute("SELECT price_old_per_1000 FROM cg_prices LIMIT 1")
    except Exception:
        try:
            conn.execute("ALTER TABLE cg_prices ADD COLUMN price_old_per_1000 REAL")
            conn.commit()
        except Exception:
            pass


def _ensure_cg_mapping(conn: sqlite3.Connection) -> None:
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS cg_mapping (
            excel_row INTEGER PRIMARY KEY,
            cutit_no TEXT NOT NULL,
            confirmed INTEGER DEFAULT 0,
            updated_at TEXT NOT NULL
        )
        """
    )
    conn.commit()


def _ensure_knife_cache(conn: sqlite3.Connection) -> None:
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS knife_cache (
            excel_row INTEGER PRIMARY KEY,
            svg_full TEXT NOT NULL,
            width_mm REAL NOT NULL,
            height_mm REAL NOT NULL,
            pdf_file TEXT,
            updated_at TEXT NOT NULL
        )
        """
    )
    conn.commit()


def init_db(conn: sqlite3.Connection) -> None:
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS packaging_items (
            excel_row INTEGER PRIMARY KEY,
            name TEXT,
            size TEXT,
            kind TEXT,
            pdf_file TEXT,
            price TEXT,
            price_new TEXT,
            qty_per_sheet TEXT,
            qty_per_year TEXT,
            updated_at TEXT NOT NULL
        )
        """
    )
    conn.commit()
    _ensure_extra_columns(conn)
    _ensure_packaging_monthly_qty(conn)
    _ensure_cutii_confirmations(conn)
    _ensure_print_tariffs(conn)
    _ensure_knife_cache(conn)
    _ensure_stock_on_hand(conn)
    _ensure_cg_knives(conn)
    _ensure_cg_prices(conn)
    _ensure_cg_mapping(conn)


def extract_gmp_code(name: str, filename: str = "") -> str:
    """Извлекает GMP-код вида ВУМ-169-01 из названия или имени файла."""
    import re
    for src in (name, filename):
        m = re.search(r'\(([A-ZА-Яa-zа-я]{2,4}-\d{2,4}-\d{2})\)', src)
        if m:
            return m.group(1).upper()
    return ""


def upsert_all(conn: sqlite3.Connection, rows: list[dict[str, Any]]) -> None:
    now = datetime.now(timezone.utc).isoformat()
    payload = [
        (
            int(r["excel_row"]),
            r.get("name") or "",
            r.get("size") or "",
            r.get("kind") or "",
            r.get("file") or "",
            r.get("price") or "",
            r.get("price_new") or "",
            r.get("qty_per_sheet") or "",
            r.get("qty_per_year") or "",
            r.get("gmp_code") or extract_gmp_code(r.get("name") or "", r.get("file") or ""),
            now,
        )
        for r in rows
    ]
    conn.executemany(
        """
        INSERT INTO packaging_items (
            excel_row, name, size, kind, pdf_file,
            price, price_new, qty_per_sheet, qty_per_year,
            gmp_code, updated_at
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(excel_row) DO UPDATE SET
            name = excluded.name,
            size = excluded.size,
            kind = excluded.kind,
            pdf_file = excluded.pdf_file,
            price = excluded.price,
            price_new = excluded.price_new,
            qty_per_sheet = excluded.qty_per_sheet,
            qty_per_year = excluded.qty_per_year,
            gmp_code = excluded.gmp_code,
            updated_at = excluded.updated_at
        """,
        payload,
    )
    conn.commit()


def load_all(conn: sqlite3.Connection) -> list[dict[str, Any]]:
    init_db(conn)
    cur = conn.execute(
        """
        SELECT excel_row, name, size, kind, pdf_file,
               price, price_new, qty_per_sheet, qty_per_year, gmp_code
        FROM packaging_items
        ORDER BY excel_row
        """
    )
    out: list[dict[str, Any]] = []
    for row in cur.fetchall():
        out.append(
            {
                "excel_row": int(row["excel_row"]),
                "name": row["name"] or "",
                "size": row["size"] or "",
                "kind": row["kind"] or "",
                "file": row["pdf_file"] or "",
                "price": row["price"] or "",
                "price_new": row["price_new"] or "",
                "qty_per_sheet": row["qty_per_sheet"] or "",
                "qty_per_year": row["qty_per_year"] or "",
                "gmp_code": row["gmp_code"] or "",
            }
        )
    return out


def row_count(conn: sqlite3.Connection) -> int:
    cur = conn.execute("SELECT COUNT(*) FROM packaging_items")
    return int(cur.fetchone()[0])


def upsert_monthly_batch(
    conn: sqlite3.Connection,
    rows: list[dict[str, Any]],
    *,
    default_source: str = "",
) -> None:
    """rows: excel_row, year, month, qty; optional source per row or default_source."""
    now = datetime.now(timezone.utc).isoformat()
    payload: list[tuple[Any, ...]] = []
    for r in rows:
        src = (r.get("source") or default_source or "").strip() or None
        payload.append(
            (
                int(r["excel_row"]),
                int(r["year"]),
                int(r["month"]),
                float(r["qty"]),
                src,
                now,
            )
        )
    if not payload:
        return
    _ensure_packaging_monthly_qty(conn)
    conn.executemany(
        """
        INSERT INTO packaging_monthly_qty (
            excel_row, year, month, qty, source, updated_at
        )
        VALUES (?, ?, ?, ?, ?, ?)
        ON CONFLICT(excel_row, year, month) DO UPDATE SET
            qty = excluded.qty,
            source = COALESCE(excluded.source, packaging_monthly_qty.source),
            updated_at = excluded.updated_at
        """,
        payload,
    )
    conn.commit()


def load_monthly_for_rows(
    conn: sqlite3.Connection,
    excel_rows: list[int],
) -> list[dict[str, Any]]:
    """Помесячные количества для указанных excel_row (пустой список = все)."""
    _ensure_packaging_monthly_qty(conn)
    if excel_rows:
        q_marks = ",".join("?" * len(excel_rows))
        cur = conn.execute(
            f"""
            SELECT excel_row, year, month, qty, source, updated_at
            FROM packaging_monthly_qty
            WHERE excel_row IN ({q_marks})
            ORDER BY excel_row, year, month
            """,
            [int(x) for x in excel_rows],
        )
    else:
        cur = conn.execute(
            """
            SELECT excel_row, year, month, qty, source, updated_at
            FROM packaging_monthly_qty
            ORDER BY excel_row, year, month
            """
        )
    return [
        {
            "excel_row": int(row["excel_row"]),
            "year": int(row["year"]),
            "month": int(row["month"]),
            "qty": float(row["qty"]),
            "source": row["source"] or "",
            "updated_at": row["updated_at"] or "",
        }
        for row in cur.fetchall()
    ]


def load_cutii_confirmations(conn: sqlite3.Connection) -> tuple[dict[int, int], dict[int, str]]:
    """Ручные сопоставления cutii (номер строки листа) → excel_row коробки; опционально зафиксированное имя из cutii."""
    _ensure_cutii_confirmations(conn)
    cur = conn.execute(
        """
        SELECT cutii_sheet_row, confirmed_excel_row, cutii_name
        FROM cutii_confirmations
        ORDER BY cutii_sheet_row
        """
    )
    mapping: dict[int, int] = {}
    names: dict[int, str] = {}
    for row in cur.fetchall():
        sr = int(row["cutii_sheet_row"])
        mapping[sr] = int(row["confirmed_excel_row"])
        cn = (row["cutii_name"] or "").strip()
        if cn:
            names[sr] = cn
    return mapping, names


def upsert_cutii_confirmations(
    conn: sqlite3.Connection,
    entries: list[dict[str, Any]],
) -> None:
    """Добавить или обновить строки; entries: cutii_sheet_row, confirmed_excel_row, опционально cutii_name."""
    if not entries:
        return
    _ensure_cutii_confirmations(conn)
    now = datetime.now(timezone.utc).isoformat()
    payload = [
        (
            int(e["cutii_sheet_row"]),
            int(e["confirmed_excel_row"]),
            (e.get("cutii_name") or "").strip() or None,
            now,
        )
        for e in entries
    ]
    conn.executemany(
        """
        INSERT INTO cutii_confirmations (
            cutii_sheet_row, confirmed_excel_row, cutii_name, updated_at
        )
        VALUES (?, ?, ?, ?)
        ON CONFLICT(cutii_sheet_row) DO UPDATE SET
            confirmed_excel_row = excluded.confirmed_excel_row,
            cutii_name = COALESCE(excluded.cutii_name, cutii_confirmations.cutii_name),
            updated_at = excluded.updated_at
        """,
        payload,
    )
    conn.commit()


# --- Тарифы печати ---


def load_tariffs(conn: sqlite3.Connection) -> list[dict[str, Any]]:
    _ensure_print_tariffs(conn)
    cur = conn.execute(
        "SELECT id, min_sheets, max_sheets, price_per_sheet FROM print_tariffs ORDER BY min_sheets"
    )
    return [
        {
            "id": int(row["id"]),
            "min_sheets": int(row["min_sheets"]),
            "max_sheets": int(row["max_sheets"]) if row["max_sheets"] is not None else None,
            "price_per_sheet": float(row["price_per_sheet"]),
        }
        for row in cur.fetchall()
    ]


def save_tariffs(conn: sqlite3.Connection, tariffs: list[dict[str, Any]]) -> None:
    """Полная перезапись тарифов (удаление + вставка)."""
    _ensure_print_tariffs(conn)
    now = datetime.now(timezone.utc).isoformat()
    conn.execute("DELETE FROM print_tariffs")
    for t in tariffs:
        conn.execute(
            """
            INSERT INTO print_tariffs (min_sheets, max_sheets, price_per_sheet, updated_at)
            VALUES (?, ?, ?, ?)
            """,
            (
                int(t["min_sheets"]),
                int(t["max_sheets"]) if t.get("max_sheets") is not None else None,
                float(t["price_per_sheet"]),
                now,
            ),
        )
    conn.commit()


# --- Кэш SVG-ножей ---


def save_knife(
    conn: sqlite3.Connection,
    excel_row: int,
    svg_full: str,
    width_mm: float,
    height_mm: float,
    pdf_file: str = "",
) -> None:
    _ensure_knife_cache(conn)
    now = datetime.now(timezone.utc).isoformat()
    conn.execute(
        """
        INSERT INTO knife_cache (excel_row, svg_full, width_mm, height_mm, pdf_file, updated_at)
        VALUES (?, ?, ?, ?, ?, ?)
        ON CONFLICT(excel_row) DO UPDATE SET
            svg_full = excluded.svg_full,
            width_mm = excluded.width_mm,
            height_mm = excluded.height_mm,
            pdf_file = excluded.pdf_file,
            updated_at = excluded.updated_at
        """,
        (int(excel_row), svg_full, float(width_mm), float(height_mm), pdf_file or "", now),
    )
    conn.commit()


def load_knife(conn: sqlite3.Connection, excel_row: int) -> dict[str, Any] | None:
    _ensure_knife_cache(conn)
    cur = conn.execute(
        "SELECT svg_full, width_mm, height_mm, pdf_file, updated_at FROM knife_cache WHERE excel_row = ?",
        (int(excel_row),),
    )
    row = cur.fetchone()
    if row is None:
        return None
    return {
        "excel_row": int(excel_row),
        "svg_full": row["svg_full"],
        "width_mm": float(row["width_mm"]),
        "height_mm": float(row["height_mm"]),
        "pdf_file": row["pdf_file"] or "",
        "updated_at": row["updated_at"] or "",
    }


def load_knives_for_rows(conn: sqlite3.Connection, excel_rows: list[int]) -> dict[int, dict[str, Any]]:
    _ensure_knife_cache(conn)
    if not excel_rows:
        cur = conn.execute("SELECT excel_row, svg_full, width_mm, height_mm, pdf_file, updated_at FROM knife_cache")
    else:
        q = ",".join("?" * len(excel_rows))
        cur = conn.execute(
            f"SELECT excel_row, svg_full, width_mm, height_mm, pdf_file, updated_at FROM knife_cache WHERE excel_row IN ({q})",
            [int(x) for x in excel_rows],
        )
    result: dict[int, dict[str, Any]] = {}
    for row in cur.fetchall():
        er = int(row["excel_row"])
        result[er] = {
            "excel_row": er,
            "svg_full": row["svg_full"],
            "width_mm": float(row["width_mm"]),
            "height_mm": float(row["height_mm"]),
            "pdf_file": row["pdf_file"] or "",
            "updated_at": row["updated_at"] or "",
        }
    return result


def load_knives_meta(conn: sqlite3.Connection) -> dict[int, dict[str, Any]]:
    """Загрузка только метаданных ножей (без тяжёлого svg_full)."""
    _ensure_knife_cache(conn)
    cur = conn.execute(
        "SELECT excel_row, width_mm, height_mm, pdf_file, updated_at FROM knife_cache"
    )
    result: dict[int, dict[str, Any]] = {}
    for row in cur.fetchall():
        er = int(row["excel_row"])
        result[er] = {
            "excel_row": er,
            "width_mm": float(row["width_mm"]),
            "height_mm": float(row["height_mm"]),
            "pdf_file": row["pdf_file"] or "",
            "updated_at": row["updated_at"] or "",
        }
    return result


def save_knives_batch(
    conn: sqlite3.Connection,
    knives: list[dict[str, Any]],
) -> None:
    """Пакетное сохранение ножей за одну транзакцию."""
    if not knives:
        return
    _ensure_knife_cache(conn)
    now = datetime.now(timezone.utc).isoformat()
    payload = [
        (
            int(k["excel_row"]),
            k["svg_full"],
            float(k["width_mm"]),
            float(k["height_mm"]),
            k.get("pdf_file") or "",
            now,
        )
        for k in knives
    ]
    conn.executemany(
        """
        INSERT INTO knife_cache (excel_row, svg_full, width_mm, height_mm, pdf_file, updated_at)
        VALUES (?, ?, ?, ?, ?, ?)
        ON CONFLICT(excel_row) DO UPDATE SET
            svg_full = excluded.svg_full,
            width_mm = excluded.width_mm,
            height_mm = excluded.height_mm,
            pdf_file = excluded.pdf_file,
            updated_at = excluded.updated_at
        """,
        payload,
    )
    conn.commit()


def propagate_knife_from_donor_to_group_rows(
    conn: sqlite3.Connection,
    donor_er: int,
    group_rows: list[dict[str, Any]],
) -> int:
    """
    Скопировать SVG-нож со строки ``donor_er`` на все остальные ``excel_row`` из ``group_rows``.

    Используется после ручного «Сохранить нож в БД»: один эталон → вся размерная группа в ``knife_cache``.
    Строка-донор не перезаписывается (у неё уже актуальная запись, в т.ч. имя PDF).
    """
    donor = load_knife(conn, int(donor_er))
    if donor is None:
        return 0
    d_er = int(donor_er)
    svg = donor["svg_full"]
    w_mm = float(donor["width_mm"])
    h_mm = float(donor["height_mm"])
    if w_mm <= 0 or h_mm <= 0:
        return 0
    batch: list[dict[str, Any]] = []
    for r in group_rows:
        er = int(r["excel_row"])
        if er == d_er:
            continue
        batch.append({
            "excel_row": er,
            "svg_full": svg,
            "width_mm": w_mm,
            "height_mm": h_mm,
            "pdf_file": f"propagated_from_er{d_er}",
        })
    if not batch:
        return 0
    save_knives_batch(conn, batch)
    return len(batch)


def propagate_knives_in_size_groups(
    conn: sqlite3.Connection,
    size_groups: list[dict[str, Any]],
    knife_meta: dict[int, dict[str, Any]],
    min_area_ratio: float = 0.18,
    size_key_filter: str | None = None,
) -> int:
    """Копировать эталонный SVG-нож на все продукты того же размера без ножа.

    Если ``size_key_filter`` задан — обрабатывается только группа с этим ``size_key``.

    Возвращает количество заполненных позиций.
    Обновляет knife_meta in-place.
    """
    from collections import Counter

    total_filled = 0
    batch_to_save: list[dict[str, Any]] = []

    for sg in size_groups:
        if size_key_filter is not None and sg.get("size_key") != size_key_filter:
            continue
        rows = sg["rows"]
        ers = [int(r["excel_row"]) for r in rows]
        with_knife = [er for er in ers if er in knife_meta and knife_meta[er]["width_mm"] > 0]
        without_knife = [er for er in ers if er not in knife_meta or knife_meta[er].get("width_mm", 0) <= 0]

        if not with_knife or not without_knife:
            continue

        dim_counter: Counter[tuple[float, float]] = Counter()
        er_by_dim: dict[tuple[float, float], list[int]] = {}
        for er in with_knife:
            m = knife_meta[er]
            w5 = round(m["width_mm"] / 5) * 5
            h5 = round(m["height_mm"] / 5) * 5
            key = (w5, h5)
            dim_counter[key] += 1
            er_by_dim.setdefault(key, []).append(er)

        best_dim, _ = dim_counter.most_common(1)[0]

        # Площадь ножа vs оценка развёртки коробки; низкий порог — больше групп получают распространение
        sample_size = sg.get("sample_size_str", "")
        expected_area = _estimate_box_area(sample_size)
        if expected_area and expected_area > 0:
            knife_area = best_dim[0] * best_dim[1]
            if knife_area < expected_area * min_area_ratio:
                continue

        candidates = er_by_dim[best_dim]
        best_er = max(
            candidates,
            key=lambda e: knife_meta[e]["width_mm"] * knife_meta[e]["height_mm"],
        )

        donor = load_knife(conn, best_er)
        if donor is None:
            continue

        for er in without_knife:
            entry = {
                "excel_row": er,
                "svg_full": donor["svg_full"],
                "width_mm": donor["width_mm"],
                "height_mm": donor["height_mm"],
                "pdf_file": f"propagated_from_er{best_er}",
            }
            batch_to_save.append(entry)
            knife_meta[er] = {
                "width_mm": donor["width_mm"],
                "height_mm": donor["height_mm"],
                "pdf_file": f"propagated_from_er{best_er}",
                "updated_at": "",
            }
            total_filled += 1

    if batch_to_save:
        save_knives_batch(conn, batch_to_save)

    return total_filled


def update_knife_dimensions(
    conn: sqlite3.Connection,
    excel_rows: list[int],
    width_mm: float,
    height_mm: float,
) -> int:
    """Обновить габариты ножа в кэше, SVG оставить прежним (коррекция метаданных)."""
    if not excel_rows or width_mm <= 0 or height_mm <= 0:
        return 0
    n = 0
    for er in excel_rows:
        row = load_knife(conn, int(er))
        if row is None:
            continue
        save_knife(
            conn,
            int(er),
            row["svg_full"],
            float(width_mm),
            float(height_mm),
            row.get("pdf_file") or "",
        )
        n += 1
    return n


def delete_knives_for_rows(conn: sqlite3.Connection, excel_rows: list[int]) -> int:
    """Удалить записи ножей из кэша для указанных excel_row."""
    if not excel_rows:
        return 0
    _ensure_knife_cache(conn)
    q = ",".join("?" * len(excel_rows))
    cur = conn.execute(
        f"DELETE FROM knife_cache WHERE excel_row IN ({q})",
        [int(x) for x in excel_rows],
    )
    conn.commit()
    return int(cur.rowcount or 0)


def _estimate_box_area(size_str: str) -> float:
    """Оценка площади развёртки коробки (2 стороны) по строке размера."""
    import re
    nums = re.findall(r'[\d]+(?:[.,]\d+)?', size_str)
    if len(nums) < 2:
        return 0.0
    vals = sorted([float(n.replace(",", ".")) for n in nums[:3]], reverse=True)
    if len(vals) >= 3:
        return 2 * (vals[0] * vals[1] + vals[0] * vals[2] + vals[1] * vals[2])
    return vals[0] * vals[1] * 2


# --- Складские остатки ---


def load_stock(conn: sqlite3.Connection) -> dict[str, float]:
    """Загрузить все складские остатки: {gmp_code: qty}."""
    _ensure_stock_on_hand(conn)
    cur = conn.execute("SELECT gmp_code, qty FROM stock_on_hand ORDER BY gmp_code")
    return {row["gmp_code"]: float(row["qty"]) for row in cur.fetchall()}


def load_stock_full(conn: sqlite3.Connection) -> list[dict[str, Any]]:
    """Загрузить складские остатки с метаданными."""
    _ensure_stock_on_hand(conn)
    cur = conn.execute(
        "SELECT gmp_code, qty, source, updated_at FROM stock_on_hand ORDER BY gmp_code"
    )
    return [
        {
            "gmp_code": row["gmp_code"],
            "qty": float(row["qty"]),
            "source": row["source"] or "",
            "updated_at": row["updated_at"] or "",
        }
        for row in cur.fetchall()
    ]


def upsert_stock_batch(
    conn: sqlite3.Connection,
    entries: list[dict[str, Any]],
    source: str = "",
) -> int:
    """Пакетная запись/обновление остатков. Возвращает количество записей."""
    if not entries:
        return 0
    _ensure_stock_on_hand(conn)
    now = datetime.now(timezone.utc).isoformat()
    payload = [
        (
            (e["gmp_code"] or "").strip().upper(),
            float(e["qty"]),
            (e.get("source") or source or "").strip() or None,
            now,
        )
        for e in entries
        if (e.get("gmp_code") or "").strip()
    ]
    if not payload:
        return 0
    conn.executemany(
        """
        INSERT INTO stock_on_hand (gmp_code, qty, source, updated_at)
        VALUES (?, ?, ?, ?)
        ON CONFLICT(gmp_code) DO UPDATE SET
            qty = excluded.qty,
            source = COALESCE(excluded.source, stock_on_hand.source),
            updated_at = excluded.updated_at
        """,
        payload,
    )
    conn.commit()
    return len(payload)


def clear_stock(conn: sqlite3.Connection) -> None:
    """Удалить все складские остатки."""
    _ensure_stock_on_hand(conn)
    conn.execute("DELETE FROM stock_on_hand")
    conn.commit()


# --- Каталог ножей типографии (CG) ---


def upsert_cg_knives(
    conn: sqlite3.Connection,
    knives: list[dict[str, Any]],
) -> int:
    if not knives:
        return 0
    _ensure_cg_knives(conn)
    now = datetime.now(timezone.utc).isoformat()
    payload = [
        (
            str(k["cutit_no"]).strip(),
            (k.get("name") or "").strip(),
            (k.get("category") or "").strip(),
            (k.get("cardboard") or "").strip(),
            now,
        )
        for k in knives
        if str(k.get("cutit_no") or "").strip()
    ]
    conn.executemany(
        """
        INSERT INTO cg_knives (cutit_no, name, category, cardboard, updated_at)
        VALUES (?, ?, ?, ?, ?)
        ON CONFLICT(cutit_no) DO UPDATE SET
            name = excluded.name,
            category = excluded.category,
            cardboard = excluded.cardboard,
            updated_at = excluded.updated_at
        """,
        payload,
    )
    conn.commit()
    return len(payload)


def load_cg_knives(conn: sqlite3.Connection) -> list[dict[str, Any]]:
    _ensure_cg_knives(conn)
    cur = conn.execute(
        "SELECT cutit_no, name, category, cardboard FROM cg_knives ORDER BY cutit_no"
    )
    return [
        {
            "cutit_no": row["cutit_no"],
            "name": row["name"] or "",
            "category": row["category"] or "",
            "cardboard": row["cardboard"] or "",
        }
        for row in cur.fetchall()
    ]


def upsert_cg_prices(
    conn: sqlite3.Connection,
    prices: list[dict[str, Any]],
) -> int:
    if not prices:
        return 0
    _ensure_cg_prices(conn)
    now = datetime.now(timezone.utc).isoformat()
    payload = [
        (
            str(p["cutit_no"]).strip(),
            str(p["finish_type"]).strip(),
            int(p["min_qty"]),
            int(p["max_qty"]) if p.get("max_qty") is not None else None,
            float(p["price_per_1000"]),
            float(p["price_old_per_1000"]) if p.get("price_old_per_1000") is not None else None,
            now,
        )
        for p in prices
        if p.get("price_per_1000") is not None
    ]
    conn.executemany(
        """
        INSERT INTO cg_prices (cutit_no, finish_type, min_qty, max_qty, price_per_1000, price_old_per_1000, updated_at)
        VALUES (?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(cutit_no, finish_type, min_qty) DO UPDATE SET
            max_qty = excluded.max_qty,
            price_per_1000 = excluded.price_per_1000,
            price_old_per_1000 = excluded.price_old_per_1000,
            updated_at = excluded.updated_at
        """,
        payload,
    )
    conn.commit()
    return len(payload)


def load_cg_prices(conn: sqlite3.Connection, cutit_no: str | None = None) -> list[dict[str, Any]]:
    _ensure_cg_prices(conn)
    if cutit_no:
        cur = conn.execute(
            "SELECT cutit_no, finish_type, min_qty, max_qty, price_per_1000, price_old_per_1000 "
            "FROM cg_prices WHERE cutit_no = ? ORDER BY finish_type, min_qty",
            (cutit_no,),
        )
    else:
        cur = conn.execute(
            "SELECT cutit_no, finish_type, min_qty, max_qty, price_per_1000, price_old_per_1000 "
            "FROM cg_prices ORDER BY cutit_no, finish_type, min_qty"
        )
    return [
        {
            "cutit_no": row["cutit_no"],
            "finish_type": row["finish_type"],
            "min_qty": int(row["min_qty"]),
            "max_qty": int(row["max_qty"]) if row["max_qty"] is not None else None,
            "price_per_1000": float(row["price_per_1000"]),
            "price_old_per_1000": float(row["price_old_per_1000"]) if row["price_old_per_1000"] is not None else None,
        }
        for row in cur.fetchall()
    ]


def cg_price_pair_at_tier(
    prices: list[dict[str, Any]],
    cutit_no: str,
    finish_type: str,
    min_qty: int,
) -> tuple[float | None, float | None]:
    """Старая и новая цена CG за 1000 шт. для конкретной ступени (finish_type + min_qty)."""
    for p in prices:
        if (
            p["cutit_no"] == cutit_no
            and p["finish_type"] == finish_type
            and int(p["min_qty"]) == int(min_qty)
        ):
            old_v = p.get("price_old_per_1000")
            return (
                float(old_v) if old_v is not None else None,
                float(p["price_per_1000"]),
            )
    return (None, None)


def cg_price_for_qty(
    prices: list[dict[str, Any]],
    finish_type: str,
    qty: int,
) -> float | None:
    """Цена за 1000 шт. для заданного типа лакирования и тиража."""
    matching = sorted(
        [p for p in prices if p["finish_type"] == finish_type],
        key=lambda p: p["min_qty"],
    )
    if not matching:
        return None
    best = None
    for p in matching:
        if qty >= p["min_qty"]:
            mx = p.get("max_qty")
            if mx is None or qty <= mx:
                best = p["price_per_1000"]
    if best is None and matching:
        best = matching[-1]["price_per_1000"]
    return best


def cg_old_price_for_qty(
    prices: list[dict[str, Any]],
    finish_type: str,
    qty: int,
) -> float | None:
    """Старая цена за 1000 шт. для заданного типа лакирования и тиража."""
    matching = sorted(
        [p for p in prices if p["finish_type"] == finish_type],
        key=lambda p: p["min_qty"],
    )
    if not matching:
        return None
    best = None
    for p in matching:
        old_p = p.get("price_old_per_1000")
        if old_p is None:
            old_p = p["price_per_1000"]
        if qty >= p["min_qty"]:
            mx = p.get("max_qty")
            if mx is None or qty <= mx:
                best = old_p
    if best is None and matching:
        best = matching[-1].get("price_old_per_1000") or matching[-1]["price_per_1000"]
    return best


def upsert_cg_mapping(
    conn: sqlite3.Connection,
    entries: list[dict[str, Any]],
) -> int:
    if not entries:
        return 0
    _ensure_cg_mapping(conn)
    now = datetime.now(timezone.utc).isoformat()
    payload = [
        (
            int(e["excel_row"]),
            str(e["cutit_no"]).strip(),
            int(e.get("confirmed", 0)),
            now,
        )
        for e in entries
        if e.get("cutit_no")
    ]
    conn.executemany(
        """
        INSERT INTO cg_mapping (excel_row, cutit_no, confirmed, updated_at)
        VALUES (?, ?, ?, ?)
        ON CONFLICT(excel_row) DO UPDATE SET
            cutit_no = excluded.cutit_no,
            confirmed = excluded.confirmed,
            updated_at = excluded.updated_at
        """,
        payload,
    )
    conn.commit()
    return len(payload)


def load_cg_mapping(conn: sqlite3.Connection) -> dict[int, dict[str, Any]]:
    _ensure_cg_mapping(conn)
    cur = conn.execute(
        "SELECT excel_row, cutit_no, confirmed FROM cg_mapping ORDER BY excel_row"
    )
    return {
        int(row["excel_row"]): {
            "cutit_no": row["cutit_no"],
            "confirmed": bool(row["confirmed"]),
        }
        for row in cur.fetchall()
    }


def clear_cg_data(conn: sqlite3.Connection) -> None:
    """Очистить все данные CG (ножи, цены, сопоставления)."""
    _ensure_cg_knives(conn)
    _ensure_cg_prices(conn)
    _ensure_cg_mapping(conn)
    conn.execute("DELETE FROM cg_prices")
    conn.execute("DELETE FROM cg_knives")
    conn.execute("DELETE FROM cg_mapping")
    conn.commit()


def sheet_price(n_sheets: int, tariffs: list[dict[str, Any]]) -> float:
    """Цена за 1 лист при тираже n_sheets. Если тарифов нет — 0."""
    if not tariffs or n_sheets <= 0:
        return 0.0
    for t in sorted(tariffs, key=lambda x: x["min_sheets"]):
        mx = t.get("max_sheets")
        if n_sheets >= t["min_sheets"] and (mx is None or n_sheets <= mx):
            return float(t["price_per_sheet"])
    return float(tariffs[-1]["price_per_sheet"])

