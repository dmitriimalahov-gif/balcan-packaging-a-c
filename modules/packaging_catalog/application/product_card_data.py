# -*- coding: utf-8 -*-
"""
Сбор данных для карточки препарата: остатки упаковки, остатки субстанции,
прогноз продаж и рекомендация по заказу.
"""

from __future__ import annotations

import sqlite3
from dataclasses import dataclass, field
from datetime import date
from typing import Any

import packaging_db as pkg_db


@dataclass
class PackagingStockInfo:
    """Остатки одного вида упаковки."""
    kind: str
    kind_label: str
    qty: float = 0.0


@dataclass
class SubstanceStockInfo:
    """Остатки субстанции (препарата)."""
    qty: float = 0.0
    unit: str = "кг"


@dataclass
class SalesForecast:
    """Прогноз продаж на основе помесячных данных."""
    avg_monthly: float = 0.0
    last_12m_total: float = 0.0
    months_of_stock: float | None = None
    order_point_months: float = 2.0
    recommended_order_qty: float = 0.0
    recommended_order_date: str = ""


@dataclass
class ProductCardData:
    """Все данные для информационного блока карточки препарата."""
    gmp_code: str = ""
    packaging_stock: list[PackagingStockInfo] = field(default_factory=list)
    substance: SubstanceStockInfo = field(default_factory=SubstanceStockInfo)
    forecast: SalesForecast = field(default_factory=SalesForecast)


_KIND_LABELS = {
    "box": "Коробка",
    "blister": "Блистер",
    "label": "Этикетка",
    "pack": "Пакет",
}


def _calc_avg_monthly(monthly_rows: list[dict[str, Any]]) -> float:
    """Среднемесячные продажи за последние 12 месяцев (или за все доступные)."""
    if not monthly_rows:
        return 0.0
    today = date.today()
    recent: list[float] = []
    for r in monthly_rows:
        y, m = int(r["year"]), int(r["month"])
        months_ago = (today.year - y) * 12 + (today.month - m)
        if 0 <= months_ago < 12:
            recent.append(float(r["qty"]))
    if not recent:
        qtys = [float(r["qty"]) for r in monthly_rows]
        return sum(qtys) / max(len(qtys), 1)
    return sum(recent) / max(len(recent), 1)


def _calc_forecast(
    avg_monthly: float,
    box_stock: float,
    last_12m_total: float,
    *,
    lead_time_months: float = 2.0,
    order_batch_months: float = 3.0,
) -> SalesForecast:
    fc = SalesForecast(
        avg_monthly=round(avg_monthly, 1),
        last_12m_total=round(last_12m_total, 1),
        order_point_months=lead_time_months,
    )
    if avg_monthly > 0 and box_stock >= 0:
        fc.months_of_stock = round(box_stock / avg_monthly, 1)
        remaining_months = fc.months_of_stock - lead_time_months
        if remaining_months <= 0:
            fc.recommended_order_date = "Заказать сейчас"
        else:
            order_date = date.today()
            full_months = int(remaining_months)
            order_date = date(
                order_date.year + (order_date.month + full_months - 1) // 12,
                (order_date.month + full_months - 1) % 12 + 1,
                min(order_date.day, 28),
            )
            fc.recommended_order_date = order_date.strftime("%B %Y")
        fc.recommended_order_qty = round(avg_monthly * order_batch_months, 0)
    return fc


def collect_product_card_data(
    db_path: str | None,
    gmp_code: str,
    box_row: dict[str, Any],
    related_rows: dict[str, dict[str, Any] | None],
    all_rows: list[dict[str, Any]],
) -> ProductCardData:
    """
    Собирает все данные для информационного блока карточки.

    Параметры:
        db_path: путь к SQLite БД (может быть None)
        gmp_code: GMP-код препарата
        box_row: строка коробки
        related_rows: {"blister": row|None, "label": row|None, "pack": row|None}
        all_rows: все строки каталога (для поиска excel_row при расчёте прогноза)
    """
    data = ProductCardData(gmp_code=gmp_code.strip().upper())

    if not db_path or not data.gmp_code:
        for kind_key in ("box", "blister", "label", "pack"):
            data.packaging_stock.append(
                PackagingStockInfo(kind=kind_key, kind_label=_KIND_LABELS.get(kind_key, kind_key))
            )
        return data

    try:
        conn = pkg_db.connect(db_path)
        pkg_db.init_db(conn)
    except Exception:
        for kind_key in ("box", "blister", "label", "pack"):
            data.packaging_stock.append(
                PackagingStockInfo(kind=kind_key, kind_label=_KIND_LABELS.get(kind_key, kind_key))
            )
        return data

    try:
        pkg_stock = pkg_db.load_packaging_stock_for_gmp(conn, data.gmp_code)
        for kind_key in ("box", "blister", "label", "pack"):
            data.packaging_stock.append(
                PackagingStockInfo(
                    kind=kind_key,
                    kind_label=_KIND_LABELS.get(kind_key, kind_key),
                    qty=pkg_stock.get(kind_key, 0.0),
                )
            )

        sub_all = pkg_db.load_substance_stock(conn, data.gmp_code)
        sub = sub_all.get(data.gmp_code)
        if sub:
            data.substance = SubstanceStockInfo(qty=sub["qty"], unit=sub["unit"])

        excel_rows_for_gmp: list[int] = []
        for r in all_rows:
            r_gmp = (r.get("gmp_code") or "").strip().upper()
            if not r_gmp:
                from packaging_db import extract_gmp_code
                r_gmp = extract_gmp_code(r.get("name") or "", r.get("file") or "").strip().upper()
            if r_gmp == data.gmp_code:
                er = r.get("excel_row")
                if er is not None:
                    excel_rows_for_gmp.append(int(er))

        monthly_rows = pkg_db.load_monthly_for_rows(conn, excel_rows_for_gmp)
        avg_monthly = _calc_avg_monthly(monthly_rows)
        last_12m = sum(float(r["qty"]) for r in monthly_rows
                       if _is_recent_12m(r))
        box_stock = pkg_stock.get("box", 0.0)
        data.forecast = _calc_forecast(avg_monthly, box_stock, last_12m)

    except Exception:
        pass
    finally:
        try:
            conn.close()
        except Exception:
            pass

    return data


def _is_recent_12m(r: dict[str, Any]) -> bool:
    today = date.today()
    y, m = int(r["year"]), int(r["month"])
    months_ago = (today.year - y) * 12 + (today.month - m)
    return 0 <= months_ago < 12
