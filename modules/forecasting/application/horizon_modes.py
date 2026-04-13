# -*- coding: utf-8 -*-
"""
Режимы горизонта прогноза (контракт из CURSOR_ERP_BOOTSTRAP §25).

Реализация расчётов — в отдельных сервисах, когда появятся заказы и склад в Postgres.
Инвариант: прогноз не перезаписывает фактические остатки и выпуск.
"""

from __future__ import annotations

from enum import Enum


class ForecastMode(str, Enum):
    CONSERVATIVE = "conservative"  # только подтверждённые заказы
    EXPECTED = "expected"  # заказы + взвешенный pipeline
    AGGRESSIVE = "aggressive"  # рост и агрессивные допущения


def describe_mode(mode: ForecastMode) -> str:
    return {
        ForecastMode.CONSERVATIVE: "Только подтверждённый спрос (фактические заказы).",
        ForecastMode.EXPECTED: "Подтверждённые заказы плюс вероятностный pipeline (квоты).",
        ForecastMode.AGGRESSIVE: "Расширенные допущения роста; требует явной маркировки в отчётах.",
    }[mode]
