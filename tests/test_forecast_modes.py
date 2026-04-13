# -*- coding: utf-8 -*-
from modules.forecasting.application.horizon_modes import ForecastMode, describe_mode


def test_describe_mode() -> None:
    assert "заказ" in describe_mode(ForecastMode.CONSERVATIVE).lower()
