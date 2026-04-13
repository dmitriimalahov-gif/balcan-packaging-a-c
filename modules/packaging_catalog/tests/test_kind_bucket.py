# -*- coding: utf-8 -*-
from modules.packaging_catalog.domain.kind_bucket import (
    build_kind_options,
    kind_bucket,
    kind_stats,
)


def test_kind_bucket_box() -> None:
    assert kind_bucket({"kind": "Коробка"}) == "box"
    assert kind_bucket({"kind": "короб"}) == "box"


def test_kind_bucket_blister() -> None:
    assert kind_bucket({"kind": "Blister"}) == "blister"
    assert kind_bucket({"kind": "блистер"}) == "blister"


def test_kind_bucket_pack() -> None:
    assert kind_bucket({"kind": "Пакет"}) == "pack"


def test_kind_bucket_label_default() -> None:
    assert kind_bucket({"kind": "Этикетка"}) == "label"
    assert kind_bucket({"kind": ""}) == "label"


def test_kind_bucket_fara_cutie() -> None:
    assert kind_bucket({"kind": "Fara cutie"}) == "label"


def test_kind_stats() -> None:
    rows = [
        {"kind": "Коробка"},
        {"kind": "Blister"},
        {"kind": "Пакет"},
        {"kind": "Этикетка"},
        {"kind": "Коробка"},
    ]
    s = kind_stats(rows)
    assert s["Коробки"] == 2
    assert s["Блистеры"] == 1
    assert s["Пакеты"] == 1
    assert s["Этикетки"] == 1


def test_build_kind_options_includes_defaults() -> None:
    opts = build_kind_options([{"kind": "Custom"}])
    assert "Коробка" in opts
    assert "Custom" in opts
