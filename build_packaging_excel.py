#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Собрать из PDF макетов Balkan: название, размер (мм), вид упаковки."""

import re
import unicodedata
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font

from packaging_pdf_sizes import (
    canonicalize_extracted_size_text,
    extract_text_from_pdf,
)

ROOT = Path(__file__).resolve().parent


def norm(s: str) -> str:
    return unicodedata.normalize("NFKC", s).lower()


def classify_packaging(text: str, filename: str) -> str:
    t = norm(text)
    f = norm(filename)

    if "ambalajului secundar" in t or "ambalaj secundar" in t:
        return "Коробка"
    if "fara secundar" in t or "fara cutie" in t or "fara secundar" in f or "fara cutie" in f:
        return "Этикетка"

    if (
        "etichet" in t
        or "etichet" in f
        or "этикет" in t
        or "этикет" in f
        or "label" in f
    ):
        return "Этикетка"

    # Вторичная упаковка по коду в имени файла (ВУМ / типичный SPM с фасовкой)
    if re.search(r"\(вум-", f) or re.search(r"\(вум_", f):
        return "Коробка"
    if re.search(r"spm-", f) and re.search(r"\bn\d", f):
        return "Коробка"
    if re.search(r"\(bum-", f) and re.search(r"\bn\d", f):
        return "Коробка"
    if re.search(r"\(smp-", f) and re.search(r"\bn\d", f):
        return "Коробка"

    if "blister" in f or "(blister)" in f:
        return "Blister"

    blister_hints = (
        "blister",
        "pvc/al",
        "pvc / al",
        "pvc-al",
        "folie pvc",
        "folie aluminiu",
        "folie al",
    )
    if any(h in t for h in blister_hints):
        return "Blister"

    pouch_hints = (
        "doypack",
        "doy pack",
        "plic",
        "plicuri",
        "stick pack",
        "sachet",
        "punga",
        "порцион",
        "pulbere orala",
        "pulbere sol. orala",
    )
    if any(h in t for h in pouch_hints) or "plic" in f:
        return "Пакет"

    if "ambalajului primar" in t or "ambalaj primar" in t:
        if any(
            x in t
            for x in (
                "fiol",
                "fiole",
                "ampul",
                "ampou",
                "flacon",
                "flacoane",
                "sticla",
                "seringa",
                "sol. inj",
                "sol inj",
                "injectabila",
                "injectab",
            )
        ):
            return "Этикетка"
        if "capsul" in t or "comprimat" in t or "comp." in f or "comp " in f:
            return "Blister"
        if "pulbere" in t and "plic" not in t and "doy" not in t:
            return "Этикетка"
        return "Этикетка"

    # Фолбэк по имени: первичные коды документов
    if re.search(r"\(пум-", f) or re.search(r"\(ppm-", f) or re.search(r"\(lab-", f):
        if re.search(r"seringa|inj|fiol|flacon|sol\.", f):
            return "Этикетка"
        if "capsul" in f or "comp" in f or "compr" in f:
            return "Blister"
    if re.search(r"этк-", f) or re.search(r"etk-", f):
        if re.search(r"flacon|fiol|fiole|ampul|inj", f):
            return "Этикетка"
        return "Этикетка"

    if "borcan" in t or "borcan" in f:
        return "Этикетка"
    return "Этикетка"


def main():
    pdfs = sorted(ROOT.glob("*.pdf"))
    rows = []

    for p in pdfs:
        text = extract_text_from_pdf(p, max_pages=4)
        if text.startswith("__READ_ERROR__"):
            rows.append(
                {
                    "name": p.stem,
                    "size": "",
                    "kind": text,
                    "file": p.name,
                }
            )
            continue
        size = canonicalize_extracted_size_text(text)
        kind = classify_packaging(text, p.name)
        rows.append(
            {
                "name": p.stem,
                "size": size,
                "kind": kind,
                "file": p.name,
            }
        )

    wb = Workbook()
    ws = wb.active
    ws.title = "Упаковка"
    headers = (
        "Название (нож CG)",
        "Категория (CG)",
        "Лаки (CG)",
        "Размер (мм)",
        "Вид",
        "Исходный PDF",
        "Размер ножа (мм)",
        "Нож CG",
        "Цена (текущая)",
        "Новая цена",
        "Кол-во на листе",
        "Кол-во за год",
        "Название (cutii)",
    )
    ws.append(headers)
    for c in ws[1]:
        c.font = Font(bold=True)

    for r in rows:
        ws.append(
            [
                "",
                "",
                "",
                r["size"],
                r["kind"],
                r["file"],
                "",
                "",
                "",
                "",
                "",
                "",
                r["name"],
            ]
        )

    out = ROOT / "Упаковка_макеты.xlsx"
    wb.save(out)
    print(f"Записано строк: {len(rows)}")
    print(f"Файл: {out}")


if __name__ == "__main__":
    main()
