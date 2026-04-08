#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Извлечение габаритов (мм) из текста PDF: расширенные шаблоны для макетов Balkan / RO / RU."""

from __future__ import annotations

import re
from pathlib import Path

from packaging_sizes import (
    canonicalize_size_mm,
    extract_gabarit_mm_values,
    normalize_size,
)

try:
    import fitz  # PyMuPDF
except ImportError:
    fitz = None  # type: ignore


def _mm_plausible(v: float) -> bool:
    return 4.0 <= v <= 950.0


def _parse_dim_token(s: str) -> float | None:
    s = (s or "").strip().replace(",", ".")
    if not s:
        return None
    try:
        v = float(s)
    except ValueError:
        return None
    return v if _mm_plausible(v) else None


def extract_text_from_pdf(path: Path, max_pages: int = 4) -> str:
    """Текст первых страниц: pypdf, при коротком результате или сбое — PyMuPDF."""
    parts: list[str] = []
    pypdf_err: str | None = None
    try:
        from pypdf import PdfReader

        r = PdfReader(str(path))
        for page in r.pages[:max_pages]:
            parts.append(page.extract_text() or "")
    except Exception as e:
        pypdf_err = str(e)
    joined = "\n".join(parts)

    def _fitz_text() -> str:
        if fitz is None:
            return ""
        try:
            doc = fitz.open(str(path))
            alt: list[str] = []
            for i in range(min(max_pages, doc.page_count)):
                alt.append(doc.load_page(i).get_text("text") or "")
            doc.close()
            return "\n".join(alt)
        except Exception:
            return ""

    alt_j = _fitz_text()
    if alt_j and len(alt_j.strip()) > len(joined.strip()):
        joined = alt_j

    if not joined.strip():
        if pypdf_err:
            return f"__READ_ERROR__ {pypdf_err}"
        return "__READ_ERROR__ no text"
    return joined


def _dims_from_sequence(nums: list[float], max_n: int = 6) -> str:
    if not nums:
        return ""
    seen: set[float] = set()
    ordered: list[float] = []
    for v in nums:
        key = round(v, 2)
        if key not in seen:
            seen.add(key)
            ordered.append(v)
        if len(ordered) >= max_n:
            break
    if not ordered:
        return ""
    inner = "×".join(str(int(v)) if abs(v - round(v)) < 0.05 else str(v) for v in ordered)
    return inner + " mm"


def extract_sizes_mm_from_text(text: str) -> str:
    """
    Ищет размеры в типичных форматах макетов (мм).
    Возвращает строку вида «a×b mm» или «a×b×c mm» до канонизации.
    """
    if not text or text.startswith("__READ_ERROR__"):
        return ""

    raw = text.replace("\r", "\n")
    t = raw.replace("х", "x").replace("Х", "x")

    candidates: list[str] = []

    # --- Явные паттерны (по убыванию специфичности) ---

    # 90 mm x 65 mm / 90mm x 65mm
    for m in re.finditer(
        r"(?i)(\d{1,4}(?:[.,]\d+)?)\s*mm\s*[x×*]\s*(\d{1,4}(?:[.,]\d+)?)\s*mm",
        t,
    ):
        a, b = _parse_dim_token(m.group(1)), _parse_dim_token(m.group(2))
        if a is not None and b is not None:
            candidates.append(_dims_from_sequence([a, b]))

    # Вырубка / печать: 110+-2 mm 36+-2 mm 35.5+-2 mm (Balkan, CMYK)
    tol_seen: set[float] = set()
    tol_vals: list[float] = []
    for m in re.finditer(
        r"(?i)(\d{1,4}(?:[.,]\d+)?)(?:\s*\+\s*-\s*|\+\-)\d+\s*mm\b",
        t,
    ):
        v = _parse_dim_token(m.group(1))
        if v is None:
            continue
        k = round(v, 2)
        if k not in tol_seen:
            tol_seen.add(k)
            tol_vals.append(v)
    if len(tol_vals) >= 2:
        candidates.append(_dims_from_sequence(tol_vals))

    # Три числа: 84 x 15 x 62 mm / 84×15×62mm
    for m in re.finditer(
        r"(?i)(\d{1,4}(?:[.,]\d+)?)\s*[x×*]\s*(\d{1,4}(?:[.,]\d+)?)\s*[x×*]\s*(\d{1,4}(?:[.,]\d+)?)\s*mm",
        t,
    ):
        a, b, c = (
            _parse_dim_token(m.group(1)),
            _parse_dim_token(m.group(2)),
            _parse_dim_token(m.group(3)),
        )
        if a is not None and b is not None and c is not None:
            candidates.append(_dims_from_sequence([a, b, c]))

    # Два числа с одним mm в конце: 80 x 57 mm / 80×57 mm
    for m in re.finditer(
        r"(?i)(\d{1,4}(?:[.,]\d+)?)\s*[x×*]\s*(\d{1,4}(?:[.,]\d+)?)\s*mm\b",
        t,
    ):
        a, b = _parse_dim_token(m.group(1)), _parse_dim_token(m.group(2))
        if a is not None and b is not None:
            candidates.append(_dims_from_sequence([a, b]))

    # Четыре числа (редко)
    for m in re.finditer(
        r"(?i)(\d{1,4}(?:[.,]\d+)?)\s*[x×*]\s*(\d{1,4}(?:[.,]\d+)?)\s*[x×*]\s*(\d{1,4}(?:[.,]\d+)?)\s*[x×*]\s*(\d{1,4}(?:[.,]\d+)?)\s*mm",
        t,
    ):
        vals = [_parse_dim_token(m.group(i)) for i in range(1, 5)]
        if all(v is not None for v in vals):
            candidates.append(_dims_from_sequence([vals[0], vals[1], vals[2], vals[3]]))  # type: ignore

    # Строка после dimensiuni / gabarit / размер (одна строка чисел)
    for m in re.finditer(
        r"(?i)(dimensiuni|dimensiune|gabarit|cutie|dimensions?|размер|габарит)[^\n]{0,50}:?\s*([0-9\s.,×x*хХ]+)\s*mm",
        t,
    ):
        chunk = m.group(2)
        nums = []
        for x in re.findall(r"\d{1,4}(?:[.,]\d+)?", chunk):
            v = _parse_dim_token(x)
            if v is not None:
                nums.append(v)
        if len(nums) >= 2:
            candidates.append(_dims_from_sequence(nums[:6]))

    # Любые «число mm» по тексту (как раньше, но с десятичными и диапазоном 4–950)
    found = re.findall(r"(?i)\b(\d{1,4}(?:[.,]\d+)?)\s*mm\b", t)
    nums_mm: list[float] = []
    for x in found:
        v = _parse_dim_token(x)
        if v is not None:
            nums_mm.append(v)
    if nums_mm:
        seen: set[float] = set()
        ordered_mm: list[float] = []
        for v in nums_mm:
            k = round(v, 2)
            if k not in seen:
                seen.add(k)
                ordered_mm.append(v)
            if len(ordered_mm) >= 6:
                break
        if len(ordered_mm) >= 2:
            candidates.append(_dims_from_sequence(ordered_mm))

    # Числа с × без mm рядом (ограничим контекст: не дозировки типа 0,5)
    for m in re.finditer(
        r"(?i)(?<!\d)(\d{2,4}(?:[.,]\d+)?)\s*[x×]\s*(\d{2,4}(?:[.,]\d+)?)(?:\s*[x×]\s*(\d{2,4}(?:[.,]\d+)?))?(?=\s*(?:mm|мм|\n|$))",
        t,
    ):
        g = [m.group(1), m.group(2), m.group(3)]
        vals = [_parse_dim_token(x) for x in g if x]
        vals = [v for v in vals if v is not None]
        if len(vals) >= 2:
            candidates.append(_dims_from_sequence(vals))

    # Выбираем «лучший» кандидат: предпочтение большему числу измерений, затем длине строки
    best = ""
    best_score = (-1, -1)
    for c in candidates:
        if not c:
            continue
        score = (c.count("×") + 1, len(c))
        if score > best_score:
            best_score = score
            best = c

    if best:
        vals = extract_gabarit_mm_values(best.replace("×", " "))
        if len(vals) < 2:
            best = ""
    return best


def canonicalize_extracted_size_text(text: str) -> str:
    """Из уже извлечённого текста PDF — строка размера или пусто."""
    if not text or text.startswith("__READ_ERROR__"):
        return ""
    raw = extract_sizes_mm_from_text(text)
    if not raw:
        return ""
    return canonicalize_size_mm(normalize_size(raw))


def extract_and_canonicalize_size_from_pdf(path: Path, max_pages: int = 4) -> str:
    """Один проход: текст PDF → размеры → канон."""
    text = extract_text_from_pdf(path, max_pages=max_pages)
    return canonicalize_extracted_size_text(text)
