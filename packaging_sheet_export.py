# -*- coding: utf-8 -*-
"""Экспорт схемы печатного листа: растры слотов и сборка PDF (PyMuPDF)."""

from __future__ import annotations

import base64
import textwrap
from datetime import datetime
from pathlib import Path
from typing import Any

import fitz

import packaging_pdf_sheet_preview as ppsp
import pdf_outline_to_svg as pdf_outline
from packaging_print_planning import PlacedRect, SheetParams

MM_TO_PT = 72.0 / 25.4


def layout_raster_kind_bucket(row: dict[str, Any] | None) -> str:
    """
    Как ``kind_bucket`` в packaging_viewer: box | blister | pack | label.
    Для растра слота по контуру PDF (стратегия crop).
    """
    if not row:
        return "box"
    raw = (row.get("kind") or "").strip()
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


# Balkan / макеты: фирменная магента с обводок Corel (#E61081)
_BRAND_RGB: tuple[float, float, float] = (230 / 255.0, 16 / 255.0, 129 / 255.0)
_BRAND_RGB_DARK: tuple[float, float, float] = (180 / 255.0, 10 / 255.0, 95 / 255.0)
_BRAND_PANEL_BG: tuple[float, float, float] = (0.98, 0.96, 0.98)
_TEXT_PRIMARY: tuple[float, float, float] = (0.16, 0.17, 0.2)
_TEXT_MUTED: tuple[float, float, float] = (0.42, 0.44, 0.48)
_RULE_LIGHT: tuple[float, float, float] = (0.88, 0.86, 0.9)
_HEADER_H_PT = 52.0
_CONTENT_LEFT = 50.0
_CONTENT_RIGHT = 545.0
_CONTENT_BODY_Y = 62.0
_CONTENT_BOTTOM = 778.0


def _resolve_unicode_font() -> Path | None:
    """Системный шрифт с кириллицей для insert_text (если есть)."""
    candidates = [
        Path("/System/Library/Fonts/Supplemental/Arial Unicode.ttf"),
        Path("/Library/Fonts/Arial Unicode.ttf"),
        Path("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"),
        Path("/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf"),
        Path("C:/Windows/Fonts/arial.ttf"),
        Path("C:/Windows/Fonts/calibri.ttf"),
    ]
    for p in candidates:
        if p.is_file():
            return p
    return None


def build_slot_png_and_outline(
    *,
    pl_active: list[PlacedRect],
    slot_er_list: list[int | None],
    rows_by_er: dict[int, dict[str, Any]],
    pdf_root: Path,
    dpi: float,
    knife_raster: bool,
    transparent_png: bool,
    show_outline: bool,
    knives_by_er: dict[int, dict[str, Any]] | None = None,
) -> tuple[list[str | None], list[str | None], list[bytes | None]]:
    """
    Рендер PNG и контур для каждого слота (без ограничения «первые N»).
    Возвращает (base64_png, inner_outline_svg, raw_png_bytes).

    Если передан ``knives_by_er`` (как из ``knife_cache``): для каждого слота **сначала** берётся
    растр и контур из **PDF этой строки** (если файл есть), иначе — из сохранённого SVG в кэше.
    """
    dpi_eff = max(36.0, min(float(dpi), 240.0))
    b64_list: list[str | None] = []
    outlines: list[str | None] = []
    raw_list: list[bytes | None] = []
    kmap = knives_by_er or {}

    for idx, p_rect in enumerate(pl_active):
        er = slot_er_list[idx] if idx < len(slot_er_list) else None
        pb: bytes | None = None
        ol: str | None = None
        if er is not None:
            row = rows_by_er.get(int(er))
            rel = (row.get("file") or "").strip() if row else ""
            p_pdf = ppsp.resolve_pdf_path(pdf_root, rel) if rel else None
            kn = kmap.get(int(er))
            kn_svg = (kn.get("svg_full") or "").strip() if kn else ""
            kn_w = float(kn.get("width_mm") or 0) if kn else 0.0
            kn_h = float(kn.get("height_mm") or 0) if kn else 0.0
            has_db_knife = bool(kn_svg and kn_w > 0 and kn_h > 0)
            if knife_raster:
                if p_pdf and p_pdf.is_file():
                    _lk = layout_raster_kind_bucket(row)
                    pb = ppsp.render_knife_bbox_fit_to_mm(
                        str(p_pdf),
                        float(p_rect.w),
                        float(p_rect.h),
                        dpi=dpi_eff,
                        transparent_bg=transparent_png,
                        layout_kind=_lk,
                    )
                if pb is None and has_db_knife:
                    pb = ppsp.render_cached_svg_knife_fit_to_mm(
                        kn_svg,
                        float(p_rect.w),
                        float(p_rect.h),
                        dpi=dpi_eff,
                        transparent_bg=transparent_png,
                    )
                if pb is None and p_pdf and p_pdf.is_file():
                    pb = ppsp.render_first_page_fit_to_mm(
                        str(p_pdf),
                        float(p_rect.w),
                        float(p_rect.h),
                        dpi=dpi_eff,
                        transparent_bg=transparent_png,
                    )
            elif p_pdf and p_pdf.is_file():
                pb = ppsp.render_first_page_fit_to_mm(
                    str(p_pdf),
                    float(p_rect.w),
                    float(p_rect.h),
                    dpi=dpi_eff,
                    transparent_bg=transparent_png,
                )
            if show_outline:
                ol = None
                if p_pdf and p_pdf.is_file():
                    ol = (
                        pdf_outline.extract_outline_svg_inner_for_slot(
                            str(p_pdf),
                            float(p_rect.w),
                            float(p_rect.h),
                            0,
                        )
                        or None
                    )
                if (ol is None or ol == "") and has_db_knife:
                    ol = pdf_outline.inner_outline_from_stored_knife_svg(
                        kn_svg,
                        float(p_rect.w),
                        float(p_rect.h),
                        kn_w,
                        kn_h,
                    ) or None
        raw_list.append(pb)
        b64_list.append(base64.b64encode(pb).decode("ascii") if pb else None)
        outlines.append(ol)

    return b64_list, outlines, raw_list


def _svg_bytes_to_png(svg_bytes: bytes, max_side_pt: float = 200.0) -> bytes | None:
    """Растеризация SVG (схема раскладки) в PNG для вставки в PDF."""
    if not svg_bytes or not svg_bytes.strip():
        return None
    try:
        src = fitz.open(stream=svg_bytes, filetype="svg")
        if src.page_count < 1:
            src.close()
            return None
        sp = src[0]
        r = sp.rect
        if r.width <= 0 or r.height <= 0:
            src.close()
            return None
        scale = min(max_side_pt / r.width, max_side_pt / r.height, 4.0)
        mat = fitz.Matrix(scale, scale)
        pix = sp.get_pixmap(matrix=mat, alpha=False)
        out = pix.tobytes("png")
        src.close()
        return out
    except Exception:
        return None


def _truncate_pdf_cell(s: str, max_chars: int) -> str:
    t = " ".join(str(s or "").replace("\n", " ").split())
    if len(t) <= max_chars:
        return t
    return t[: max_chars - 1] + "…"


def _insert_font_on_page(page: fitz.Page, font_path: Path | None) -> str:
    if font_path and font_path.is_file():
        try:
            page.insert_font("exf", fontfile=str(font_path))
            return "exf"
        except Exception:
            pass
    return "helv"


def _summary_draw_brand_header(pg: fitz.Page, a4w: float, fn: str) -> None:
    """Шапка в стиле фирменных макетов (магента #E61081)."""
    pg.draw_rect(fitz.Rect(0, 0, a4w, _HEADER_H_PT), color=_BRAND_RGB, fill=_BRAND_RGB, width=0)
    pg.draw_rect(
        fitz.Rect(0, _HEADER_H_PT, a4w, _HEADER_H_PT + 2.5),
        color=_BRAND_RGB_DARK,
        fill=_BRAND_RGB_DARK,
        width=0,
    )
    pg.insert_text(
        (_CONTENT_LEFT, 24),
        "BALKAN PHARMACEUTICALS",
        fontname=fn,
        fontsize=13.5,
        color=(1, 1, 1),
    )
    pg.insert_text(
        (_CONTENT_LEFT, 40),
        "Сводка по печатному листу · внутренний документ",
        fontname=fn,
        fontsize=8.2,
        color=(1.0, 0.92, 0.97),
    )


def _summary_draw_footer(pg: fitz.Page, a4w: float, a4h: float, fn: str, page_no: int) -> None:
    pg.draw_line(
        fitz.Point(_CONTENT_LEFT, a4h - 36),
        fitz.Point(_CONTENT_RIGHT, a4h - 36),
        color=_RULE_LIGHT,
        width=0.75,
    )
    ts = datetime.now().strftime("%Y-%m-%d %H:%M")
    pg.insert_text(
        (_CONTENT_LEFT, a4h - 22),
        f"Сформировано: {ts}  ·  Balkan Pharmaceuticals  ·  стр. {page_no}",
        fontname=fn,
        fontsize=7.2,
        color=_TEXT_MUTED,
    )


def _summary_wrap_lines(text: str, width_chars: int) -> list[str]:
    t = " ".join(str(text).replace("\r", "").split())
    if not t:
        return []
    return textwrap.wrap(
        t,
        width=max(12, width_chars),
        break_long_words=True,
        break_on_hyphens=True,
    )


def _append_rich_summary_pages(
    doc: fitz.Document,
    *,
    font_path: Path | None,
    title_line: str,
    summary: dict[str, Any],
) -> None:
    """Сводка A4: фирменная шапка Balkan (#E61081), таблицы, схема, без обрезки — только переносы."""
    a4w, a4h = 595.0, 842.0
    page: fitz.Page | None = None
    fn = "helv"
    y = _CONTENT_BODY_Y
    page_no = 0

    def new_page() -> None:
        nonlocal page, y, fn, page_no
        if page is not None:
            _summary_draw_footer(page, a4w, a4h, fn, page_no)
        page = doc.new_page(width=a4w, height=a4h)
        page_no += 1
        fn = _insert_font_on_page(page, font_path)
        _summary_draw_brand_header(page, a4w, fn)
        y = _CONTENT_BODY_Y

    def ensure_space(need_pt: float) -> None:
        nonlocal y
        if page is None or y + need_pt > _CONTENT_BOTTOM:
            new_page()

    def emit_lines(lines: list[str], fs: float) -> None:
        nonlocal y
        lh = fs * 1.28
        for ln in lines:
            ensure_space(lh + 2)
            assert page is not None
            page.insert_text(
                (_CONTENT_LEFT, y),
                ln,
                fontname=fn,
                fontsize=fs,
                color=_TEXT_PRIMARY,
            )
            y += lh
        y += 4

    def emit_paragraph(text: str, fs: float = 9.0, ch: int = 72) -> None:
        emit_lines(_summary_wrap_lines(text, ch), fs)

    def section_title(label: str) -> None:
        nonlocal y
        ensure_space(26)
        assert page is not None
        page.draw_rect(fitz.Rect(_CONTENT_LEFT, y - 1, _CONTENT_LEFT + 3.5, y + 12), fill=_BRAND_RGB, width=0)
        page.insert_text(
            (_CONTENT_LEFT + 10, y + 9),
            label,
            fontname=fn,
            fontsize=10.8,
            color=_BRAND_RGB_DARK,
        )
        y += 20

    def panel_start() -> float:
        """Возвращает y начала панели (для рамки)."""
        nonlocal y
        ensure_space(16)
        return y

    def panel_end(y0: float) -> None:
        nonlocal y
        assert page is not None
        if y > y0 + 2:
            page.draw_rect(
                fitz.Rect(_CONTENT_LEFT - 4, y0 - 6, _CONTENT_RIGHT + 4, y + 2),
                color=_RULE_LIGHT,
                width=0.55,
            )

    new_page()

    emit_paragraph(title_line or "—", fs=10.0, ch=68)
    y += 2

    sm = summary.get("sheet_meta") or {}
    if sm:
        section_title("Параметры листа")
        y0 = panel_start()
        emit_paragraph(
            f"Формат: {sm.get('w', '—')} × {sm.get('h', '—')} мм. Поля: {sm.get('m', '—')} мм. "
            f"Зазор между оттисками: X {sm.get('gx', '—')} мм, Y {sm.get('gy', '—')} мм. "
            f"Слотов на листе: {sm.get('n_slots', '—')}. Листов в партии: {sm.get('n_sheets', '—')}.",
            fs=9.0,
            ch=70,
        )
        panel_end(y0)

    kn = summary.get("knife_note") or ""
    if kn:
        section_title("Габариты ножа / ячейки")
        y0 = panel_start()
        emit_paragraph(str(kn), fs=9.0, ch=72)
        panel_end(y0)

    svg_b = summary.get("layout_svg_bytes")
    if isinstance(svg_b, (bytes, bytearray)) and svg_b:
        png_thumb = _svg_bytes_to_png(bytes(svg_b), max_side_pt=200.0)
        if png_thumb:
            section_title("Схема раскладки (контуры слотов)")
            tw, th = 200.0, 142.0
            ensure_space(th + 28)
            assert page is not None
            r_img = fitz.Rect(_CONTENT_RIGHT - tw, y, _CONTENT_RIGHT, y + th)
            try:
                page.insert_image(r_img, stream=png_thumb, keep_proportion=True)
            except Exception:
                pass
            emit_paragraph(
                "Уменьшенная копия векторной схемы: расположение слотов и контуры ножей на печатном листе. "
                "Полноразмерный макет — на первой странице PDF.",
                fs=8.6,
                ch=58,
            )
            y = max(y, r_img.y1 + 10)

    exn = summary.get("extras_note") or ""
    if exn:
        section_title("Отделка и доплаты к печати")
        y0 = panel_start()
        for part in str(exn).split("; "):
            part = part.strip()
            if part:
                emit_paragraph(part, fs=9.0, ch=72)
        panel_end(y0)

    econ = summary.get("economics")
    if isinstance(econ, dict) and econ.get("has_data"):
        section_title("Экономика (планировщик)")
        y0 = panel_start()
        for line in econ.get("lines") or []:
            emit_paragraph(str(line), fs=9.0, ch=72)
        panel_end(y0)

    party = summary.get("party_rows") or []
    if party:
        section_title("Партия по видам")
        fs_t = 8.4
        lh = fs_t * 1.35
        x_name = _CONTENT_LEFT
        x_nums = _CONTENT_LEFT + 248
        ensure_space(lh + 8)
        assert page is not None
        page.draw_rect(fitz.Rect(x_name - 2, y - 4, _CONTENT_RIGHT + 2, y + lh - 2), fill=_BRAND_PANEL_BG, width=0)
        page.insert_text((x_name, y + fs_t * 0.85), "Наименование", fontname=fn, fontsize=fs_t, color=_BRAND_RGB_DARK)
        page.insert_text((x_nums, y + fs_t * 0.85), "Яч.  Лист.  Оттиск.   БД   €/1000", fontname=fn, fontsize=fs_t, color=_BRAND_RGB_DARK)
        y += lh + 4
        for row in party:
            name = str(row.get("name", row.get("Название", "")) or "")
            c1 = str(row.get("cells", row.get("Ячеек на 1 листе", "")))
            c2 = str(row.get("sheets", row.get("Листов в партии", "")))
            c3 = str(row.get("imprints", row.get("Всего оттисков (ячейки×листы)", "")))
            c4 = str(row.get("db_qps", row.get("Кол-во на листе (БД)", "")))
            c5 = str(row.get("eur_per_1000", "—"))
            name_lines = _summary_wrap_lines(name, 36) or ["—"]
            row_h = max(len(name_lines) * lh, lh) + 6
            ensure_space(row_h + 4)
            assert page is not None
            yy = y
            for i, nl in enumerate(name_lines):
                page.insert_text((x_name, yy + fs_t * 0.85 + i * lh), nl, fontname=fn, fontsize=fs_t, color=_TEXT_PRIMARY)
            num_line = f"{c1:>4}  {c2:>5}  {c3:>7}  {str(c4)[:8]:>8}  {c5:>8}"
            page.insert_text((x_nums, y + fs_t * 0.85), num_line, fontname=fn, fontsize=fs_t, color=_TEXT_PRIMARY)
            y += row_h
        y += 4

    slots = summary.get("slot_rows") or []
    if slots:
        section_title("Назначение слотов")
        for sr in slots:
            slot_txt = (
                f"Слот {sr.get('slot', '?')}: Excel-строка {sr.get('er', '—')}. "
                f"{str(sr.get('label', '') or '—')}"
            )
            emit_paragraph(slot_txt, fs=8.8, ch=72)

    cg_rows = summary.get("cg_rows") or []
    if cg_rows:
        section_title("Оценка CG (€ за 1000 шт.)")
        fs_t = 8.2
        lh = fs_t * 1.32
        for cr in cg_rows:
            head = (
                f"ER {cr.get('er', '—')}  ·  нож {cr.get('cutit', '—')}  ·  {cr.get('finish', '—')}  ·  "
                f"тираж {cr.get('qty', '—')} шт.  ·  €/1000: {cr.get('eur_per_1000', '—')}"
            )
            emit_paragraph(head, fs=fs_t, ch=74)
            nm = str(cr.get("name", "") or "").strip()
            if nm:
                emit_paragraph(f"Наименование: {nm}", fs=fs_t - 0.2, ch=76)
            y += 2

    tech = summary.get("tech_lines") or []
    if tech:
        section_title("Параметры превью / экспорта")
        y0 = panel_start()
        for t in tech:
            emit_paragraph(str(t), fs=8.4, ch=74)
        panel_end(y0)

    legacy = summary.get("legacy_stats_lines") or []
    if legacy:
        section_title("Полный текстовый дамп (детали)")
        for line in legacy:
            emit_paragraph(str(line), fs=7.8, ch=78)

    if page is not None:
        _summary_draw_footer(page, a4w, a4h, fn, page_no)


def sheet_layout_to_pdf_bytes(
    sheet_params: SheetParams,
    pl_active: list[PlacedRect],
    slot_png_bytes: list[bytes | None],
    stats_lines: list[str] | None = None,
    *,
    title_line: str = "",
    summary: dict[str, Any] | None = None,
) -> bytes:
    """
    Страница 1: визуальный лист с PNG по слотам; опционально мини-схема SVG в углу.
    Далее A4: сводка — таблицы, схема, CG, экономика (если передано в ``summary``).
    """
    doc = fitz.open()
    wp = float(sheet_params.width_mm) * MM_TO_PT
    hp = float(sheet_params.height_mm) * MM_TO_PT
    page = doc.new_page(width=wp, height=hp)
    page.draw_rect(page.rect, color=(1, 1, 1), fill=(1, 1, 1))

    mpt = float(sheet_params.margin_mm) * MM_TO_PT
    iw = float(sheet_params.width_mm) * MM_TO_PT - 2 * mpt
    ih = float(sheet_params.height_mm) * MM_TO_PT - 2 * mpt
    inner_r = fitz.Rect(mpt, mpt, mpt + iw, mpt + ih)
    page.draw_rect(inner_r, color=(0.55, 0.55, 0.55), width=0.6, dashes="[3 3] 0")

    for i, pr in enumerate(pl_active):
        x0 = mpt + float(pr.x) * MM_TO_PT
        y0 = mpt + float(pr.y) * MM_TO_PT
        x1 = x0 + float(pr.w) * MM_TO_PT
        y1 = y0 + float(pr.h) * MM_TO_PT
        rect = fitz.Rect(x0, y0, x1, y1)
        stream = slot_png_bytes[i] if i < len(slot_png_bytes) else None
        if stream:
            try:
                page.insert_image(rect, stream=stream, keep_proportion=True)
            except Exception:
                pass
        page.draw_rect(rect, color=(0.09, 0.29, 0.62), width=0.9)

    layout_corner = summary.get("layout_svg_bytes") if summary else None
    if isinstance(layout_corner, (bytes, bytearray)) and layout_corner:
        png_c = _svg_bytes_to_png(bytes(layout_corner), max_side_pt=min(140.0, wp * 0.22))
        if png_c:
            cw = min(wp * 0.22, 140.0)
            ch = min(hp * 0.22, 100.0)
            inset = fitz.Rect(wp - cw - 10, hp - ch - 10, wp - 10, hp - 10)
            try:
                page.insert_image(inset, stream=png_c, keep_proportion=True)
            except Exception:
                pass
            try:
                page.insert_text(
                    (wp - cw - 8, hp - ch - 14),
                    "Раскладка (схема)",
                    fontsize=6,
                    color=(0.2, 0.2, 0.2),
                )
            except Exception:
                pass

    font_path = _resolve_unicode_font()
    if summary:
        summ = dict(summary)
        if stats_lines:
            summ.setdefault("legacy_stats_lines", list(stats_lines))
        _append_rich_summary_pages(doc, font_path=font_path, title_line=title_line, summary=summ)
    else:
        a4w, a4h = 595.0, 842.0
        p2 = doc.new_page(width=a4w, height=a4h)
        fontname = _insert_font_on_page(p2, font_path)
        left, top0 = 48.0, 52.0
        y = top0
        lh_title = 14.0
        lh_body = 10.5
        fs_title = 11.0
        fs_body = 9.0
        if title_line:
            p2.insert_text((left, y), title_line, fontname=fontname, fontsize=fs_title)
            y += lh_title
        y += 4.0
        for line in stats_lines or []:
            if y > a4h - 48.0:
                p2 = doc.new_page(width=a4w, height=a4h)
                fontname = _insert_font_on_page(p2, font_path)
                y = top0
            p2.insert_text((left, y), _truncate_pdf_cell(line, 100), fontname=fontname, fontsize=fs_body)
            y += lh_body

    out = doc.tobytes(deflate=True, garbage=4, clean=True)
    doc.close()
    return out
