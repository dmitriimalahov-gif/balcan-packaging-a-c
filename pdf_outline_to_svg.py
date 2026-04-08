#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Извлечение векторных обводок из PDF (PyMuPDF get_drawings) в SVG.

Логика «нож vs мусор» (макеты Balkan / Corel):

- **Нож** — сплошная обводка препресс-цвета (магента/красный) по контуру дизайна.
- **Бегунки** — часто тот же цвет, но **пунктир** (dash pattern в PDF).
- **Размеры и стрелки** — тёмно-серые линии; не попадают в палитру ножа и отсекаются как вспомогательный серый.

Сначала берутся только **сплошные** линии цвета ножа; если их нет (редкий экспорт), повторяется отбор **с пунктиром**.
"""

from __future__ import annotations

import math
import re
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Any

import fitz

_PT_TO_MM = 25.4 / 72.0

# Эталон Corel/препресс + близкие оттенки из разных экспортов в PDF
_DEFAULT_TARGET_HEX = (
    "E61081",
    "DD0031",
    "E02020",
    "C81081",
    "D81B60",
    "E3007C",
    "EC008C",
    "ED1C24",
    "EF4136",
    "F50057",
    "FF0066",
    "E91E8C",
)

# Euclidean RGB в кубе 0..1 (макс. ~√3). Выше — шире «попадание» в цвет ножа.
DEFAULT_KNIFE_COLOR_TOLERANCE = 0.36
# Часть PDF даёт тонкие векторные обводки.
DEFAULT_KNIFE_MIN_WIDTH_PT = 0.18


def _hex_to_rgb01(h: str) -> tuple[float, float, float]:
    s = h.strip().lstrip("#")
    if len(s) != 6:
        raise ValueError(f"Нужен hex из 6 символов: {h!r}")
    return (int(s[0:2], 16) / 255.0, int(s[2:4], 16) / 255.0, int(s[4:6], 16) / 255.0)


def _parse_rgb01(stroke: Any) -> tuple[float, float, float] | None:
    if not isinstance(stroke, (list, tuple)) or len(stroke) < 3:
        return None
    try:
        r, g, b = float(stroke[0]), float(stroke[1]), float(stroke[2])
    except (TypeError, ValueError):
        return None
    return (
        max(0.0, min(1.0, r)),
        max(0.0, min(1.0, g)),
        max(0.0, min(1.0, b)),
    )


def _rgb_distance(a: tuple[float, float, float], b: tuple[float, float, float]) -> float:
    return math.sqrt((a[0] - b[0]) ** 2 + (a[1] - b[1]) ** 2 + (a[2] - b[2]) ** 2)


def _matches_any_target(
    rgb: tuple[float, float, float],
    targets: list[tuple[float, float, float]],
    tolerance: float,
) -> bool:
    return any(_rgb_distance(rgb, t) <= tolerance for t in targets)


def _is_near_gray_auxiliary(
    rgb: tuple[float, float, float],
    *,
    ref_hex: str = "34302F",
    max_dist: float = 0.14,
) -> bool:
    """Тёмно-серый препресс (размерные линии, подложки) рядом с типичным #34302F."""
    ref = _hex_to_rgb01(ref_hex)
    return _rgb_distance(rgb, ref) <= max_dist


def _drawing_uses_dash_pattern(drawing: dict[str, Any]) -> bool:
    """
    True, если у обводки задан непустой штрих (пунктир).

    PyMuPDF отдаёт ``dashes`` строкой вида ``\"[] 0\"`` (сплошная) или
    ``\"[ 7 2 2 2 ] 0\"`` (пунктир — типичные бегунки рядом с ножом).
    """
    raw = drawing.get("dashes")
    if raw is None:
        return False
    if isinstance(raw, str):
        s = raw.strip()
        if not s or s.startswith("[]"):
            return False
        if "[" not in s or "]" not in s:
            return False
        try:
            i0, i1 = s.index("["), s.index("]")
        except ValueError:
            return False
        inner = s[i0 + 1 : i1].replace(",", " ").split()
        for part in inner:
            try:
                if float(part) > 1e-6:
                    return True
            except ValueError:
                continue
        return False
    if isinstance(raw, (list, tuple)):
        if len(raw) == 0:
            return False
        spec: Any = raw
        if (
            len(raw) == 2
            and isinstance(raw[0], (list, tuple))
            and isinstance(raw[1], (int, float))
        ):
            spec = raw[0]
        for x in spec:
            if isinstance(x, (int, float)) and float(x) > 1e-6:
                return True
    return False


def _xy_from_ptlike(p: Any) -> tuple[float, float] | None:
    """PyMuPDF Point / tuple с .x .y."""
    if p is None:
        return None
    try:
        return (float(p.x), float(p.y))
    except AttributeError:
        pass
    if isinstance(p, (list, tuple)) and len(p) >= 2:
        try:
            return (float(p[0]), float(p[1]))
        except (TypeError, ValueError):
            pass
    return None


def _bbox_pdf_from_items(items: list[Any]) -> tuple[float, float, float, float] | None:
    """Ограничивающий прямоугольник по сегментам path в координатах страницы PDF (pt)."""
    min_x = min_y = float("inf")
    max_x = max_y = float("-inf")

    def add_pt(x: float, y: float) -> None:
        nonlocal min_x, min_y, max_x, max_y
        min_x = min(min_x, x)
        min_y = min(min_y, y)
        max_x = max(max_x, x)
        max_y = max(max_y, y)

    for it in items:
        if not it:
            continue
        op = it[0] if isinstance(it, (list, tuple)) and it else None
        if op == "l":
            _, p1, p2 = it
            for p in (p1, p2):
                xy = _xy_from_ptlike(p)
                if xy:
                    add_pt(xy[0], xy[1])
        elif op == "re":
            _, r = it
            add_pt(float(r.x0), float(r.y0))
            add_pt(float(r.x1), float(r.y1))
        elif op == "c":
            if len(it) == 5:
                for p in it[1:5]:
                    xy = _xy_from_ptlike(p)
                    if xy:
                        add_pt(xy[0], xy[1])
            else:
                _, curve = it
                try:
                    for i in range(0, 8, 2):
                        add_pt(float(curve[i]), float(curve[i + 1]))
                except (IndexError, TypeError, ValueError):
                    pass
        elif op == "qu" and len(it) >= 2:
            q = it[1]
            try:
                for corner in (q.ul, q.ur, q.lr, q.ll):
                    xy = _xy_from_ptlike(corner)
                    if xy:
                        add_pt(xy[0], xy[1])
            except (AttributeError, TypeError):
                pass
    if min_x == float("inf"):
        return None
    return (min_x, min_y, max_x, max_y)


def _union_pdf_rects(
    a: tuple[float, float, float, float],
    b: tuple[float, float, float, float],
) -> tuple[float, float, float, float]:
    return (min(a[0], b[0]), min(a[1], b[1]), max(a[2], b[2]), max(a[3], b[3]))


def _path_d_from_items(items: list[Any], rect: fitz.Rect) -> str:
    """Линии и кривые в координатах страницы PDF → атрибут d (ось Y как в PDF)."""
    parts: list[str] = []
    for it in items:
        if not it:
            continue
        op = it[0] if isinstance(it, (list, tuple)) and it else None
        if op == "l":
            _, p1, p2 = it
            parts.append(f"M {p1.x:.4f} {p1.y:.4f} L {p2.x:.4f} {p2.y:.4f}")
        elif op == "re":
            _, r = it
            x0, y0, x1, y1 = r.x0, r.y0, r.x1, r.y1
            parts.append(
                f"M {x0:.4f} {y0:.4f} L {x1:.4f} {y0:.4f} L {x1:.4f} {y1:.4f} L {x0:.4f} {y1:.4f} Z"
            )
        elif op == "c":
            if len(it) == 5:
                _, p0, p1, p2, p3 = it
                a = _xy_from_ptlike(p0)
                b = _xy_from_ptlike(p1)
                c_ = _xy_from_ptlike(p2)
                d = _xy_from_ptlike(p3)
                if a and b and c_ and d:
                    parts.append(
                        f"M {a[0]:.4f} {a[1]:.4f} C {b[0]:.4f} {b[1]:.4f} "
                        f"{c_[0]:.4f} {c_[1]:.4f} {d[0]:.4f} {d[1]:.4f}"
                    )
            else:
                _, curve = it
                try:
                    x0, y0 = curve[0], curve[1]
                    x1, y1 = curve[2], curve[3]
                    x2, y2 = curve[4], curve[5]
                    x3, y3 = curve[6], curve[7]
                    parts.append(
                        f"M {x0:.4f} {y0:.4f} C {x1:.4f} {y1:.4f} {x2:.4f} {y2:.4f} {x3:.4f} {y3:.4f}"
                    )
                except (IndexError, TypeError):
                    pass
        elif op == "qu" and len(it) >= 2:
            q = it[1]
            try:
                ul, ur, lr, ll = q.ul, q.ur, q.lr, q.ll
                ula, ura, lra, lla = (
                    _xy_from_ptlike(ul),
                    _xy_from_ptlike(ur),
                    _xy_from_ptlike(lr),
                    _xy_from_ptlike(ll),
                )
                if ula and ura and lra and lla:
                    parts.append(
                        f"M {ula[0]:.4f} {ula[1]:.4f} L {ura[0]:.4f} {ura[1]:.4f} "
                        f"L {lra[0]:.4f} {lra[1]:.4f} L {lla[0]:.4f} {lla[1]:.4f} Z"
                    )
            except (AttributeError, TypeError):
                pass
    if not parts:
        return ""
    return " ".join(parts)


def _svg_escape_d(s: str) -> str:
    return re.sub(r'[<>&"]', "", s)


def _pdf_document_usable(doc: fitz.Document) -> bool:
    """Документ разблокирован и первая страница читается (иначе растр/SVG часто падают тихо)."""
    try:
        if getattr(doc, "needs_pass", False) and not doc.authenticate(""):
            return False
        if doc.page_count < 1:
            return False
        p0 = doc.load_page(0)
        _ = p0.rect
        return True
    except Exception:
        return False


def open_pdf_document(path_str: str) -> fitz.Document | None:
    """
    Открыть PDF для чтения векторов/страниц.

    - Шифрование с пустым паролем: ``authenticate("")`` — без этого PyMuPDF не отдаёт контент,
      а ``get_drawings()`` пустой (типичный случай «PDF с паролем печати» из InDesign/Acrobat).
    - Если ``open(path)`` не удался или первая страница не грузится — повторное открытие из байт
      с диска (длинные пути, частично битые xref у отдельных экспортов).
    """
    doc: fitz.Document | None = None
    try:
        doc = fitz.open(path_str)
        if _pdf_document_usable(doc):
            return doc
    except Exception:
        pass
    if doc is not None:
        try:
            doc.close()
        except Exception:
            pass
        doc = None

    try:
        blob = Path(path_str).read_bytes()
    except OSError:
        return None
    if not blob:
        return None
    try:
        doc = fitz.open(stream=blob, filetype="pdf")
        if _pdf_document_usable(doc):
            return doc
    except Exception:
        pass
    if doc is not None:
        try:
            doc.close()
        except Exception:
            pass
    return None


def _open_pdf(path_str: str) -> fitz.Document | None:
    return open_pdf_document(path_str)


def _gather_knife_path_tags_and_bboxes(
    drawings: list[Any],
    page_rect: fitz.Rect,
    *,
    targets: list[tuple[float, float, float]],
    color_tolerance: float,
    min_width_pt: float,
    max_width_pt: float | None,
    exclude_gray_auxiliary: bool,
    gray_exclude_hex: str,
    output_stroke_hex: str,
    exclude_dashed_matching_targets: bool,
) -> tuple[list[str], list[tuple[float, float, float, float]]]:
    """Пути SVG и rect-ы для bbox; ``exclude_dashed_matching_targets`` отсекает бегунки."""
    out_rgb = _hex_to_rgb01(output_stroke_hex)
    R_out, G_out, B_out = int(out_rgb[0] * 255), int(out_rgb[1] * 255), int(out_rgb[2] * 255)
    path_tags: list[str] = []
    bboxes: list[tuple[float, float, float, float]] = []
    r = page_rect

    for d in drawings:
        stroke = d.get("stroke")
        if stroke is None:
            stroke = d.get("color")
        w_raw = d.get("width")
        w_pt = float(w_raw) if w_raw is not None else 1.0
        if w_pt < min_width_pt:
            continue
        if max_width_pt is not None and w_pt > max_width_pt:
            continue
        rgb = _parse_rgb01(stroke)
        if rgb is None:
            continue
        if exclude_gray_auxiliary and _is_near_gray_auxiliary(rgb, ref_hex=gray_exclude_hex):
            continue
        if not _matches_any_target(rgb, targets, color_tolerance):
            continue
        if exclude_dashed_matching_targets and _drawing_uses_dash_pattern(d):
            continue
        items = d.get("items") or []
        d_attr = _path_d_from_items(items, r)
        if not d_attr:
            continue
        dr = d.get("rect")
        if dr is not None:
            try:
                bboxes.append((dr.x0, dr.y0, dr.x1, dr.y1))
            except Exception:
                pass
        sw = max(0.05, w_pt)
        path_tags.append(
            f'<path d="{_svg_escape_d(d_attr)}" fill="none" stroke="rgb({R_out},{G_out},{B_out})" '
            f'stroke-width="{sw:.4f}" stroke-linecap="round" stroke-linejoin="round"/>'
        )
    return path_tags, bboxes


def _build_outline_group_page_mm(
    page: fitz.Page,
    *,
    target_hex_colors: list[str] | None = None,
    color_tolerance: float = DEFAULT_KNIFE_COLOR_TOLERANCE,
    min_width_pt: float = DEFAULT_KNIFE_MIN_WIDTH_PT,
    max_width_pt: float | None = None,
    exclude_gray_auxiliary: bool = True,
    gray_exclude_hex: str = "34302F",
    output_stroke_hex: str = "E61081",
) -> tuple[str, fitz.Rect, float, float, list[tuple[float, float, float, float]], str] | None:
    """
    Группа `<g>`: координаты страницы PDF → мм в системе (0,0)…(page_w_mm, page_h_mm).
    Последний элемент — только ``<path>`` без обёртки (для viewbox content).
    """
    colors = target_hex_colors if target_hex_colors else list(_DEFAULT_TARGET_HEX)
    targets = [_hex_to_rgb01(h) for h in colors]

    r = page.rect
    pw = max(r.width, 0.01)
    ph = max(r.height, 0.01)
    try:
        drawings = page.get_drawings()
    except Exception:
        drawings = []

    path_tags, bboxes = _gather_knife_path_tags_and_bboxes(
        drawings,
        r,
        targets=targets,
        color_tolerance=color_tolerance,
        min_width_pt=min_width_pt,
        max_width_pt=max_width_pt,
        exclude_gray_auxiliary=exclude_gray_auxiliary,
        gray_exclude_hex=gray_exclude_hex,
        output_stroke_hex=output_stroke_hex,
        exclude_dashed_matching_targets=True,
    )
    if not path_tags:
        path_tags, bboxes = _gather_knife_path_tags_and_bboxes(
            drawings,
            r,
            targets=targets,
            color_tolerance=color_tolerance,
            min_width_pt=min_width_pt,
            max_width_pt=max_width_pt,
            exclude_gray_auxiliary=exclude_gray_auxiliary,
            gray_exclude_hex=gray_exclude_hex,
            output_stroke_hex=output_stroke_hex,
            exclude_dashed_matching_targets=False,
        )

    if not path_tags:
        return None

    k = _PT_TO_MM
    page_w_mm = pw * k
    page_h_mm = ph * k
    inner = "\n".join(path_tags)
    group = (
        f'<g transform="translate({-r.x0 * k:.6f},{-r.y0 * k:.6f}) scale({k:.8f})">'
        f"{inner}"
        f"</g>"
    )
    return (group, r, page_w_mm, page_h_mm, bboxes, inner)


def _knife_bbox_union_pdf_points(
    page: fitz.Page,
    *,
    target_hex_colors: list[str] | None = None,
    color_tolerance: float = DEFAULT_KNIFE_COLOR_TOLERANCE,
    min_width_pt: float = DEFAULT_KNIFE_MIN_WIDTH_PT,
    max_width_pt: float | None = None,
    exclude_gray_auxiliary: bool = True,
    gray_exclude_hex: str = "34302F",
) -> tuple[float, float, float, float] | None:
    """
    Объединённый bbox отфильтрованных обводок в координатах страницы PDF (pt).
    Те же критерии, что и у контура в ``_build_outline_group_page_mm`` (в т.ч. пунктир бегунков).
    """
    colors = target_hex_colors if target_hex_colors else list(_DEFAULT_TARGET_HEX)
    targets = [_hex_to_rgb01(h) for h in colors]
    r = page.rect
    try:
        drawings = page.get_drawings()
    except Exception:
        drawings = []

    def _pass_union(skip_dashed: bool) -> tuple[float, float, float, float] | None:
        union: tuple[float, float, float, float] | None = None
        for d in drawings:
            stroke = d.get("stroke")
            if stroke is None:
                stroke = d.get("color")
            w_raw = d.get("width")
            w_pt = float(w_raw) if w_raw is not None else 1.0
            if w_pt < min_width_pt:
                continue
            if max_width_pt is not None and w_pt > max_width_pt:
                continue
            rgb = _parse_rgb01(stroke)
            if rgb is None:
                continue
            if exclude_gray_auxiliary and _is_near_gray_auxiliary(rgb, ref_hex=gray_exclude_hex):
                continue
            if not _matches_any_target(rgb, targets, color_tolerance):
                continue
            if skip_dashed and _drawing_uses_dash_pattern(d):
                continue
            items = d.get("items") or []
            d_attr = _path_d_from_items(items, r)
            if not d_attr:
                continue
            bb: tuple[float, float, float, float] | None = None
            dr = d.get("rect")
            if dr is not None:
                try:
                    bb = (float(dr.x0), float(dr.y0), float(dr.x1), float(dr.y1))
                except Exception:
                    bb = None
            ib = _bbox_pdf_from_items(items)
            if ib is not None:
                bb = _union_pdf_rects(bb, ib) if bb is not None else ib
            elif bb is None:
                continue
            union = _union_pdf_rects(union, bb) if union is not None else bb
        return union

    u = _pass_union(True)
    if u is None:
        u = _pass_union(False)
    return u


def knife_bbox_union_pdf_points(
    page: fitz.Page,
    *,
    target_hex_colors: list[str] | None = None,
    color_tolerance: float = DEFAULT_KNIFE_COLOR_TOLERANCE,
    min_width_pt: float = DEFAULT_KNIFE_MIN_WIDTH_PT,
    max_width_pt: float | None = None,
    exclude_gray_auxiliary: bool = True,
    gray_exclude_hex: str = "34302F",
) -> tuple[float, float, float, float] | None:
    """
    Объединённый bbox отфильтрованных обводок в координатах страницы PDF (pt).
    Те же фильтры, что у экспорта контура. Для открытого документа — без повторного чтения файла.
    """
    return _knife_bbox_union_pdf_points(
        page,
        target_hex_colors=target_hex_colors,
        color_tolerance=color_tolerance,
        min_width_pt=min_width_pt,
        max_width_pt=max_width_pt,
        exclude_gray_auxiliary=exclude_gray_auxiliary,
        gray_exclude_hex=gray_exclude_hex,
    )


def knife_bbox_pdf_union_points(
    path_str: str,
    page_index: int = 0,
    **kwargs: Any,
) -> tuple[float, float, float, float] | None:
    """
    Union bbox в координатах страницы PDF (pt) по пути к файлу.
    Параметры фильтра — как у ``knife_bbox_mm_from_pdf`` / ``extract_outline_svg_*``.
    """
    doc = _open_pdf(path_str)
    if doc is None:
        return None
    try:
        if doc.page_count < 1 or page_index < 0 or page_index >= doc.page_count:
            return None
        page = doc.load_page(page_index)
        return knife_bbox_union_pdf_points(
            page,
            target_hex_colors=kwargs.get("target_hex_colors"),
            color_tolerance=float(kwargs.get("color_tolerance", DEFAULT_KNIFE_COLOR_TOLERANCE)),
            min_width_pt=float(kwargs.get("min_width_pt", DEFAULT_KNIFE_MIN_WIDTH_PT)),
            max_width_pt=kwargs.get("max_width_pt"),
            exclude_gray_auxiliary=bool(kwargs.get("exclude_gray_auxiliary", True)),
            gray_exclude_hex=str(kwargs.get("gray_exclude_hex", "34302F")),
        )
    except Exception:
        return None
    finally:
        try:
            doc.close()
        except Exception:
            pass


def knife_bbox_mm_from_pdf(
    path_str: str,
    page_index: int = 0,
    **kwargs: Any,
) -> tuple[float, float] | None:
    """
    Габариты одного «ножа» в мм: ширина и высота по ограничивающему прямоугольнику
    отфильтрованных обводок (те же параметры, что у ``extract_outline_svg_*``).

    Возвращает ``(width_mm, height_mm)`` или ``None``, если линий нет или PDF недоступен.
    """
    doc = _open_pdf(path_str)
    if doc is None:
        return None
    try:
        if doc.page_count < 1 or page_index < 0 or page_index >= doc.page_count:
            return None
        page = doc.load_page(page_index)
        u = knife_bbox_union_pdf_points(
            page,
            target_hex_colors=kwargs.get("target_hex_colors"),
            color_tolerance=float(kwargs.get("color_tolerance", DEFAULT_KNIFE_COLOR_TOLERANCE)),
            min_width_pt=float(kwargs.get("min_width_pt", DEFAULT_KNIFE_MIN_WIDTH_PT)),
            max_width_pt=kwargs.get("max_width_pt"),
            exclude_gray_auxiliary=bool(kwargs.get("exclude_gray_auxiliary", True)),
            gray_exclude_hex=str(kwargs.get("gray_exclude_hex", "34302F")),
        )
        if u is None:
            return None
        x0, y0, x1, y1 = u
        k = _PT_TO_MM
        w_mm = max(0.0, (x1 - x0) * k)
        h_mm = max(0.0, (y1 - y0) * k)
        if w_mm <= 0 or h_mm <= 0:
            return None
        return (w_mm, h_mm)
    except Exception:
        return None
    finally:
        try:
            doc.close()
        except Exception:
            pass


def inner_outline_from_stored_knife_svg(
    svg_full: str,
    slot_w_mm: float,
    slot_h_mm: float,
    knife_w_mm: float,
    knife_h_mm: float,
) -> str:
    """
    Фрагмент ``<g transform=...>…</g>`` для слота схемы листа из уже сохранённого SVG ножа
    (габариты ``knife_w_mm × knife_h_mm`` в мм), без чтения PDF.
    """
    if slot_w_mm <= 0 or slot_h_mm <= 0 or knife_w_mm <= 0 or knife_h_mm <= 0:
        return ""
    raw = (svg_full or "").strip()
    if not raw:
        return ""
    try:
        root = ET.fromstring(raw)
    except ET.ParseError:
        return ""
    local = root.tag.split("}")[-1] if "}" in root.tag else root.tag
    if (local or "").lower() != "svg":
        return ""
    inner_parts: list[str] = []
    for child in list(root):
        inner_parts.append(ET.tostring(child, encoding="unicode"))
    if not inner_parts:
        return ""
    inner = "".join(inner_parts)
    s = min(float(slot_w_mm) / knife_w_mm, float(slot_h_mm) / knife_h_mm)
    s = max(0.0001, min(s, 1000.0))
    ox = (float(slot_w_mm) - knife_w_mm * s) / 2.0
    oy = (float(slot_h_mm) - knife_h_mm * s) / 2.0
    return (
        f'<g transform="translate({ox:.4f},{oy:.4f}) scale({s:.8f})">'
        f"{inner}"
        f"</g>"
    )


def extract_outline_svg_inner_for_slot(
    path_str: str,
    slot_w_mm: float,
    slot_h_mm: float,
    page_index: int = 0,
    **kwargs: Any,
) -> str:
    """
    Фрагмент `<g>...</g>` для вставки в слот печатного листа (мм): контур вписан в
    ``slot_w_mm × slot_h_mm`` как ``contain``. Пустая строка, если нет линий или ошибка.
    """
    if slot_w_mm <= 0 or slot_h_mm <= 0:
        return ""
    doc = _open_pdf(path_str)
    if doc is None:
        return ""
    try:
        if doc.page_count < 1 or page_index < 0 or page_index >= doc.page_count:
            return ""
        page = doc.load_page(page_index)
        built = _build_outline_group_page_mm(
            page,
            target_hex_colors=kwargs.get("target_hex_colors"),
            color_tolerance=float(kwargs.get("color_tolerance", DEFAULT_KNIFE_COLOR_TOLERANCE)),
            min_width_pt=float(kwargs.get("min_width_pt", DEFAULT_KNIFE_MIN_WIDTH_PT)),
            max_width_pt=kwargs.get("max_width_pt"),
            exclude_gray_auxiliary=bool(kwargs.get("exclude_gray_auxiliary", True)),
            gray_exclude_hex=str(kwargs.get("gray_exclude_hex", "34302F")),
            output_stroke_hex=str(kwargs.get("output_stroke_hex", "E61081")),
        )
        if built is None:
            return ""
        group_page, _r, page_w_mm, page_h_mm, _b, _inner = built
        if page_w_mm <= 0 or page_h_mm <= 0:
            return ""
        s = min(slot_w_mm / page_w_mm, slot_h_mm / page_h_mm)
        s = max(0.0001, min(s, 1000.0))
        ox = (slot_w_mm - page_w_mm * s) / 2.0
        oy = (slot_h_mm - page_h_mm * s) / 2.0
        return (
            f'<g transform="translate({ox:.4f},{oy:.4f}) scale({s:.8f})">'
            f"{group_page}"
            f"</g>"
        )
    except Exception:
        return ""
    finally:
        try:
            doc.close()
        except Exception:
            pass


def extract_knife_data_from_pdf(
    path_str: str,
    page_index: int = 0,
    *,
    target_hex_colors: list[str] | None = None,
    color_tolerance: float = DEFAULT_KNIFE_COLOR_TOLERANCE,
    min_width_pt: float = DEFAULT_KNIFE_MIN_WIDTH_PT,
    max_width_pt: float | None = None,
    exclude_gray_auxiliary: bool = True,
    gray_exclude_hex: str = "34302F",
    output_stroke_hex: str = "E61081",
) -> tuple[str, float, float] | None:
    """
    Одно открытие PDF → (svg_full, width_mm, height_mm) или None.
    SVG с viewbox="content". Заменяет раздельные вызовы
    knife_bbox_mm_from_pdf + extract_outline_svg_from_pdf.
    """
    doc = _open_pdf(path_str)
    if doc is None:
        return None
    try:
        if doc.page_count < 1 or page_index < 0 or page_index >= doc.page_count:
            return None
        page = doc.load_page(page_index)

        u = knife_bbox_union_pdf_points(
            page,
            target_hex_colors=target_hex_colors,
            color_tolerance=color_tolerance,
            min_width_pt=min_width_pt,
            max_width_pt=max_width_pt,
            exclude_gray_auxiliary=exclude_gray_auxiliary,
            gray_exclude_hex=gray_exclude_hex,
        )
        if u is None:
            return None
        x0, y0, x1, y1 = u
        k = _PT_TO_MM
        w_mm = max(0.0, (x1 - x0) * k)
        h_mm = max(0.0, (y1 - y0) * k)
        if w_mm <= 0 or h_mm <= 0:
            return None

        built = _build_outline_group_page_mm(
            page,
            target_hex_colors=target_hex_colors,
            color_tolerance=color_tolerance,
            min_width_pt=min_width_pt,
            max_width_pt=max_width_pt,
            exclude_gray_auxiliary=exclude_gray_auxiliary,
            gray_exclude_hex=gray_exclude_hex,
            output_stroke_hex=output_stroke_hex,
        )
        if built is None:
            return None
        _group, _r, _pw, _ph, bboxes, inner_paths = built

        if bboxes:
            min_bx = min(b[0] for b in bboxes)
            min_by = min(b[1] for b in bboxes)
            max_bx = max(b[2] for b in bboxes)
            max_by = max(b[3] for b in bboxes)
            vb_w = max(0.01, (max_bx - min_bx) * k)
            vb_h = max(0.01, (max_by - min_by) * k)
            svg_body = (
                f'<g transform="translate({-min_bx * k:.6f},{-min_by * k:.6f}) scale({k:.8f})">'
                f"{inner_paths}"
                f"</g>"
            )
        else:
            vb_w, vb_h = _pw, _ph
            svg_body = _group

        svg_full = (
            f'<?xml version="1.0" encoding="UTF-8"?>\n'
            f'<svg xmlns="http://www.w3.org/2000/svg" '
            f'width="{vb_w:.4f}mm" height="{vb_h:.4f}mm" '
            f'viewBox="0 0 {vb_w:.4f} {vb_h:.4f}" '
            f'style="shape-rendering:geometricPrecision;fill-rule:evenodd">\n'
            f"<title>outline from PDF</title>\n"
            f"{svg_body}\n"
            f"</svg>"
        )
        return (svg_full, w_mm, h_mm)
    except Exception:
        return None
    finally:
        try:
            doc.close()
        except Exception:
            pass


def try_extract_knife_from_pdf(
    path_str: str,
    page_index: int = 0,
    *,
    target_hex_colors: list[str] | None = None,
    output_stroke_hex: str = "E61081",
) -> tuple[str, float, float] | None:
    """
    Извлечь нож из PDF: сначала стандартные допуски, при пустом результате — второй проход
    с более мягким цветом и более тонкими линиями (часть наших PDF не попадает в первый проход).
    """
    r = extract_knife_data_from_pdf(
        path_str,
        page_index,
        target_hex_colors=target_hex_colors,
        output_stroke_hex=output_stroke_hex,
    )
    if r is not None:
        return r
    return extract_knife_data_from_pdf(
        path_str,
        page_index,
        target_hex_colors=target_hex_colors,
        color_tolerance=min(0.58, DEFAULT_KNIFE_COLOR_TOLERANCE + 0.2),
        min_width_pt=0.1,
        exclude_gray_auxiliary=True,
        output_stroke_hex=output_stroke_hex,
    )


def extract_outline_svg_from_pdf(
    path_str: str,
    page_index: int = 0,
    *,
    target_hex_colors: list[str] | None = None,
    color_tolerance: float = DEFAULT_KNIFE_COLOR_TOLERANCE,
    min_width_pt: float = DEFAULT_KNIFE_MIN_WIDTH_PT,
    max_width_pt: float | None = None,
    exclude_gray_auxiliary: bool = True,
    gray_exclude_hex: str = "34302F",
    viewbox: str = "page",
    output_stroke_hex: str = "E61081",
) -> str:
    """
    Возвращает полный XML/SVG или пустую строку, если ничего не отобрано.

    ``viewbox``: ``\"page\"`` — холст размером со страницу в мм; ``\"content\"`` — обрезка по bbox из rect у drawing.
    """
    if viewbox not in ("page", "content"):
        viewbox = "page"

    doc = _open_pdf(path_str)
    if doc is None:
        return ""
    try:
        if doc.page_count < 1 or page_index < 0 or page_index >= doc.page_count:
            return ""
        page = doc.load_page(page_index)
        built = _build_outline_group_page_mm(
            page,
            target_hex_colors=target_hex_colors,
            color_tolerance=color_tolerance,
            min_width_pt=min_width_pt,
            max_width_pt=max_width_pt,
            exclude_gray_auxiliary=exclude_gray_auxiliary,
            gray_exclude_hex=gray_exclude_hex,
            output_stroke_hex=output_stroke_hex,
        )
        if built is None:
            return ""
        group, r, page_w_mm, page_h_mm, bboxes, inner_paths = built
        k = _PT_TO_MM

        if viewbox == "content" and bboxes:
            min_x = min(b[0] for b in bboxes)
            min_y = min(b[1] for b in bboxes)
            max_x = max(b[2] for b in bboxes)
            max_y = max(b[3] for b in bboxes)
            vb_w = max(0.01, (max_x - min_x) * k)
            vb_h = max(0.01, (max_y - min_y) * k)
            svg_body = (
                f'<g transform="translate({-min_x * k:.6f},{-min_y * k:.6f}) scale({k:.8f})">'
                f"{inner_paths}"
                f"</g>"
            )
        else:
            vb_w, vb_h = page_w_mm, page_h_mm
            svg_body = group

        return (
            f'<?xml version="1.0" encoding="UTF-8"?>\n'
            f'<svg xmlns="http://www.w3.org/2000/svg" '
            f'width="{vb_w:.4f}mm" height="{vb_h:.4f}mm" '
            f'viewBox="0 0 {vb_w:.4f} {vb_h:.4f}" '
            f'style="shape-rendering:geometricPrecision;fill-rule:evenodd">\n'
            f"<title>outline from PDF</title>\n"
            f"{svg_body}\n"
            f"</svg>"
        )
    except Exception:
        return ""
    finally:
        try:
            doc.close()
        except Exception:
            pass
