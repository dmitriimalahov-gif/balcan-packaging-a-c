#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Аналитика цветов для PDF макетов (коробка, блистер, пакет, этикетка): сбор, кластеризация, экспорт."""

from __future__ import annotations

import io
import math
from collections import Counter, defaultdict
from pathlib import Path
from typing import Any

import fitz
import pandas as pd
from PIL import Image

from packaging_pdf_sheet_preview import resolve_pdf_path


CORE_PALETTE: dict[str, tuple[int, int, int]] = {
    "Красный": (220, 20, 60),
    "Синий": (25, 90, 180),
    "Жёлтый": (240, 200, 40),
    "Фиолетовый": (128, 64, 170),
    "Зелёный": (34, 139, 34),
    "Чёрный": (20, 20, 20),
}

# Приближение sRGB -> Pantone Solid Coated (для оценки, не для финальной спецификации печати).
PANTONE_APPROX: list[tuple[str, tuple[int, int, int]]] = [
    ("186 C", (200, 16, 46)),
    ("199 C", (213, 0, 50)),
    ("485 C", (218, 41, 28)),
    ("032 C", (238, 39, 55)),
    ("185 C", (228, 0, 43)),
    ("293 C", (0, 87, 184)),
    ("286 C", (0, 56, 168)),
    ("2728 C", (0, 71, 187)),
    ("Reflex Blue C", (0, 20, 137)),
    ("287 C", (0, 48, 135)),
    ("102 C", (245, 225, 39)),
    ("116 C", (255, 205, 0)),
    ("1235 C", (255, 200, 46)),
    ("2685 C", (88, 44, 131)),
    ("2597 C", (97, 32, 142)),
    ("266 C", (117, 63, 172)),
    ("355 C", (0, 154, 68)),
    ("348 C", (0, 132, 61)),
    ("361 C", (67, 176, 42)),
    ("Black C", (45, 41, 38)),
    ("Cool Gray 11 C", (83, 86, 90)),
    ("Warm Red C", (249, 66, 58)),
    ("021 C", (238, 127, 0)),
    ("1655 C", (252, 76, 2)),
    ("2925 C", (0, 167, 225)),
    ("299 C", (0, 169, 224)),
    ("375 C", (139, 197, 63)),
    ("White", (255, 255, 255)),
]


def _to_u8_rgb(color_obj: Any) -> tuple[int, int, int] | None:
    if color_obj is None:
        return None
    if isinstance(color_obj, (int, float)):
        g = int(max(0.0, min(1.0, float(color_obj))) * 255.0)
        return (g, g, g)
    if isinstance(color_obj, (list, tuple)):
        if not color_obj:
            return None
        vals = [float(x) for x in color_obj[:3]]
        if any(math.isnan(v) for v in vals):
            return None
        if max(vals) <= 1.0:
            vals = [v * 255.0 for v in vals]
        return tuple(int(max(0.0, min(255.0, v))) for v in vals)  # type: ignore[return-value]
    return None


def _rgb_to_hex(rgb: tuple[int, int, int]) -> str:
    return f"#{rgb[0]:02X}{rgb[1]:02X}{rgb[2]:02X}"


def hex_to_rgb_u8(h: str) -> tuple[int, int, int]:
    """#RRGGBB или RRGGBB → (R,G,B) 0..255; при ошибке — серый."""
    s = (h or "").strip().lstrip("#")
    if len(s) == 6:
        try:
            return (int(s[0:2], 16), int(s[2:4], 16), int(s[4:6], 16))
        except ValueError:
            pass
    return (128, 128, 128)


def rgb_to_cmyk_percent(rgb: tuple[int, int, int]) -> tuple[float, float, float, float]:
    """sRGB 0..255 -> CMYK 0..100 (без ICC-профиля)."""
    r = rgb[0] / 255.0
    g = rgb[1] / 255.0
    b = rgb[2] / 255.0
    k = 1.0 - max(r, g, b)
    if k >= 0.999:
        return (0.0, 0.0, 0.0, 100.0)
    c = (1.0 - r - k) / (1.0 - k)
    m = (1.0 - g - k) / (1.0 - k)
    y = (1.0 - b - k) / (1.0 - k)
    return (c * 100.0, m * 100.0, y * 100.0, k * 100.0)


def cmyk_percent_str(rgb: tuple[int, int, int]) -> str:
    c, m, y, k = rgb_to_cmyk_percent(rgb)
    return f"C{c:.1f} M{m:.1f} Y{y:.1f} K{k:.1f}"


def _dist_rgb(a: tuple[int, int, int], b: tuple[int, int, int]) -> float:
    return math.sqrt((a[0] - b[0]) ** 2 + (a[1] - b[1]) ** 2 + (a[2] - b[2]) ** 2)


def nearest_pantone_approx(rgb: tuple[int, int, int]) -> tuple[str, float]:
    best_name = "—"
    best_dist = 1e9
    for name, candidate in PANTONE_APPROX:
        d = _dist_rgb(rgb, candidate)
        if d < best_dist:
            best_dist = d
            best_name = name
    return best_name, round(best_dist, 2)


def _enrich_color_row(rgb: tuple[int, int, int]) -> dict[str, Any]:
    pantone_name, pantone_delta = nearest_pantone_approx(rgb)
    return {
        "hex": _rgb_to_hex(rgb),
        "cmyk": cmyk_percent_str(rgb),
        "pantone_approx": pantone_name,
        "pantone_delta": pantone_delta,
    }


def canonical_color_analytics_bucket(item: dict[str, Any]) -> str:
    """
    Категория для анализа цветов: box | blister | pack | label.
    Согласовано с kind_bucket в packaging_viewer / layout_raster_kind_bucket.
    """
    raw = (item.get("kind") or "").strip()
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


def collect_rows_for_color_bucket(rows: list[dict[str, Any]], bucket: str) -> list[dict[str, Any]]:
    """Строки с полем «Вид», попадающим в заданный bucket."""
    b = (bucket or "").strip().lower()
    return [r for r in rows if canonical_color_analytics_bucket(r) == b]


def collect_blister_rows(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Оставить только позиции блистеров."""
    return collect_rows_for_color_bucket(rows, "blister")


def _extract_drawing_colors(page: fitz.Page) -> list[tuple[tuple[int, int, int], float]]:
    out: list[tuple[tuple[int, int, int], float]] = []
    try:
        drawings = page.get_drawings()
    except Exception:
        drawings = []
    for d in drawings:
        rect = d.get("rect")
        area = 1.0
        if rect is not None:
            try:
                area = max(1.0, float(rect.width) * float(rect.height))
            except Exception:
                area = 1.0
        stroke = _to_u8_rgb(d.get("color"))
        fill = _to_u8_rgb(d.get("fill"))
        if stroke is not None:
            out.append((stroke, max(1.0, area * 0.15)))
        if fill is not None:
            out.append((fill, max(1.0, area)))
    return out


def _extract_raster_colors(page: fitz.Page, max_colors: int = 8) -> list[tuple[tuple[int, int, int], float]]:
    """Fallback по растровому представлению первой страницы."""
    out: list[tuple[tuple[int, int, int], float]] = []
    try:
        pix = page.get_pixmap(matrix=fitz.Matrix(0.12, 0.12), colorspace=fitz.csRGB, alpha=False)
    except Exception:
        return out
    if pix.width <= 0 or pix.height <= 0:
        return out
    samples = pix.samples
    ctr: Counter[tuple[int, int, int]] = Counter()
    step = 9  # разреживание
    n = len(samples)
    i = 0
    while i + 2 < n:
        rgb = (samples[i], samples[i + 1], samples[i + 2])
        # выкидываем почти белый фон
        if not (rgb[0] > 244 and rgb[1] > 244 and rgb[2] > 244):
            ctr[rgb] += 1
        i += 3 * step
    for rgb, cnt in ctr.most_common(max_colors):
        out.append((rgb, float(cnt)))
    return out


def extract_pdf_colors(pdf_path: Path, page_index: int = 0) -> dict[str, Any]:
    """
    Извлекает цвета из PDF.

    Возвращает:
    - colors: list[{"rgb": (r,g,b), "hex": "#RRGGBB", "weight": float}]
    - ok: bool
    - error: str
    """
    try:
        doc = fitz.open(str(pdf_path))
    except Exception as e:
        return {"ok": False, "error": str(e), "colors": []}
    try:
        if doc.page_count <= 0 or page_index >= doc.page_count:
            return {"ok": False, "error": "empty_pdf", "colors": []}
        page = doc.load_page(page_index)
        raw = _extract_drawing_colors(page)
        if not raw:
            raw = _extract_raster_colors(page)
        if not raw:
            return {"ok": False, "error": "no_colors", "colors": []}
        ctr: defaultdict[tuple[int, int, int], float] = defaultdict(float)
        for rgb, w in raw:
            ctr[rgb] += float(w)
        colors = [
            {"rgb": rgb, "hex": _rgb_to_hex(rgb), "weight": float(w)}
            for rgb, w in sorted(ctr.items(), key=lambda x: x[1], reverse=True)
        ]
        return {"ok": True, "error": "", "colors": colors}
    finally:
        try:
            doc.close()
        except Exception:
            pass


def cluster_colors(
    weighted_colors: list[tuple[tuple[int, int, int], float]],
    threshold: float = 28.0,
) -> list[dict[str, Any]]:
    """Простая кластеризация по RGB-расстоянию."""
    clusters: list[dict[str, Any]] = []
    for rgb, w in weighted_colors:
        target = None
        best = 1e9
        for c in clusters:
            d = _dist_rgb(rgb, c["center"])
            if d < threshold and d < best:
                best = d
                target = c
        if target is None:
            clusters.append({"center": rgb, "weight": float(w), "members": [(rgb, float(w))]})
        else:
            target["members"].append((rgb, float(w)))
            target["weight"] += float(w)
            tw = target["weight"]
            sr = sum(m[0][0] * m[1] for m in target["members"])
            sg = sum(m[0][1] * m[1] for m in target["members"])
            sb = sum(m[0][2] * m[1] for m in target["members"])
            target["center"] = (int(sr / tw), int(sg / tw), int(sb / tw))
    clusters.sort(key=lambda c: c["weight"], reverse=True)
    return [
        {
            "center": c["center"],
            "hex": _rgb_to_hex(c["center"]),
            "weight": c["weight"],
            "members": len(c["members"]),
        }
        for c in clusters
    ]


def _risk_by_delta(delta: float) -> str:
    if delta < 20:
        return "низкий"
    if delta < 45:
        return "средний"
    return "высокий"


def recommend_palette_changes(
    position_clusters: list[dict[str, Any]],
    *,
    mode: str = "core",
    dominant_palette: list[tuple[int, int, int]] | None = None,
    palette_items: list[tuple[str, tuple[int, int, int]]] | None = None,
) -> list[dict[str, Any]]:
    """Рекомендации по замене цветов для одной позиции."""
    if palette_items is not None:
        palette = list(palette_items)
    elif mode == "core":
        palette = list(CORE_PALETTE.items())
    else:
        dom = dominant_palette or []
        palette = [(f"Dominant {i+1}", rgb) for i, rgb in enumerate(dom)]
    recs: list[dict[str, Any]] = []
    for c in position_clusters:
        src = c["center"]
        if not palette:
            recs.append(
                {
                    "from_rgb": src,
                    "from_hex": c["hex"],
                    "to_rgb": src,
                    "to_hex": c["hex"],
                    "target": "same",
                    "delta": 0.0,
                    "risk": "низкий",
                    "weight": c["weight"],
                }
            )
            continue
        best_name = ""
        best_rgb = palette[0][1]
        best_d = 1e9
        for name, rgb in palette:
            d = _dist_rgb(src, rgb)
            if d < best_d:
                best_d = d
                best_name = name
                best_rgb = rgb
        recs.append(
            {
                "from_rgb": src,
                "from_hex": c["hex"],
                "to_rgb": best_rgb,
                "to_hex": _rgb_to_hex(best_rgb),
                "target": best_name,
                "delta": round(best_d, 2),
                "risk": _risk_by_delta(best_d),
                "weight": float(c["weight"]),
            }
        )
    return recs


def build_color_stats(
    blister_rows: list[dict[str, Any]],
    *,
    pdf_root: Path,
    cluster_threshold: float = 28.0,
    top_n_dominant: int = 8,
    summary_total_label: str = "Всего блистеров",
) -> dict[str, Any]:
    """
    Основной анализ по списку позиций (один вид упаковки).

    ``summary_total_label`` — первая строка сводки (например «Всего коробок»).

    Возвращает словарь с summary_df, positions_df, dominant_palette.
    """
    all_weighted: list[tuple[tuple[int, int, int], float]] = []
    positions: list[dict[str, Any]] = []
    ok_count = 0
    fail_count = 0
    for item in blister_rows:
        er = int(item["excel_row"])
        file_val = (item.get("file") or "").strip()
        p = resolve_pdf_path(pdf_root, file_val)
        if p is None or not p.is_file():
            fail_count += 1
            positions.append(
                {
                    "excel_row": er,
                    "name": (item.get("name") or "")[:140],
                    "pdf": file_val,
                    "status": "missing_pdf",
                    "top_colors": "",
                    "unique_clusters": 0,
                    "core_reco": "",
                    "dominant_reco": "",
                    "risk_core": "",
                    "risk_dominant": "",
                }
            )
            continue
        ext = extract_pdf_colors(p)
        if not ext["ok"]:
            fail_count += 1
            positions.append(
                {
                    "excel_row": er,
                    "name": (item.get("name") or "")[:140],
                    "pdf": p.name,
                    "status": f"error:{ext.get('error') or 'unknown'}",
                    "top_colors": "",
                    "unique_clusters": 0,
                    "core_reco": "",
                    "dominant_reco": "",
                    "risk_core": "",
                    "risk_dominant": "",
                }
            )
            continue
        ok_count += 1
        weighted = [(tuple(c["rgb"]), float(c["weight"])) for c in ext["colors"]]
        all_weighted.extend(weighted)
        clusters = cluster_colors(weighted, threshold=cluster_threshold)
        top_colors = ", ".join(c["hex"] for c in clusters[:5])
        positions.append(
            {
                "excel_row": er,
                "name": (item.get("name") or "")[:140],
                "pdf": p.name,
                "status": "ok",
                "top_colors": top_colors,
                "unique_clusters": len(clusters),
                "_clusters": clusters,
            }
        )

    global_clusters = cluster_colors(all_weighted, threshold=cluster_threshold)
    dominant_palette = [tuple(c["center"]) for c in global_clusters[: max(1, top_n_dominant)]]

    pos_rows: list[dict[str, Any]] = []
    for row in positions:
        clusters = row.get("_clusters") or []
        if not clusters:
            pos_rows.append(
                {
                    **{k: v for k, v in row.items() if not str(k).startswith("_")},
                    "core_reco": "",
                    "dominant_reco": "",
                    "risk_core": "",
                    "risk_dominant": "",
                }
            )
            continue
        core_rec = recommend_palette_changes(clusters, mode="core")
        dom_rec = recommend_palette_changes(clusters, mode="dominant", dominant_palette=dominant_palette)
        core_txt = "; ".join(f"{r['from_hex']}→{r['to_hex']}" for r in core_rec[:4])
        dom_txt = "; ".join(f"{r['from_hex']}→{r['to_hex']}" for r in dom_rec[:4])
        risk_core = Counter(r["risk"] for r in core_rec).most_common(1)[0][0]
        risk_dom = Counter(r["risk"] for r in dom_rec).most_common(1)[0][0]
        pos_rows.append(
            {
                **{k: v for k, v in row.items() if not str(k).startswith("_")},
                "core_reco": core_txt,
                "dominant_reco": dom_txt,
                "risk_core": risk_core,
                "risk_dominant": risk_dom,
            }
        )

    summary_rows = [
        {"Метрика": summary_total_label, "Значение": len(blister_rows)},
        {"Метрика": "Успешно прочитано PDF", "Значение": ok_count},
        {"Метрика": "Ошибок/не найдено PDF", "Значение": fail_count},
        {"Метрика": "Уникальных цветовых кластеров (общ.)", "Значение": len(global_clusters)},
    ]
    for i, c in enumerate(global_clusters[:10], start=1):
        summary_rows.append(
            {
                "Метрика": f"Top {i} цвет",
                "Значение": f"{c['hex']} (вес {int(c['weight'])}, members {c['members']})",
            }
        )

    global_rows: list[dict[str, Any]] = []
    for i, c in enumerate(global_clusters):
        rgb = tuple(c["center"])
        global_rows.append(
            {
                "rank": i + 1,
                "hex": _rgb_to_hex(rgb),
                "weight": c["weight"],
                "members": c["members"],
                "cmyk": cmyk_percent_str(rgb),
                "pantone_approx": nearest_pantone_approx(rgb)[0],
                "pantone_delta": nearest_pantone_approx(rgb)[1],
            }
        )

    palette_ref_rows: list[dict[str, Any]] = []
    for name, rgb in CORE_PALETTE.items():
        enriched = _enrich_color_row(rgb)
        palette_ref_rows.append(
            {
                "palette_type": "core",
                "name": name,
                "rank": None,
                "hex": enriched["hex"],
                "cmyk": enriched["cmyk"],
                "pantone_approx": enriched["pantone_approx"],
                "pantone_delta": enriched["pantone_delta"],
            }
        )
    for i, rgb in enumerate(dominant_palette, start=1):
        enriched = _enrich_color_row(rgb)
        palette_ref_rows.append(
            {
                "palette_type": "dominant",
                "name": f"Dominant {i}",
                "rank": i,
                "hex": enriched["hex"],
                "cmyk": enriched["cmyk"],
                "pantone_approx": enriched["pantone_approx"],
                "pantone_delta": enriched["pantone_delta"],
            }
        )

    return {
        "summary_df": pd.DataFrame(summary_rows),
        "positions_df": pd.DataFrame(pos_rows),
        "global_colors_df": pd.DataFrame(global_rows),
        "palette_reference_df": pd.DataFrame(palette_ref_rows),
        "dominant_palette": [_rgb_to_hex(x) for x in dominant_palette],
    }


def export_color_report_csv(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False).encode("utf-8-sig")


def export_color_report_xlsx(sheets: dict[str, pd.DataFrame]) -> bytes:
    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as wr:
        for name, df in sheets.items():
            safe = (name or "sheet")[:31]
            df.to_excel(wr, sheet_name=safe, index=False)
    return bio.getvalue()


def _build_dominant_palette(
    blister_rows: list[dict[str, Any]],
    *,
    pdf_root: Path,
    cluster_threshold: float,
    top_n_dominant: int,
) -> list[tuple[int, int, int]]:
    all_weighted: list[tuple[tuple[int, int, int], float]] = []
    for item in blister_rows:
        file_val = (item.get("file") or "").strip()
        p = resolve_pdf_path(pdf_root, file_val)
        if p is None or not p.is_file():
            continue
        ext = extract_pdf_colors(p)
        if not ext["ok"]:
            continue
        weighted = [(tuple(c["rgb"]), float(c["weight"])) for c in ext["colors"]]
        all_weighted.extend(weighted)
    if not all_weighted:
        return []
    global_clusters = cluster_colors(all_weighted, threshold=cluster_threshold)
    return [tuple(c["center"]) for c in global_clusters[: max(1, top_n_dominant)]]


# Серебристая «фольга» для основы блистера в превью (светлые/белые области PDF).
BLISTER_PREVIEW_SILVER_RGB: tuple[int, int, int] = (198, 200, 206)
BLISTER_SILVER_LIGHT_THRESHOLD = 244


def apply_blister_silver_foil_base(
    samples: bytes,
    *,
    silver_rgb: tuple[int, int, int] | None = None,
    light_threshold: int = BLISTER_SILVER_LIGHT_THRESHOLD,
) -> bytes:
    """Заменяет почти белые пиксели на серебристый тон (основа блистера в превью)."""
    sr, sg, sb = silver_rgb or BLISTER_PREVIEW_SILVER_RGB
    thr = int(light_threshold)
    buf = bytearray(samples)
    for i in range(0, len(buf), 3):
        r, g, b = buf[i], buf[i + 1], buf[i + 2]
        if r > thr and g > thr and b > thr:
            buf[i] = sr
            buf[i + 1] = sg
            buf[i + 2] = sb
    return bytes(buf)


def render_pdf_with_selected_colors(
    pdf_path: Path,
    selected_colors: list[tuple[int, int, int]],
    *,
    dpi: float = 96.0,
    cluster_threshold: float = 28.0,
    light_base: str = "keep",
) -> bytes | None:
    """
    Рендерит первую страницу PDF, оставляя только пиксели, принадлежащие
    кластерам из ``selected_colors``. Остальные заменяются белым.

    Возвращает PNG-байты или None при ошибке.
    """
    try:
        doc = fitz.open(str(pdf_path))
    except Exception:
        return None
    try:
        if doc.page_count <= 0:
            return None
        page = doc.load_page(0)
        zoom = max(0.2, float(dpi) / 72.0)
        pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), colorspace=fitz.csRGB, alpha=False)
        if pix.width <= 0 or pix.height <= 0:
            return None

        ext = extract_pdf_colors(pdf_path)
        if not ext["ok"]:
            return None
        weighted = [(tuple(c["rgb"]), float(c["weight"])) for c in ext["colors"]]
        clusters = cluster_colors(weighted, threshold=cluster_threshold)
        if not clusters:
            return None

        keep_centers: list[tuple[int, int, int]] = []
        for c in clusters:
            center = tuple(c["center"])
            for sel in selected_colors:
                if _dist_rgb(center, sel) < cluster_threshold:
                    keep_centers.append(center)
                    break

        raw = bytes(pix.samples)
        buf = bytearray(raw)
        match_radius = cluster_threshold * 1.5
        for i in range(0, len(buf), 3):
            rgb = (buf[i], buf[i + 1], buf[i + 2])
            if rgb[0] > 244 and rgb[1] > 244 and rgb[2] > 244:
                continue
            matched = False
            for kc in keep_centers:
                if _dist_rgb(rgb, kc) <= match_radius:
                    matched = True
                    break
            if not matched:
                buf[i] = 255
                buf[i + 1] = 255
                buf[i + 2] = 255

        out_samples = bytes(buf)
        if (light_base or "keep").strip().lower() == "silver":
            out_samples = apply_blister_silver_foil_base(out_samples)
        return _pixmap_to_png_bytes(out_samples, pix.width, pix.height)
    finally:
        try:
            doc.close()
        except Exception:
            pass


def find_first_pdf_path_in_rows(
    rows: list[dict[str, Any]],
    pdf_root: Path,
) -> tuple[Path | None, dict[str, Any] | None]:
    """Первая позиция в списке с существующим PDF."""
    for item in rows:
        file_val = (item.get("file") or "").strip()
        p = resolve_pdf_path(pdf_root, file_val)
        if p is not None and p.is_file():
            return p, item
    return None, None


def find_first_blister_pdf_path(
    blister_rows: list[dict[str, Any]],
    pdf_root: Path,
) -> tuple[Path | None, dict[str, Any] | None]:
    """Первый блистер с существующим PDF (совместимость)."""
    return find_first_pdf_path_in_rows(blister_rows, pdf_root)


def _recolor_samples_by_clusters(
    samples: bytes,
    source_centers: list[tuple[int, int, int]],
    target_centers: list[tuple[int, int, int]],
    *,
    pixel_match_radius: float,
) -> bytes:
    if not source_centers or not target_centers:
        return samples
    n = min(len(source_centers), len(target_centers))
    src = source_centers[:n]
    dst = target_centers[:n]
    buf = bytearray(samples)
    for i in range(0, len(buf), 3):
        rgb = (buf[i], buf[i + 1], buf[i + 2])
        if rgb[0] > 244 and rgb[1] > 244 and rgb[2] > 244:
            continue
        best_idx = -1
        best_d = 1e9
        for idx, center in enumerate(src):
            d = _dist_rgb(rgb, center)
            if d < best_d:
                best_d = d
                best_idx = idx
        if best_idx >= 0 and best_d <= pixel_match_radius:
            tr, tg, tb = dst[best_idx]
            buf[i] = tr
            buf[i + 1] = tg
            buf[i + 2] = tb
    return bytes(buf)


def _pixmap_to_png_bytes(samples: bytes, width: int, height: int) -> bytes:
    img = Image.frombytes("RGB", (width, height), samples)
    bio = io.BytesIO()
    img.save(bio, format="PNG")
    return bio.getvalue()


def blister_recolor_comparison_pngs(
    pdf_path: Path,
    *,
    dpi: float,
    cluster_threshold: float,
    pixel_match_radius: float,
    mode: str,
    dominant_palette: list[tuple[int, int, int]],
    palette_items: list[tuple[str, tuple[int, int, int]]] | None,
    light_base: str = "silver",
) -> tuple[bytes, bytes] | None:
    """
    Два PNG: исходник и перекраска.

    ``light_base``:
    - ``silver`` — светлые области как у блистера (серебристая фольга в превью);
    - ``keep`` — без замены белого (коробка, пакет, этикетка).
    """
    try:
        doc = fitz.open(str(pdf_path))
    except Exception:
        return None
    try:
        if doc.page_count <= 0:
            return None
        page = doc.load_page(0)
        zoom = max(0.2, float(dpi) / 72.0)
        pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), colorspace=fitz.csRGB, alpha=False)
        if pix.width <= 0 or pix.height <= 0:
            return None
        ext = extract_pdf_colors(pdf_path)
        if not ext["ok"]:
            return None
        weighted = [(tuple(c["rgb"]), float(c["weight"])) for c in ext["colors"]]
        clusters = cluster_colors(weighted, threshold=cluster_threshold)
        if not clusters:
            return None
        recs = recommend_palette_changes(
            clusters,
            mode=mode,
            dominant_palette=dominant_palette,
            palette_items=palette_items,
        )
        source_centers = [tuple(c["center"]) for c in clusters]
        target_centers = [tuple(r["to_rgb"]) for r in recs]
        raw = bytes(pix.samples)
        recolored = _recolor_samples_by_clusters(
            raw,
            source_centers,
            target_centers,
            pixel_match_radius=float(pixel_match_radius),
        )
        if (light_base or "silver").strip().lower() == "keep":
            orig_out = raw
            new_out = recolored
        else:
            orig_out = apply_blister_silver_foil_base(raw)
            new_out = apply_blister_silver_foil_base(recolored)
        return (
            _pixmap_to_png_bytes(orig_out, pix.width, pix.height),
            _pixmap_to_png_bytes(new_out, pix.width, pix.height),
        )
    finally:
        try:
            doc.close()
        except Exception:
            pass


def build_default_preview_palette_items(
    mode: str,
    blister_rows: list[dict[str, Any]],
    *,
    pdf_root: Path,
    cluster_threshold: float,
    top_n_dominant: int,
) -> list[tuple[str, tuple[int, int, int]]]:
    """Палитра по умолчанию для превью (как без ручной подстройки)."""
    if mode == "core":
        return list(CORE_PALETTE.items())
    dom = _build_dominant_palette(
        blister_rows,
        pdf_root=pdf_root,
        cluster_threshold=cluster_threshold,
        top_n_dominant=top_n_dominant,
    )
    return [(f"Dominant {i + 1}", rgb) for i, rgb in enumerate(dom)]


def build_recolor_preview_pdf_bytes(
    blister_rows: list[dict[str, Any]],
    *,
    pdf_root: Path,
    mode: str,
    cluster_threshold: float = 28.0,
    top_n_dominant: int = 8,
    item_from_1: int = 1,
    item_to_1: int | None = None,
    dpi: int = 96,
    pixel_match_radius: float = 30.0,
    palette_items: list[tuple[str, tuple[int, int, int]]] | None = None,
    light_base: str = "silver",
    preview_base_caption: str = "серебро",
) -> dict[str, Any]:
    """Собирает PDF-превью с растровой перекраской первой страницы каждой позиции.

    ``item_from_1`` / ``item_to_1`` — номера в переданном списке **включительно**, с 1.
    ``item_to_1=None`` — до конца списка.

    ``light_base``: ``silver`` (блистер) или ``keep`` (белый фон как в PDF).
    ``preview_base_caption`` — подпись в заголовке страницы PDF.
    """
    if palette_items is not None and len(palette_items) == 0:
        return {
            "ok": False,
            "error": "Палитра превью пуста — выберите хотя бы один цвет.",
            "pdf_bytes": b"",
            "generated": 0,
            "skipped": 0,
        }
    n_all = len(blister_rows)
    if n_all == 0:
        return {
            "ok": False,
            "error": "Нет позиций для превью.",
            "pdf_bytes": b"",
            "generated": 0,
            "skipped": 0,
        }
    lo1 = max(1, int(item_from_1))
    hi1 = int(item_to_1) if item_to_1 is not None else n_all
    lo1 = min(lo1, n_all)
    hi1 = min(max(hi1, lo1), n_all)
    rows = blister_rows[lo1 - 1 : hi1]
    dominant_palette: list[tuple[int, int, int]] = []
    if palette_items is None:
        dominant_palette = _build_dominant_palette(
            rows,
            pdf_root=pdf_root,
            cluster_threshold=cluster_threshold,
            top_n_dominant=top_n_dominant,
        )
    out_doc = fitz.open()
    generated = 0
    skipped = 0

    for item in rows:
        file_val = (item.get("file") or "").strip()
        pdf_path = resolve_pdf_path(pdf_root, file_val)
        if pdf_path is None or not pdf_path.is_file():
            skipped += 1
            continue

        ext = extract_pdf_colors(pdf_path)
        if not ext["ok"]:
            skipped += 1
            continue
        weighted = [(tuple(c["rgb"]), float(c["weight"])) for c in ext["colors"]]
        clusters = cluster_colors(weighted, threshold=cluster_threshold)
        if not clusters:
            skipped += 1
            continue
        recs = recommend_palette_changes(
            clusters,
            mode=mode,
            dominant_palette=dominant_palette,
            palette_items=palette_items,
        )
        source_centers = [tuple(c["center"]) for c in clusters]
        target_centers = [tuple(r["to_rgb"]) for r in recs]

        src_doc = None
        try:
            src_doc = fitz.open(str(pdf_path))
            if src_doc.page_count <= 0:
                skipped += 1
                continue
            page = src_doc.load_page(0)
            zoom = max(0.2, float(dpi) / 72.0)
            pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), colorspace=fitz.csRGB, alpha=False)
            raw = bytes(pix.samples)
            recolored_samples = _recolor_samples_by_clusters(
                raw,
                source_centers,
                target_centers,
                pixel_match_radius=float(pixel_match_radius),
            )
            if (light_base or "silver").strip().lower() == "keep":
                orig_png = _pixmap_to_png_bytes(raw, pix.width, pix.height)
                new_png = _pixmap_to_png_bytes(recolored_samples, pix.width, pix.height)
            else:
                orig_silver = apply_blister_silver_foil_base(raw)
                new_silver = apply_blister_silver_foil_base(recolored_samples)
                orig_png = _pixmap_to_png_bytes(orig_silver, pix.width, pix.height)
                new_png = _pixmap_to_png_bytes(new_silver, pix.width, pix.height)

            w_pt = 595.0
            h_pt = 842.0
            out_page = out_doc.new_page(width=w_pt, height=h_pt)
            pal_note = "custom" if palette_items else mode
            _base_note = (preview_base_caption or "серебро").strip()
            title = (
                f"Row {int(item.get('excel_row') or 0)} | "
                f"{(item.get('name') or '')[:90]} | palette={pal_note} | основа: {_base_note}"
            )
            out_page.insert_text((36, 26), title, fontsize=9)
            _left_lbl = (
                "Исходник (как в PDF)"
                if (light_base or "").strip().lower() == "keep"
                else "Исходник (серебр. основа)"
            )
            out_page.insert_text((36, 40), _left_lbl, fontsize=8)
            out_page.insert_text((303, 40), "После перекраски", fontsize=8)
            gap = 6.0
            side_w = (w_pt - 72.0 - gap) / 2.0
            y0 = 48.0
            y1 = h_pt - 28.0
            left_rect = fitz.Rect(36.0, y0, 36.0 + side_w, y1)
            right_rect = fitz.Rect(36.0 + side_w + gap, y0, w_pt - 36.0, y1)
            out_page.insert_image(left_rect, stream=orig_png, keep_proportion=True)
            out_page.insert_image(right_rect, stream=new_png, keep_proportion=True)
            generated += 1
        except Exception:
            skipped += 1
        finally:
            if src_doc is not None:
                try:
                    src_doc.close()
                except Exception:
                    pass

    if generated <= 0:
        return {
            "ok": False,
            "error": "Не удалось собрать превью: нет успешно обработанных PDF.",
            "pdf_bytes": b"",
            "generated": generated,
            "skipped": skipped,
        }
    try:
        data = out_doc.tobytes()
    finally:
        try:
            out_doc.close()
        except Exception:
            pass
    return {
        "ok": True,
        "error": "",
        "pdf_bytes": data,
        "generated": generated,
        "skipped": skipped,
    }

