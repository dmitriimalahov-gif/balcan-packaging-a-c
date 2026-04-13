# -*- coding: utf-8 -*-
"""
Обогащение строк каталога данными CG и knife_cache из SQLite (без Streamlit).
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

import packaging_db as pkg_db

from modules.packaging_catalog.application.makety_display import parse_qty_int_for_cg

CG_FINISH_LABELS_MAKETY: dict[str, str] = {
    "lac_wb": "Lac WB (водный лак)",
    "uv_no_foil": "UV без фольги",
    "uv_foil": "UV с фольгой",
}


def apply_makety_cg_derived_from_db(db_path: Path, rows: list[dict[str, Any]]) -> None:
    """
    Поля каталога CG, размер ножа из knife_cache и текущая цена CG (за 1000 шт., приоритет отделки lac_wb).
    При сопоставлении excel_row → cutit и наличии прайса перезаписывает item['price'].
    """
    if not rows:
        return
    if not db_path.is_file():
        for item in rows:
            item["_cg_cutit_no"] = ""
            item["_cg_knife_name"] = ""
            item["_cg_category"] = ""
            item["_cg_lacquers"] = ""
            item["_knife_size_mm"] = ""
        return
    conn = pkg_db.connect(db_path)
    try:
        pkg_db.init_db(conn)
        cg_map = pkg_db.load_cg_mapping(conn)
        cg_knives = pkg_db.load_cg_knives(conn)
        knives_by = {k["cutit_no"]: k for k in cg_knives}
        cg_prices = pkg_db.load_cg_prices(conn)
        knife_meta = pkg_db.load_knives_meta(conn)
        finish_pref = ("lac_wb", "uv_no_foil", "uv_foil")
        for item in rows:
            er = int(item["excel_row"])
            cutit = ""
            m = cg_map.get(er)
            if m:
                cutit = (m.get("cutit_no") or "").strip()
            kinfo = knives_by.get(cutit) if cutit else None
            item["_cg_cutit_no"] = cutit
            item["_cg_knife_name"] = (kinfo.get("name") or "").strip() if kinfo else ""
            item["_cg_category"] = (kinfo.get("category") or "").strip() if kinfo else ""
            pr_c = [p for p in cg_prices if p["cutit_no"] == cutit]
            fts = sorted(set(str(p["finish_type"]) for p in pr_c))
            lac_labels = [CG_FINISH_LABELS_MAKETY.get(f, f) for f in fts]
            item["_cg_lacquers"] = ", ".join(lac_labels)
            km = knife_meta.get(er)
            w0 = float(km["width_mm"]) if km else 0.0
            h0 = float(km["height_mm"]) if km else 0.0
            if km and w0 > 0 and h0 > 0:
                item["_knife_size_mm"] = f"{w0:.1f} × {h0:.1f} mm"
            else:
                item["_knife_size_mm"] = ""
            if cutit and pr_c:
                qty = parse_qty_int_for_cg(item.get("qty_per_year") or "")
                if qty <= 0:
                    qty = 1
                ft = next((f for f in finish_pref if any(p["finish_type"] == f for p in pr_c)), None)
                if ft is None:
                    ft = str(pr_c[0]["finish_type"])
                val = pkg_db.cg_price_for_qty(pr_c, ft, qty)
                if val is not None:
                    item["price"] = f"{val:.2f}"
    finally:
        conn.close()
