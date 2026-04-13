-- Витрины для аналитики в PostgreSQL (Materialized VIEW при необходимости обновлять по расписанию).
-- Подключение BI: Metabase / Lightdash к той же БД или read-replica (см. docs/analytics_metabase.md).

-- Каталог с признаками заполненности
CREATE OR REPLACE VIEW v_packaging_items_enriched AS
SELECT
    pi.id,
    pi.excel_row,
    pi.name,
    pi.size,
    pi.kind,
    pi.pdf_file,
    pi.gmp_code,
    (pi.pdf_file IS NOT NULL AND TRIM(pi.pdf_file) <> '') AS has_pdf,
    (pi.gmp_code IS NOT NULL AND TRIM(pi.gmp_code) <> '') AS has_gmp,
    pi.updated_at
FROM packaging_items pi;

-- Агрегат помесячных объёмов по позиции
CREATE OR REPLACE VIEW v_monthly_qty_totals AS
SELECT
    excel_row,
    year,
    SUM(qty) AS qty_year,
    MAX(updated_at) AS last_updated
FROM packaging_monthly_qty
GROUP BY excel_row, year;

-- Сводка: позиция + суммарный год (текущий max year в данных можно фильтровать в BI)
CREATE OR REPLACE VIEW v_catalog_with_yearly_qty AS
SELECT
    e.*,
    y.year,
    y.qty_year
FROM v_packaging_items_enriched e
LEFT JOIN v_monthly_qty_totals y ON y.excel_row = e.excel_row;
