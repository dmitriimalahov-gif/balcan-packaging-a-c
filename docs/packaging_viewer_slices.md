# Вертикальные срезы `packaging_viewer.py`

Файл [`packaging_viewer.py`](../packaging_viewer.py) (~7k строк) разбивается по **экранам/вкладкам** и **чистым функциям**. Приоритет переноса: каталог → печать → cutii → планировщик → цвета.

## Уже вынесено в модули

| Логика | Новое расположение |
|--------|---------------------|
| Чтение каталога SQLite | [`modules/packaging_catalog/infrastructure/catalog_sqlite.py`](../modules/packaging_catalog/infrastructure/catalog_sqlite.py) |
| Чтение каталога Postgres + сценарий выбора бэкенда | `catalog_postgres.py`, [`application/catalog_read_service.py`](../modules/packaging_catalog/application/catalog_read_service.py) |
| HTTP `/api/v1/items` | [`modules/packaging_catalog/api/router.py`](../modules/packaging_catalog/api/router.py) |
| Подписи годового объёма / парсинг qty для CG | [`application/makety_display.py`](../modules/packaging_catalog/application/makety_display.py) |
| Заголовки Excel «Макеты» (`HEADERS`) | [`domain/makety_excel_config.py`](../modules/packaging_catalog/domain/makety_excel_config.py) |
| Репозитории SQLite items / monthly | [`infrastructure/items_repository.py`](../modules/packaging_catalog/infrastructure/items_repository.py), [`monthly_repository.py`](../modules/packaging_catalog/infrastructure/monthly_repository.py) |
| Контракты импорта (обёртка над `packaging_schemas`) | [`domain/import_contracts.py`](../modules/packaging_catalog/domain/import_contracts.py) |
| Нормализация заголовков Excel, карта столбцов, разбор строки | [`application/excel_headers.py`](../modules/packaging_catalog/application/excel_headers.py) |
| Чтение/запись .xlsx «Макеты» (openpyxl) | [`application/excel_io.py`](../modules/packaging_catalog/application/excel_io.py) — обогащение из БД передаётся колбэком `enrich_from_db` |

## Крупные UI-входы (`def render_*`)

| Функция | Назначение | Целевой модуль / API |
|---------|------------|----------------------|
| `render_cutii_tab` | Cutii, помесячные объёмы | `modules/packaging_catalog` + `/api/v1/cutii/*` |
| `render_print_orders_tab` | Печать, заявки | `modules/planning` + print API |
| `render_planner_tab` | Планировщик, профиль Excel | `modules/planning` |
| `render_makety_cg_supplier_prices_by_kind` | Цены CG в таблице макетов | catalog + commercial |
| `render_packaging_color_analytics*` | Анализ цветов PDF | отдельный analytics-модуль |
| `render_pdf_*` | Превью PDF | infrastructure + web viewer |

## Вспомогательные блоки без `render_`

- **`HEADERS` / `MAKETY_*`** — конфиг макетов Excel → `domain` или `application/config` каталога.
- **Загрузка/сохранение Excel** → `application/import_export` + существующие `build_packaging_excel` / `packaging_profile_excel`.
- **`main()`** — только маршрутизация вкладок и вызов сервисов.

## Следующие шаги

1. По желанию: вызывать `apply_makety_cg_derived_from_db` из [`makety_cg_enrichment.py`](../modules/packaging_catalog/application/makety_cg_enrichment.py) по умолчанию внутри `excel_io.save_*` без колбэка из viewer.
2. Подключить экран каталога в [`web/`](../web/) к write-API, когда появится `PUT /items` (сейчас в шапке SPA — индикатор `GET /health`).

Уже вынесено: `apply_makety_cg_derived_from_db`, `CG_FINISH_LABELS_MAKETY`, [`merge_kind_values_from_sqlite`](../modules/packaging_catalog/application/makety_kind_merge.py) (Streamlit-обёртка `merge_kind_from_db` остаётся в viewer).
