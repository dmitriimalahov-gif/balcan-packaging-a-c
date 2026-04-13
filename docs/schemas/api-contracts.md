# API: контракты (черновик)

Базовый префикс: **`/api/v1`** (FastAPI приложение [`api/main.py`](../../api/main.py)).

## Реализовано

| Метод | Путь | Описание |
|-------|------|----------|
| GET | `/health` | Проверка живости (`{ "status": "ok" }`, используется SPA) |
| GET | `/api/v1/items` | Каталог макетов (read-only), тело `{ items, total }` |
| GET | `/api/v1/clients` | Список клиентов (ERP-заготовка, может быть пустым) |
| GET | `/api/v1/events/recent` | Последние доменные события (аудит/интеграции) |

### `CatalogItem`

Поля: `id` (nullable для SQLite), `excel_row`, `name`, `size`, `kind`, `file`, `price`, `price_new`, `qty_per_sheet`, `qty_per_year`, `gmp_code`, `updated_at`.

## Планируется (см. architecture_domain_inventory)

- `PUT /api/v1/items/{excel_row}`, импорт Excel, print layout, cutii, planner, color analytics — по мере выноса из Streamlit.

Ошибки: стабилизировать формат `{"detail": ...}` (Pydantic/FastAPI по умолчанию).
