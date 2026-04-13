# Схема БД: модели и источники правды

## PostgreSQL (Alembic)

Источник: [`db/models.py`](../../db/models.py), миграции в [`alembic/versions/`](../../alembic/versions/).

Основные таблицы (каталог и справочники):

- `packaging_items` — позиции каталога (`id` BIGSERIAL, `excel_row` UNIQUE).
- `packaging_monthly_qty`, `cutii_confirmations`, `print_tariffs`, `print_finish_extras`.
- `knife_cache`, `stock_on_hand`, `cg_knives`, `cg_prices`, `cg_mapping`.
- `clients` — контрагенты (`code` UNIQUE, `name`).
- `domain_events` — журнал событий (`event_type`, `payload` JSON, `created_at`).

Миграция: `002_clients_events` (после `001_initial`).

Команды:

```bash
export PACKAGING_DATABASE_URL=postgresql+psycopg://user:pass@host/dbname
alembic upgrade head
```

## SQLite (legacy)

Источник: [`packaging_db.py`](../../packaging_db.py) — `init_db`, `upsert_all`, таблицы по именам как в Postgres где применимо.

Используется Streamlit и read-only API при отсутствии `PACKAGING_DATABASE_URL`.

## Витрины аналитики (Postgres)

Файл [`analytics/views.sql`](../../analytics/views.sql) — представления для BI (Metabase и т.д.).
