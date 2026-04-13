# Доменные события

Цель: слабая связность модулей и аудит (bootstrap §31). Реализация: таблица **`domain_events`** в PostgreSQL.

## Запись события

| Поле | Тип | Описание |
|------|-----|----------|
| id | BIGSERIAL | PK |
| event_type | TEXT | Например `import_completed`, `layout_exported` |
| payload | JSONB | Контекст (без PII или с согласованной маскировкой) |
| created_at | TIMESTAMPTZ | UTC |

## Планируемые типы (имена)

- `import_completed` — успешный импорт макетов / cutii.
- `layout_exported` — экспорт раскладки PDF/SVG.
- `catalog_synced` — синхронизация с Excel.

Чтение: `GET /api/v1/events/recent?limit=50` (read-only).

Правило: **факты производства и остатков** не заменяются событиями; события — дополнение для интеграций и журнала.
