# Аналитика и Metabase

SQL-витрины лежат в [`analytics/views.sql`](../analytics/views.sql). Примените их к PostgreSQL после миграции (`psql -f analytics/views.sql` или отдельная миграция Alembic).

**Metabase:** добавьте подключение к той же базе (или к read-replica с доступом только на чтение). Создайте дашборды поверх `v_packaging_items_enriched`, `v_monthly_qty_totals`, `v_catalog_with_yearly_qty`. Для тяжёлых отчётов замените `VIEW` на `MATERIALIZED VIEW` и настройте `REFRESH` по cron или через Metabase «пересчёт».

SQLite остаётся без этих представлений до переноса на Postgres.
