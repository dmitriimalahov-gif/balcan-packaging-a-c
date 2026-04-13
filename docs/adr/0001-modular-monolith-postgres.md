# ADR-0001: Модульный монолит и PostgreSQL для новых сущностей

## Статус

Принято

## Контекст

Репозиторий вырос из Streamlit-приложения с SQLite. Целевая платформа — ERP/MES с трассируемостью и аналитикой ([`docs/reference/CURSOR_ERP_BOOTSTRAP.md`](../reference/CURSOR_ERP_BOOTSTRAP.md)).

## Решение

1. Развивать **модульный монолит** в каталоге `modules/` с разделением `domain` / `application` / `infrastructure` / `api` по модулям.
2. **PostgreSQL + Alembic** — канон для новых таблиц и ERP-расширений; SQLite остаётся для текущего Streamlit и API до полного паритета.
3. **Доменная модель в центре**; Streamlit и HTTP — тонкие адаптеры.
4. Доменные **события** фиксировать в `domain_events` (outbox/event log) по мере появления сценариев.

## Последствия

- Дублирование схемы SQLite/Postgres допустимо переходный период; миграция данных — скриптами (см. `scripts/migrate_sqlite_to_postgres.py`).
- Новые фичи предпочтительно не добавлять в `packaging_viewer.py` цельными блоками — выносить в `modules/*`.
