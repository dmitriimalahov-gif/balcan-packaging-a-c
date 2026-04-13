# Операционные правила агента (репозиторий a-c)

Краткий контракт для Cursor и разработчиков. **Полная версия целевой платформы:** [`docs/reference/CURSOR_ERP_BOOTSTRAP.md`](docs/reference/CURSOR_ERP_BOOTSTRAP.md).

## Миссия (кратко)

Эволюция текущего инструмента каталога макетов / печати / планирования листа в сторону **модульного ERP/MES** для упаковки: домен в центре, UI и БД — адаптеры, прогнозы не перезаписывают факты.

## Что есть в этом репозитории сейчас

- **Streamlit:** [`packaging_viewer.py`](packaging_viewer.py) — основной UI (постепенно истончаем).
- **Данные:** [`packaging_db.py`](packaging_db.py) (SQLite) и PostgreSQL через Alembic ([`db/`](db/), [`alembic/`](alembic/)).
- **HTTP API:** [`api/`](api/) (FastAPI, read-only каталог и расширения).
- **SPA:** [`web/`](web/) (React + Vite).
- **Модули домена (рост):** [`modules/`](modules/) — `packaging_catalog`, `planning`, `forecasting`.

## Инварианты (из bootstrap, применимые сразу)

- Не дублировать бизнес-правила во фронтенде — источник правды на бэкенде.
- Миграции схемы только через Alembic для PostgreSQL.
- Не ломать рабочий Streamlit без паритета в API+SPA для сценария ([`docs/STREAMLIT_AND_PRODUCTION.md`](docs/STREAMLIT_AND_PRODUCTION.md)).
- Прогноз и факт разделять в модели данных, когда появятся прогнозные сущности.

## Один чат, несколько «агентов» (Bootstrap §18)

В Cursor **одна модель на окно чата**. Несколько агентов реализованы так:

1. **[`.cursor/rules/orchestrator.mdc`](.cursor/rules/orchestrator.mdc)** (`alwaysApply: true`) — в нетривиальном ответе в начале указывай строку **`Роль: …`**; если пользователь написал префикс **`[Бэкенд]`**, `[SQL]`, `[Архитектор]` и т.д., ведущая роль **строго** по префиксу.
2. **[`.cursor/rules/role-*.mdc`](.cursor/rules/)** — восемь ролей (Архитектор, Домен, Бэкенд, Фронт, SQL, Планирование, QA, Доки) с **globs**: правила подключаются к контексту открытых файлов.
3. Тематические правила (`packaging_viewer.mdc`, `api-backend.mdc`, …) дополняют роли по путям.

Инструкция для команды: **[`docs/cursor_roles.md`](docs/cursor_roles.md)**. Полный эталон ролей: [`docs/reference/CURSOR_ERP_BOOTSTRAP.md`](docs/reference/CURSOR_ERP_BOOTSTRAP.md) §18–§19.

## Workflow задачи

1. Понять цель и затронутые модули.  
2. Короткий план (файлы, БД, API, тесты, доки).  
3. Сфокусированная реализация.  
4. Тесты / линты.  
5. Краткое резюме и риски.

## Наблюдаемость и воркер

- [`docs/observability.md`](docs/observability.md) — логи, `domain_events`, фоновые задачи (`worker/`).

## Документация

- Обзор системы: [`docs/architecture/system-overview.md`](docs/architecture/system-overview.md)
- Зависимости модулей: [`docs/architecture/module-dependencies.md`](docs/architecture/module-dependencies.md)
- Домен упаковки: [`docs/domain/packaging.md`](docs/domain/packaging.md)
- Схемы и контракты: [`docs/schemas/`](docs/schemas/)
- ADR: [`docs/adr/`](docs/adr/)

## Правила Cursor

- [`.cursor/README.md`](.cursor/README.md) — краткая карта каталога `.cursor/`.
- [`.cursor/rules/`](.cursor/rules/) — `core.mdc` + `orchestrator.mdc` всегда; остальное по glob’ам и описанию.
