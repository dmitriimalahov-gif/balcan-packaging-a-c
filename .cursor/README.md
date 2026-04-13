# Cursor: правила и роли проекта a-c

- **`rules/core.mdc`** — всегда включён: язык, слои, миграции, bootstrap-инварианты.
- **`rules/orchestrator.mdc`** — всегда включён: **один чат**, явная **роль** в ответе, префиксы `[Бэкенд]` и т.д.
- **`rules/role-*.mdc`** — роли из ERP Bootstrap §18; цепляются по **globs** к открытым/редактируемым файлам.
- **Тематические правила** (`packaging_viewer.mdc`, `api-backend.mdc`, …) — узкие указания по путям.

Подробно для команды: [`docs/cursor_roles.md`](../docs/cursor_roles.md).
