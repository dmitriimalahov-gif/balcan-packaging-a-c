# Наблюдаемость и фоновые задачи

## Логи

- API (uvicorn): настройте уровень логирования через `UVICORN_LOG_LEVEL`.
- Streamlit: сообщения в консоль; для prod предпочтительно вынести сценарии в API+worker.

## Аудит и события

- Таблица `domain_events` (PostgreSQL) — журнал для `import_completed`, `layout_exported` и т.д.
- Запись из кода: [`modules.erp_foundation.application.events_service.record_domain_event`](../modules/erp_foundation/application/events_service.py).
- Чтение: `GET /api/v1/events/recent` (при настроенном `PACKAGING_DATABASE_URL`).

## Воркер

- Каталог [`worker/`](../worker/): точка входа `python -m worker.runner`.
- Рекомендуемая интеграция: **Redis + RQ** или **Celery** для задач анализа PDF, массового экспорта Excel, прогонов планировщика.
- Зависимости: черновик в [`requirements-worker.txt`](../requirements-worker.txt).

## Трассировка

- При росте нагрузки: OpenTelemetry для FastAPI и воркеров; корреляция `job_id` в логах и в `domain_events.payload`.
