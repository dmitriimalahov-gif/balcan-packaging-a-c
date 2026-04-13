# -*- coding: utf-8 -*-
"""
Заготовка воркера для тяжёлых задач (bootstrap §11).

Подключение RQ/Celery — по выбору команды; пока CLI-заглушка для проверки окружения.
"""

from __future__ import annotations

import argparse


def main() -> int:
    p = argparse.ArgumentParser(description="Packaging background worker (stub)")
    p.parse_args()
    print(
        "worker: заглушка. Подключите RQ/Celery и очередь Redis — см. requirements-worker.txt "
        "и docs/observability.md"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
