# AGENTS.md

## Cursor Cloud specific instructions

### Обзор продукта

Streamlit-приложение для управления фармацевтической упаковкой компании Balkan.
Единственный сервис — без внешних БД, очередей или API.
Данные: локальная SQLite (`packaging_data.db`, автосоздаётся) + Excel (`Упаковка_макеты.xlsx`).

### Запуск приложения

```bash
export PATH="$HOME/.local/bin:$PATH"
streamlit run packaging_viewer.py --server.port 8501 --server.headless true
```

### Проверка синтаксиса (lint)

В репозитории нет конфигурации линтера. Используйте `py_compile`:

```bash
python3 -m py_compile packaging_viewer.py
```

---

### Многоуровневая архитектура модулей

Проект — плоская директория из 15 Python-файлов. Модули организованы в 3 уровня по ответственности.

#### Уровень 1 — Фундамент (утилиты без зависимостей на другие модули проекта)

| Модуль | Ответственность | Внешние зависимости |
|---|---|---|
| `packaging_sizes.py` | Парсинг/каноникализация размеров (мм), ключи группировки | — (stdlib only) |
| `packaging_db.py` | SQLite-схема (9 таблиц) + CRUD для всех сущностей | sqlite3 (stdlib) |

#### Уровень 2 — Доменная логика (зависят от уровня 1)

| Модуль | Зависит от | Ответственность |
|---|---|---|
| `packaging_pdf_sizes.py` | `packaging_sizes` | Извлечение размеров из текста PDF (pypdf/PyMuPDF) |
| `pdf_outline_to_svg.py` | — | Извлечение векторных контуров (ножей) из PDF → SVG |
| `packaging_pdf_sheet_preview.py` | `pdf_outline_to_svg` | Рендеринг превью PDF для раскладки |
| `import_cutii_forecast.py` | `packaging_db` | Импорт прогнозов cutii из внешнего Excel |
| `packaging_print_planning.py` | `packaging_sizes`, `import_cutii_forecast` | Алгоритм раскладки оттисков на печатный лист |
| `packaging_sheet_export.py` | `packaging_pdf_sheet_preview`, `pdf_outline_to_svg`, `packaging_print_planning` | Экспорт схемы печатного листа в PDF |

#### Уровень 3 — Точки входа

| Модуль | Зависит от | Тип |
|---|---|---|
| **`packaging_viewer.py`** | `packaging_db`, `packaging_sizes`, `packaging_pdf_sheet_preview`, `pdf_outline_to_svg` | **Основное Streamlit-приложение** (~5300 строк) |
| `build_packaging_excel.py` | `packaging_pdf_sizes` | CLI: сборка Excel из PDF-макетов |
| `extract_pdf_outline_svg.py` | `pdf_outline_to_svg` | CLI: извлечение контура PDF → SVG |
| `unify_packaging_sizes.py` | `packaging_db`, `packaging_sizes` | CLI: каноникализация размеров в БД/Excel |
| `refill_db_missing_sizes_from_pdf.py` | `packaging_db`, `packaging_pdf_sizes` | CLI: заполнение пустых размеров из PDF в БД |
| `refill_missing_sizes_from_pdf.py` | `packaging_db`, `packaging_pdf_sizes` | CLI: заполнение пустых размеров из PDF в Excel |
| `migrate_prochee_to_etiketka.py` | `packaging_db` | CLI: миграция «Прочее» → «Этикетка» |

#### Граф зависимостей (упрощённый)

```
packaging_viewer.py (Streamlit UI)
├── packaging_db.py              (SQLite CRUD)
├── packaging_sizes.py           (размеры, каноникализация)
├── packaging_pdf_sheet_preview.py
│   └── pdf_outline_to_svg.py    (PDF → SVG ножи)
└── [через Streamlit UI вызывает]
    ├── packaging_print_planning.py
    │   ├── packaging_sizes.py
    │   └── import_cutii_forecast.py
    │       └── packaging_db.py
    └── packaging_sheet_export.py
        ├── packaging_pdf_sheet_preview.py
        ├── pdf_outline_to_svg.py
        └── packaging_print_planning.py
```

### Вкладки Streamlit-приложения (4 основных)

| Вкладка | Функциональность |
|---|---|
| **Макеты** | Просмотр/редактирование 852+ позиций упаковки; фильтры по категориям (Коробка, Блистер, Пакет, Этикетка); редактирование размеров, видов, количеств; сохранение в Excel + SQLite |
| **Cutii-cutii → коробки** | Fuzzy-matching названий cutii к позициям каталога; ручное подтверждение сопоставлений |
| **Печать и заявки** | Конфигурация печатного листа; загрузка заявок; тарифы печати |
| **Планировщик** | Оптимизация раскладки оттисков; расчёт тиражей и стоимости; экспорт PDF со схемой листа |

### Данные и хранилище

| Файл | Описание | В git? |
|---|---|---|
| `Упаковка_макеты.xlsx` | Основной Excel-каталог упаковки (13 столбцов, лист «Макеты») | ✅ |
| `packaging_data.db` | SQLite (автосоздаётся при запуске из Excel) | ❌ gitignored |
| `*.pdf` | PDF-макеты упаковки (большой объём) | ❌ gitignored |
| `cutii_import_report.csv` | Отчёт импорта cutii | ✅ |
| `cutii_pending_review.csv` | Позиции cutii на ручное подтверждение | ✅ |

### Важные особенности для агентов

- **`packaging_viewer.py` — очень большой файл (~245k символов, ~5300 строк).** Редактируйте точечно; всегда используйте `offset`/`limit` при чтении.
- PDF-макеты отсутствуют в git — приложение работает без них, но миниатюры не отрисовываются.
- SQLite DB автоматически создаётся/мигрируется при первом запуске из данных Excel.
- CLI-скрипты (уровень 3, кроме viewer) — утилиты разовой обработки, не нуждаются в запуске для работы основного приложения.
- `requirements-viewer.txt` — единственный файл зависимостей (streamlit, pymupdf, openpyxl, rapidfuzz, pandas).
