# AGENTS.md

## Cursor Cloud specific instructions

### Overview

This is a Python/Streamlit app for pharmaceutical packaging management ("Макеты упаковки").
Single-service architecture — no external databases, queues, or APIs required.
Data stored in local SQLite (`packaging_data.db`, auto-created) and Excel (`Упаковка_макеты.xlsx`).

### Running the app

```bash
streamlit run packaging_viewer.py --server.port 8501 --server.headless true
```

The `streamlit` binary is installed to `~/.local/bin`; ensure `PATH` includes it:

```bash
export PATH="$HOME/.local/bin:$PATH"
```

### Lint / syntax check

No dedicated linter config in the repo. Use `py_compile` for syntax validation:

```bash
python3 -m py_compile packaging_viewer.py
```

### Key files

| File | Role |
|---|---|
| `packaging_viewer.py` | Main Streamlit app (very large, ~6k+ lines) |
| `packaging_db.py` | SQLite schema + CRUD |
| `packaging_sizes.py` | Dimension parsing/canonicalization |
| `packaging_pdf_sizes.py` | Extract dimensions from PDF text |
| `packaging_print_planning.py` | Print sheet layout algorithm |
| `packaging_sheet_export.py` | Export print sheet layout to PDF |
| `pdf_outline_to_svg.py` | Extract vector knife outlines from PDF |
| `requirements-viewer.txt` | pip dependencies |

### Notes

- PDF mockup files (`*.pdf`) are gitignored due to size. The app still works without them — thumbnails won't render but all data features function normally.
- The SQLite DB file (`packaging_data.db`) is also gitignored; it is auto-created on first run from the Excel data.
- `packaging_viewer.py` is a very large single file (~245k chars). Edits should be carefully targeted.
