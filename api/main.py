# -*- coding: utf-8 -*-
"""
Запуск: из корня репозитория
  PACKAGING_DB_PATH=/path/to.db uvicorn api.main:app --reload
или PostgreSQL:
  PACKAGING_DATABASE_URL=postgresql+psycopg://user:pass@localhost/db uvicorn api.main:app --reload
"""

from __future__ import annotations

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware

from modules.erp_foundation.api.router import router as erp_router
from modules.packaging_catalog.api.router import router as catalog_router

app = FastAPI(title="Packaging catalog API", version="0.2.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=[
        "http://127.0.0.1:5173",
        "http://localhost:5173",
    ],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

app.include_router(catalog_router)
app.include_router(erp_router)


@app.get("/health")
def health() -> dict[str, str]:
    return {"status": "ok"}
