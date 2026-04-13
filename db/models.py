# -*- coding: utf-8 -*-
"""Postgres-схема, зеркало основных таблиц из packaging_db (SQLite)."""

from __future__ import annotations

from datetime import datetime

from sqlalchemy import (
    JSON,
    BigInteger,
    Boolean,
    DateTime,
    Float,
    ForeignKey,
    Integer,
    PrimaryKeyConstraint,
    String,
    Text,
)
from sqlalchemy.orm import DeclarativeBase, Mapped, mapped_column


class Base(DeclarativeBase):
    pass


class PackagingItem(Base):
    __tablename__ = "packaging_items"

    id: Mapped[int] = mapped_column(BigInteger, primary_key=True, autoincrement=True)
    excel_row: Mapped[int] = mapped_column(Integer, unique=True, nullable=False, index=True)
    name: Mapped[str | None] = mapped_column(Text)
    size: Mapped[str | None] = mapped_column(Text)
    kind: Mapped[str | None] = mapped_column(Text)
    pdf_file: Mapped[str | None] = mapped_column(Text)
    price: Mapped[str | None] = mapped_column(Text)
    price_new: Mapped[str | None] = mapped_column(Text)
    qty_per_sheet: Mapped[str | None] = mapped_column(Text)
    qty_per_year: Mapped[str | None] = mapped_column(Text)
    gmp_code: Mapped[str | None] = mapped_column(Text)
    updated_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), nullable=False)


class PackagingMonthlyQty(Base):
    __tablename__ = "packaging_monthly_qty"
    __table_args__ = (PrimaryKeyConstraint("excel_row", "year", "month"),)

    excel_row: Mapped[int] = mapped_column(Integer, ForeignKey("packaging_items.excel_row"), index=True)
    year: Mapped[int] = mapped_column(Integer, nullable=False)
    month: Mapped[int] = mapped_column(Integer, nullable=False)
    qty: Mapped[float] = mapped_column(Float, nullable=False)
    source: Mapped[str | None] = mapped_column(Text)
    updated_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), nullable=False)


class CutiiConfirmation(Base):
    __tablename__ = "cutii_confirmations"

    cutii_sheet_row: Mapped[int] = mapped_column(Integer, primary_key=True)
    confirmed_excel_row: Mapped[int] = mapped_column(Integer, nullable=False)
    cutii_name: Mapped[str | None] = mapped_column(Text)
    updated_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), nullable=False)


class PrintTariff(Base):
    __tablename__ = "print_tariffs"

    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    min_sheets: Mapped[int] = mapped_column(Integer, nullable=False)
    max_sheets: Mapped[int | None] = mapped_column(Integer)
    price_per_sheet: Mapped[float] = mapped_column(Float, nullable=False)
    updated_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), nullable=False)


class PrintFinishExtra(Base):
    __tablename__ = "print_finish_extras"

    code: Mapped[str] = mapped_column(String(64), primary_key=True)
    label: Mapped[str] = mapped_column(Text, nullable=False)
    extra_per_sheet: Mapped[float] = mapped_column(Float, nullable=False)
    updated_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), nullable=False)


class KnifeCache(Base):
    __tablename__ = "knife_cache"

    excel_row: Mapped[int] = mapped_column(Integer, ForeignKey("packaging_items.excel_row"), primary_key=True)
    svg_full: Mapped[str] = mapped_column(Text, nullable=False)
    width_mm: Mapped[float] = mapped_column(Float, nullable=False)
    height_mm: Mapped[float] = mapped_column(Float, nullable=False)
    pdf_file: Mapped[str | None] = mapped_column(Text)
    updated_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), nullable=False)


class StockOnHand(Base):
    __tablename__ = "stock_on_hand"

    gmp_code: Mapped[str] = mapped_column(Text, primary_key=True)
    qty: Mapped[float] = mapped_column(Float, nullable=False)
    source: Mapped[str | None] = mapped_column(Text)
    updated_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), nullable=False)


class CgKnife(Base):
    __tablename__ = "cg_knives"

    cutit_no: Mapped[str] = mapped_column(Text, primary_key=True)
    name: Mapped[str | None] = mapped_column(Text)
    category: Mapped[str | None] = mapped_column(Text)
    cardboard: Mapped[str | None] = mapped_column(Text)
    updated_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), nullable=False)


class CgPrice(Base):
    __tablename__ = "cg_prices"
    __table_args__ = (PrimaryKeyConstraint("cutit_no", "finish_type", "min_qty"),)

    cutit_no: Mapped[str] = mapped_column(Text, nullable=False, index=True)
    finish_type: Mapped[str] = mapped_column(Text, nullable=False)
    min_qty: Mapped[int] = mapped_column(Integer, nullable=False)
    max_qty: Mapped[int | None] = mapped_column(Integer)
    price_per_1000: Mapped[float] = mapped_column(Float, nullable=False)
    price_old_per_1000: Mapped[float | None] = mapped_column(Float)
    updated_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), nullable=False)


class CgMapping(Base):
    __tablename__ = "cg_mapping"

    excel_row: Mapped[int] = mapped_column(Integer, ForeignKey("packaging_items.excel_row"), primary_key=True)
    cutit_no: Mapped[str] = mapped_column(Text, nullable=False)
    confirmed: Mapped[bool] = mapped_column(Boolean, nullable=False, default=False)
    updated_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), nullable=False)


class Client(Base):
    """ERP: клиент (заготовка; первый вертикальный срез коммерции)."""

    __tablename__ = "clients"

    id: Mapped[int] = mapped_column(BigInteger, primary_key=True, autoincrement=True)
    code: Mapped[str] = mapped_column(Text, unique=True, nullable=False, index=True)
    name: Mapped[str] = mapped_column(Text, nullable=False)
    created_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), nullable=False)
    updated_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), nullable=False)


class DomainEvent(Base):
    """Журнал доменных событий для интеграций и аудита (bootstrap §31)."""

    __tablename__ = "domain_events"

    id: Mapped[int] = mapped_column(BigInteger, primary_key=True, autoincrement=True)
    event_type: Mapped[str] = mapped_column(Text, nullable=False, index=True)
    payload: Mapped[dict | list | None] = mapped_column(JSON, nullable=True)
    created_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), nullable=False)
