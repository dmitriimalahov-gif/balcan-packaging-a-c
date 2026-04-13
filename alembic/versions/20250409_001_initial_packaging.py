# -*- coding: utf-8 -*-
"""Начальная схема Postgres (зеркало SQLite из packaging_db).

Revision ID: 001_initial
Revises:
Create Date: 2025-04-09

"""

from __future__ import annotations

from typing import Sequence, Union

import sqlalchemy as sa
from alembic import op

revision: str = "001_initial"
down_revision: Union[str, None] = None
branch_labels: Union[str, Sequence[str], None] = None
depends_on: Union[str, Sequence[str], None] = None


def upgrade() -> None:
    op.create_table(
        "packaging_items",
        sa.Column("id", sa.BigInteger(), sa.Identity(always=False), nullable=False),
        sa.Column("excel_row", sa.Integer(), nullable=False),
        sa.Column("name", sa.Text(), nullable=True),
        sa.Column("size", sa.Text(), nullable=True),
        sa.Column("kind", sa.Text(), nullable=True),
        sa.Column("pdf_file", sa.Text(), nullable=True),
        sa.Column("price", sa.Text(), nullable=True),
        sa.Column("price_new", sa.Text(), nullable=True),
        sa.Column("qty_per_sheet", sa.Text(), nullable=True),
        sa.Column("qty_per_year", sa.Text(), nullable=True),
        sa.Column("gmp_code", sa.Text(), nullable=True),
        sa.Column("updated_at", sa.DateTime(timezone=True), nullable=False),
        sa.PrimaryKeyConstraint("id"),
        sa.UniqueConstraint("excel_row"),
    )

    op.create_table(
        "packaging_monthly_qty",
        sa.Column("excel_row", sa.Integer(), nullable=False),
        sa.Column("year", sa.Integer(), nullable=False),
        sa.Column("month", sa.Integer(), nullable=False),
        sa.Column("qty", sa.Float(), nullable=False),
        sa.Column("source", sa.Text(), nullable=True),
        sa.Column("updated_at", sa.DateTime(timezone=True), nullable=False),
        sa.ForeignKeyConstraint(["excel_row"], ["packaging_items.excel_row"]),
        sa.PrimaryKeyConstraint("excel_row", "year", "month"),
    )

    op.create_table(
        "cutii_confirmations",
        sa.Column("cutii_sheet_row", sa.Integer(), nullable=False),
        sa.Column("confirmed_excel_row", sa.Integer(), nullable=False),
        sa.Column("cutii_name", sa.Text(), nullable=True),
        sa.Column("updated_at", sa.DateTime(timezone=True), nullable=False),
        sa.PrimaryKeyConstraint("cutii_sheet_row"),
    )

    op.create_table(
        "print_tariffs",
        sa.Column("id", sa.Integer(), sa.Identity(always=False), nullable=False),
        sa.Column("min_sheets", sa.Integer(), nullable=False),
        sa.Column("max_sheets", sa.Integer(), nullable=True),
        sa.Column("price_per_sheet", sa.Float(), nullable=False),
        sa.Column("updated_at", sa.DateTime(timezone=True), nullable=False),
        sa.PrimaryKeyConstraint("id"),
    )

    op.create_table(
        "print_finish_extras",
        sa.Column("code", sa.String(length=64), nullable=False),
        sa.Column("label", sa.Text(), nullable=False),
        sa.Column("extra_per_sheet", sa.Float(), nullable=False),
        sa.Column("updated_at", sa.DateTime(timezone=True), nullable=False),
        sa.PrimaryKeyConstraint("code"),
    )

    op.create_table(
        "knife_cache",
        sa.Column("excel_row", sa.Integer(), nullable=False),
        sa.Column("svg_full", sa.Text(), nullable=False),
        sa.Column("width_mm", sa.Float(), nullable=False),
        sa.Column("height_mm", sa.Float(), nullable=False),
        sa.Column("pdf_file", sa.Text(), nullable=True),
        sa.Column("updated_at", sa.DateTime(timezone=True), nullable=False),
        sa.ForeignKeyConstraint(["excel_row"], ["packaging_items.excel_row"]),
        sa.PrimaryKeyConstraint("excel_row"),
    )

    op.create_table(
        "stock_on_hand",
        sa.Column("gmp_code", sa.Text(), nullable=False),
        sa.Column("qty", sa.Float(), nullable=False),
        sa.Column("source", sa.Text(), nullable=True),
        sa.Column("updated_at", sa.DateTime(timezone=True), nullable=False),
        sa.PrimaryKeyConstraint("gmp_code"),
    )

    op.create_table(
        "cg_knives",
        sa.Column("cutit_no", sa.Text(), nullable=False),
        sa.Column("name", sa.Text(), nullable=True),
        sa.Column("category", sa.Text(), nullable=True),
        sa.Column("cardboard", sa.Text(), nullable=True),
        sa.Column("updated_at", sa.DateTime(timezone=True), nullable=False),
        sa.PrimaryKeyConstraint("cutit_no"),
    )

    op.create_table(
        "cg_prices",
        sa.Column("cutit_no", sa.Text(), nullable=False),
        sa.Column("finish_type", sa.Text(), nullable=False),
        sa.Column("min_qty", sa.Integer(), nullable=False),
        sa.Column("max_qty", sa.Integer(), nullable=True),
        sa.Column("price_per_1000", sa.Float(), nullable=False),
        sa.Column("price_old_per_1000", sa.Float(), nullable=True),
        sa.Column("updated_at", sa.DateTime(timezone=True), nullable=False),
        sa.PrimaryKeyConstraint("cutit_no", "finish_type", "min_qty"),
    )

    op.create_table(
        "cg_mapping",
        sa.Column("excel_row", sa.Integer(), nullable=False),
        sa.Column("cutit_no", sa.Text(), nullable=False),
        sa.Column("confirmed", sa.Boolean(), nullable=False, server_default="false"),
        sa.Column("updated_at", sa.DateTime(timezone=True), nullable=False),
        sa.ForeignKeyConstraint(["excel_row"], ["packaging_items.excel_row"]),
        sa.PrimaryKeyConstraint("excel_row"),
    )


def downgrade() -> None:
    op.drop_table("cg_mapping")
    op.drop_table("cg_prices")
    op.drop_table("cg_knives")
    op.drop_table("stock_on_hand")
    op.drop_table("knife_cache")
    op.drop_table("print_finish_extras")
    op.drop_table("print_tariffs")
    op.drop_table("cutii_confirmations")
    op.drop_table("packaging_monthly_qty")
    op.drop_table("packaging_items")
