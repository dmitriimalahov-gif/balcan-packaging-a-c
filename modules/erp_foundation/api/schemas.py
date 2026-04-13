# -*- coding: utf-8 -*-
from pydantic import BaseModel


class ClientOut(BaseModel):
    id: int
    code: str
    name: str
    created_at: str
    updated_at: str


class ClientsListResponse(BaseModel):
    items: list[ClientOut]


class DomainEventOut(BaseModel):
    id: int
    event_type: str
    payload: dict | list | None
    created_at: str


class DomainEventsResponse(BaseModel):
    items: list[DomainEventOut]
