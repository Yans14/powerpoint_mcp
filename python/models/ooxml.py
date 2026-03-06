from __future__ import annotations

from datetime import datetime
from typing import Any

from pydantic import BaseModel, Field


class BridgeRequest(BaseModel):
    id: str
    method: str
    params: dict[str, Any] = Field(default_factory=dict)


class BridgeErrorPayload(BaseModel):
    code: str
    message: str
    details: dict[str, Any] | None = None


class BridgeSuccessResponse(BaseModel):
    id: str
    result: Any


class BridgeFailureResponse(BaseModel):
    id: str
    error: BridgeErrorPayload


class SlideSummary(BaseModel):
    index: int
    title: str
    layout: str
    shape_count: int


class PresentationState(BaseModel):
    presentation_id: str
    slide_count: int
    slides: list[SlideSummary]


class PresentationSession(BaseModel):
    id: str
    original_path: str
    working_path: str
    engine: str
    dirty: bool = False
    opened_at: datetime
