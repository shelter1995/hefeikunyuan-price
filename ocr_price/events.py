from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from typing import Any


@dataclass(frozen=True)
class PipelineEvent:
    stage: str
    level: str
    message: str
    data: dict[str, Any] = field(default_factory=dict)
    time: str = field(default_factory=lambda: datetime.now().isoformat(timespec="seconds"))


def event_to_dict(event: PipelineEvent) -> dict[str, Any]:
    return {
        "time": event.time,
        "stage": event.stage,
        "level": event.level,
        "message": event.message,
        "data": event.data,
    }


def append_event(
    result: dict[str, Any],
    *,
    stage: str,
    level: str,
    message: str,
    data: dict[str, Any] | None = None,
) -> dict[str, Any]:
    payload = event_to_dict(
        PipelineEvent(stage=stage, level=level, message=message, data=data or {})
    )
    result.setdefault("events", []).append(payload)
    return payload
