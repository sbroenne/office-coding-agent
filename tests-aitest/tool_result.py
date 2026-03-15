from __future__ import annotations

from dataclasses import dataclass
from typing import Any


@dataclass(slots=True)
class ToolResult:
    """Simple result wrapper returned by in-memory simulator tools."""

    success: bool
    value: Any = None
    error: str | None = None
