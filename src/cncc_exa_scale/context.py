"""Shared runtime context."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any

import logging


@dataclass(frozen=True)
class AppContext:
    project_root: Path
    config: dict[str, Any]
    logger: logging.Logger

    def resolve_path(self, value: str | Path) -> Path:
        path = Path(value)
        if path.is_absolute():
            return path
        return self.project_root / path
