from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from loguru import logger


class DownloadRegistry:
    """Хранит множество order_id, по которым отчёт уже успешно скачан (на портал)."""

    def __init__(self, path: Path) -> None:
        self.path = path
        self.path.parent.mkdir(parents=True, exist_ok=True)

    def load(self) -> set[str]:
        if not self.path.exists():
            return set()
        try:
            data = json.loads(self.path.read_text(encoding="utf-8"))
        except (json.JSONDecodeError, OSError) as e:
            logger.warning("Не удалось прочитать реестр {}: {}", self.path, e)
            return set()
        if isinstance(data, dict) and "order_ids" in data:
            raw = data["order_ids"]
        elif isinstance(data, list):
            raw = data
        else:
            return set()
        return {str(x) for x in raw}

    def add(self, order_id: str) -> None:
        ids = self.load()
        ids.add(str(order_id))
        self._save_ids(ids)

    def _save_ids(self, ids: set[str]) -> None:
        payload: dict[str, Any] = {"order_ids": sorted(ids)}
        self._atomic_write_json(payload)

    def _atomic_write_json(self, payload: dict[str, Any]) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        text = json.dumps(payload, ensure_ascii=False, indent=2)
        tmp = self.path.with_suffix(".json.tmp")
        tmp.write_text(text, encoding="utf-8")
        tmp.replace(self.path)
