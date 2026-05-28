from __future__ import annotations

import sys
from pathlib import Path

from loguru import logger

from cvetopt.core.settings import EnvSettings

_CONFIGURED = False


def configure_logging(env: EnvSettings | None = None) -> None:
    """
    Настраивает loguru:
    - вывод в консоль (stderr),
    - ежедневные файлы data/logs/YYYY-MM-DD.log,
    - удаление файлов старше 30 дней.
    """
    global _CONFIGURED
    if _CONFIGURED:
        return

    if env is None:
        env = EnvSettings()

    logs_dir: Path = env.project_root / "data" / "logs"
    logs_dir.mkdir(parents=True, exist_ok=True)

    logger.remove()
    logger.add(
        sys.stderr,
        level=env.log_level.upper(),
        enqueue=True,
        backtrace=False,
        diagnose=False,
    )
    logger.add(
        str(logs_dir / "{time:YYYY-MM-DD}.log"),
        level=env.log_level.upper(),
        rotation="00:00",
        retention="30 days",
        encoding="utf-8",
        enqueue=True,
        backtrace=True,
        diagnose=False,
    )
    _CONFIGURED = True
