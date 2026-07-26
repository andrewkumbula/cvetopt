from __future__ import annotations

import getpass
import sys
import tempfile
from pathlib import Path

from loguru import logger

from cvetopt.core.settings import EnvSettings

_CONFIGURED = False
_LOG_DIR: Path | None = None
_LOG_SUFFIX = ""


def _safe_user() -> str:
    try:
        user = getpass.getuser()
    except Exception:
        user = "user"
    return "".join(ch for ch in user if ch.isalnum() or ch in "-_") or "user"


def active_log_dir() -> Path | None:
    """Папка, куда реально удалось писать журнал (может отличаться от data/logs)."""
    return _LOG_DIR


def active_log_suffix() -> str:
    """Суффикс имени файла журнала: пустой либо «-Имя», если общий файл занят."""
    return _LOG_SUFFIX


def _candidate_dirs(project_root: Path) -> list[Path]:
    return [
        project_root / "data" / "logs",
        Path(tempfile.gettempdir()) / "cvetopt-logs",
    ]


def configure_logging(env: EnvSettings | None = None) -> None:
    """
    Настраивает loguru:
    - вывод в консоль (stderr),
    - ежедневные файлы data/logs/YYYY-MM-DD.log,
    - удаление файлов старше 30 дней.

    Файл журнала не критичен: если папка или файл недоступны (например,
    созданы другой учётной записью Windows), пробуем файл с именем пользователя,
    затем TEMP, и в крайнем случае работаем только с консольным выводом —
    сервер из-за журнала падать не должен.
    """
    global _CONFIGURED, _LOG_DIR, _LOG_SUFFIX
    if _CONFIGURED:
        return

    if env is None:
        env = EnvSettings()

    level = env.log_level.upper()
    logger.remove()
    logger.add(
        sys.stderr,
        level=level,
        enqueue=True,
        backtrace=False,
        diagnose=False,
    )

    problems: list[str] = []
    for logs_dir in _candidate_dirs(env.project_root):
        try:
            logs_dir.mkdir(parents=True, exist_ok=True)
        except OSError as e:
            problems.append(f"{logs_dir}: {e}")
            continue
        for suffix in ("", f"-{_safe_user()}"):
            pattern = str(logs_dir / f"{{time:YYYY-MM-DD}}{suffix}.log")
            try:
                logger.add(
                    pattern,
                    level=level,
                    rotation="00:00",
                    retention="30 days",
                    encoding="utf-8",
                    enqueue=True,
                    backtrace=True,
                    diagnose=False,
                )
            except OSError as e:
                problems.append(f"{pattern}: {e}")
                continue
            _LOG_DIR = logs_dir
            _LOG_SUFFIX = suffix
            if problems:
                logger.warning(
                    "Журнал пишется в {} (не удалось: {})",
                    pattern,
                    "; ".join(problems),
                )
            _CONFIGURED = True
            return

    logger.warning(
        "Журнал в файл недоступен, работаем без него: {}",
        "; ".join(problems) or "неизвестная причина",
    )
    _CONFIGURED = True
