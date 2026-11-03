"""
Что сегодня сделано — для шагов на вкладке «Запуск». Переживает перезапуск сервера.

  data/state/runs/ГГГГ-ММ-ДД.json   — завершённые прогоны: кнопка, время, итог, причина ошибки;
  data/state/steps/ГГГГ-ММ-ДД.json  — отметки шагов: «сделано вручную» или «отметка снята».

Отметка «снята» в момент T значит: прогоны, начатые до T, шаг больше не закрывают.
Файлы старше _KEEP_DAYS удаляются при записи.
"""

from __future__ import annotations

import json
import re
import threading
from datetime import date, datetime, timedelta, timezone
from pathlib import Path

from loguru import logger

from cvetopt.core.models import JobState, JobStatus

_KEEP_DAYS = 14
_STEP_ID_RE = re.compile(r"^[a-z0-9-]{1,40}$")
STEP_DONE = "done"
STEP_UNDONE = "undone"

_lock = threading.Lock()


def job_kind(portal_id: str) -> tuple[str, bool]:
    """«auto:balance_auto+delmir:7d» → («balance_auto+delmir», True): кнопка и был ли автозапуск."""
    auto = portal_id.startswith("auto:")
    base = portal_id.removeprefix("auto:")
    return base.split(":", 1)[0], auto


def short_error(error: str | None) -> str:
    """Причина ошибки для шага на главной: без «RuntimeError: » и ссылки на журнал."""
    text = (error or "").split("Подробности в журнале", 1)[0].strip()
    head, sep, rest = text.partition(": ")
    if sep and head.isidentifier() and head.endswith(("Error", "Exception")):
        text = rest
    return text


def _local(dt: datetime) -> datetime:
    """JobState хранит время в UTC без зоны — переводим в местное время сервера."""
    return dt.replace(tzinfo=timezone.utc).astimezone().replace(tzinfo=None)


def run_entry(job: JobState) -> dict:
    kind, auto = job_kind(job.portal_id)
    return {
        "id": job.id,
        "kind": kind,
        "auto": auto,
        "status": job.status.value,
        "started": _local(job.created_at).isoformat(timespec="seconds"),
        "error": short_error(job.error) if job.status == JobStatus.failed else "",
    }


def _day_file(root: Path, folder: str, day: date) -> Path:
    return root / "data" / "state" / folder / f"{day.isoformat()}.json"


def _read(path: Path, default):
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return default


def _write(path: Path, data) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(".tmp")
    tmp.write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8")
    tmp.replace(path)
    _prune(path.parent)


def _prune(folder: Path) -> None:
    cutoff = (date.today() - timedelta(days=_KEEP_DAYS)).isoformat()
    for old in folder.glob("*.json"):
        if old.stem < cutoff:
            old.unlink(missing_ok=True)


def record_finished_run(root: Path, job: JobState) -> None:
    """Вызывается, когда прогон завершился (готово / ошибка / остановлен)."""
    entry = run_entry(job)
    day = date.fromisoformat(entry["started"][:10])
    path = _day_file(root, "runs", day)
    try:
        with _lock:
            runs = [r for r in _read(path, []) if r.get("id") != job.id]
            runs.append(entry)
            _write(path, runs)
    except OSError as e:
        # Журнал дня — подсказка для главной, прогон из-за него падать не должен.
        logger.warning("Не удалось записать прогон в {}: {}", path, e)


def load_runs(root: Path, day: date) -> list[dict]:
    return _read(_day_file(root, "runs", day), [])


def load_marks(root: Path, day: date) -> dict[str, dict]:
    return _read(_day_file(root, "steps", day), {})


def set_mark(root: Path, step_id: str, state: str | None) -> dict[str, dict]:
    """state: STEP_DONE / STEP_UNDONE, None — убрать отметку. Возвращает отметки дня."""
    if not _STEP_ID_RE.match(step_id):
        raise ValueError(f"Неизвестный шаг: {step_id!r}")
    if state not in (STEP_DONE, STEP_UNDONE, None):
        raise ValueError(f"Неизвестное состояние: {state!r}")
    path = _day_file(root, "steps", date.today())
    with _lock:
        marks = _read(path, {})
        if state is None:
            marks.pop(step_id, None)
        else:
            marks[step_id] = {"state": state, "at": datetime.now().isoformat(timespec="seconds")}
        _write(path, marks)
    return marks
