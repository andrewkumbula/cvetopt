"""
История прогонов для вкладки «Журнал»: календарь, последний запуск каждой кнопки, автозапуски.

Прогоны пишутся в data/state/runs (day_log) с момента появления этого журнала. Более ранние
восстанавливаются из дневных журналов data/logs: там нет названия кнопки, поэтому вид прогона
узнаём по характерным строкам, итог — по «Job … failed/cancelled» и последней строке «Ошибка: …».
"""

from __future__ import annotations

import re
from collections.abc import Iterable
from datetime import date, datetime, timedelta
from pathlib import Path

from loguru import logger

from cvetopt.core import day_log
from cvetopt.core.settings import ScheduleConfig

KIND_LABELS = {
    "balance_auto+delmir": "Заполнить Auto_new.xls",
    "holland_full_cycle": "Почта, Excel и перевод",
    "holland_translated": "Перевести инвойсы",
    "sklad_template_copy": "Создать шаблон",
    "mix_separation": "Разобрать миксы",
    "gypsophila_split": "Отделить гипсофилу",
    "biflorica": "Скачать отчёты Biflorica",
    "mail_attachments": "Скачать вложения из почты",
    "auto1_pipeline": "Auto1 до файла для склада",
    "holland_translate": "Перевести Голландия_1",
    "ecuador_create": "Создать Эквадор",
    "delmir_transport": "Транспорт трак с del-mir",
    "delmir_check": "Проверить del-mir",
    "selfcheck": "Проверить программу",
}

# Автозапуски из scheduler.py: что запускается по очереди. Первый шаг — признак, что автозапуск был.
AUTORUNS = {
    "weekly_chain": ("Голландия и шаблон", ["balance_auto+delmir", "holland_full_cycle", "sklad_template_copy"]),
    "biflorica": ("Отчёты Biflorica", ["biflorica"]),
}

# Порядок важен: цепочки пишут строки вложенных шагов («Шаблон:», «Перевод:»), поэтому
# сначала узнаём цепочку по её собственной строке.
_KIND_MARKERS: list[tuple[str, str]] = [
    ("selfcheck", "Проверка: тестовые файлы"),
    ("holland_full_cycle", "Шаг 1 из 3: забираю вложения"),
    ("balance_auto+delmir", "marketPlaceNew/balance"),
    ("balance_auto+delmir", "(баланс Biflorica)"),
    ("delmir_check", "Проверка del-mir:"),
    ("delmir_transport", "Транспорт трак"),
    ("mix_separation", "Миксы:"),
    ("gypsophila_split", "Гипсофила:"),
    ("holland_translated", "Переведено:"),
    ("sklad_template_copy", "Шаблон:"),
    # Скачивание Biflorica само создаёт Эквадор — проверяем раньше «Эквадор:».
    ("biflorica", "Уже в реестре заказов"),
    ("biflorica", "Архив Biflorica:"),
    ("ecuador_create", "Эквадор:"),
    ("holland_translate", "Перевод: папка склада"),
    ("auto1_pipeline", "Книга Auto_new:"),
    ("mail_attachments", "Переопределён период поиска писем"),
]

_TS = r"(\d{4}-\d\d-\d\d \d\d:\d\d:\d\d)\.\d+ \|"
_JOB_LINE = re.compile(_TS + r".*?\[job ([0-9a-f-]{36})\] \S+ \S+ \| (.*)$")
_JOB_END = re.compile(r"Job ([0-9a-f-]{36}) (failed|cancelled)")
_AUTORUN = re.compile(_TS + r".*Автозапуск: (.+)$")
_AUTORUN_GAP = timedelta(minutes=15)


def _kind_of(lines: list[str]) -> str | None:
    for kind, marker in _KIND_MARKERS:
        if any(marker in line for line in lines):
            return kind
    return None


def _autorun_id(name: str) -> str | None:
    if "Auto_new" in name:
        return "weekly_chain"
    if "Biflorica" in name:
        return "biflorica"
    return None


def parse_log(lines: Iterable[str]) -> list[dict]:
    """Прогоны из дневного журнала — в том же виде, что day_log.run_entry."""
    jobs: dict[str, dict] = {}
    ended: dict[str, str] = {}
    autoruns: list[tuple[datetime, str]] = []
    for raw in lines:
        line = raw.rstrip("\n")
        m = _JOB_LINE.match(line)
        if m:
            ts, job_id, text = m.groups()
            job = jobs.setdefault(job_id, {"first": ts, "lines": []})
            job["last"] = ts
            job["lines"].append(text)
            continue
        m = _JOB_END.search(line)
        if m:
            ended[m.group(1)] = m.group(2)
            continue
        m = _AUTORUN.match(line)
        if m and (task := _autorun_id(m.group(2))):
            autoruns.append((datetime.fromisoformat(m.group(1)), task))

    entries = []
    queue: list[str] = []
    anchor: datetime | None = None
    for job_id, job in sorted(jobs.items(), key=lambda kv: kv[1]["first"]):
        kind = _kind_of(job["lines"])
        if kind is None:
            continue
        started = datetime.fromisoformat(job["first"])
        # Автозапуск: шаги цепочки идут друг за другом сразу после строки «Автозапуск: …».
        while autoruns and autoruns[0][0] <= started:
            anchor, task = autoruns.pop(0)
            queue = list(AUTORUNS[task][1])
        auto = bool(queue) and queue[0] == kind and anchor is not None and started - anchor <= _AUTORUN_GAP
        if auto:
            queue.pop(0)
            anchor = datetime.fromisoformat(job["last"])
        else:
            queue = []

        last_line = job["lines"][-1]
        status = ended.get(job_id)
        if status is None:
            if last_line.startswith("Ошибка:"):
                status = "failed"
            elif last_line.startswith(("Прогон прерван", "Прогон остановлен")):
                status = "cancelled"
            else:
                status = "completed"
        error = ""
        if status == "failed":
            errors = [line for line in job["lines"] if line.startswith("Ошибка:")]
            error = day_log.short_error(errors[-1].removeprefix("Ошибка:").strip()) if errors else ""
        entries.append({
            "id": job_id,
            "kind": kind,
            "auto": auto,
            "status": status,
            "started": started.isoformat(timespec="seconds"),
            "error": error,
            "from_log": True,
        })
    return entries


def backfill_from_logs(root: Path, log_dir: Path | None) -> int:
    """Дописывает в журнал дней прогоны из data/logs. Повторный вызов ничего не дублирует."""
    if log_dir is None or not log_dir.is_dir():
        return 0
    added = 0
    for path in sorted(log_dir.glob("*.log")):
        try:
            with path.open(encoding="utf-8", errors="replace") as fh:
                entries = parse_log(fh)
        except OSError as e:
            logger.warning("Журнал {} не прочитан: {}", path, e)
            continue
        by_day: dict[date, list[dict]] = {}
        for entry in entries:
            by_day.setdefault(date.fromisoformat(entry["started"][:10]), []).append(entry)
        for day, day_entries in by_day.items():
            added += day_log.merge_runs(root, day, day_entries)
    return added


def history(root: Path, first: date, last: date) -> dict[str, list[dict]]:
    days: dict[str, list[dict]] = {}
    day = first
    while day <= last:
        runs = day_log.load_runs(root, day)
        if runs:
            days[day.isoformat()] = runs
        day += timedelta(days=1)
    return days


def last_runs(root: Path, before: date) -> dict[str, dict]:
    """Последний прогон каждой кнопки до дня before (не включая)."""
    result: dict[str, dict] = {}
    for day in reversed(day_log.run_days(root)):
        if day >= before:
            continue
        for entry in reversed(day_log.load_runs(root, day)):
            result.setdefault(entry["kind"], entry)
        if len(result) >= len(KIND_LABELS):
            break
    return result


def _next_due(now: datetime, weekday: int, hour: int, minute: int) -> datetime:
    due = now.replace(hour=hour, minute=minute, second=0, microsecond=0)
    due += timedelta(days=(weekday - now.weekday()) % 7)
    if due <= now:
        due += timedelta(days=7)
    return due


def autorun_overview(root: Path, schedule: ScheduleConfig, now: datetime, first: date) -> list[dict]:
    """По каждому автозапуску: расписание, следующий запуск, последний результат и пропуски с first."""
    known_days = day_log.run_days(root)
    since = max(first, known_days[0]) if known_days else now.date()
    result = []
    for task_id, (title, kinds) in AUTORUNS.items():
        cfg = getattr(schedule, task_id)
        enabled = schedule.enabled and cfg.enabled
        item: dict = {
            "id": task_id,
            "title": title,
            "steps": [KIND_LABELS[k] for k in kinds],
            "enabled": enabled,
            "weekday": cfg.weekday_index,
            "time": f"{cfg.hour:02d}:{cfg.minute:02d}",
            "next": "",
            "last": None,
            "missed": [],
        }
        if enabled:
            item["next"] = _next_due(now, cfg.weekday_index, cfg.hour, cfg.minute).isoformat(timespec="minutes")

        day = now.date()
        while day >= since:
            runs = [r for r in day_log.load_runs(root, day) if r.get("auto") and r["kind"] in kinds]
            due = datetime.combine(day, datetime.min.time()).replace(hour=cfg.hour, minute=cfg.minute)
            if runs and item["last"] is None:
                failed = [r for r in runs if r["status"] == "failed"]
                item["last"] = {
                    "day": day.isoformat(),
                    "started": runs[0]["started"],
                    "status": "failed" if failed else runs[-1]["status"],
                    "error": f"{KIND_LABELS[failed[0]['kind']]}: {failed[0]['error']}" if failed else "",
                }
            elif (
                enabled
                and not runs
                and day.weekday() == cfg.weekday_index
                and now - due > _AUTORUN_GAP
            ):
                item["missed"].append(day.isoformat())
            day -= timedelta(days=1)
        result.append(item)
    return result
