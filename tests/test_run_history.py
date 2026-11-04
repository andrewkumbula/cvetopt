import tempfile
import unittest
from datetime import date, datetime, timedelta
from pathlib import Path

from cvetopt.core import day_log, run_history
from cvetopt.core.settings import ScheduleConfig, ScheduleTaskConfig

A = "aaaaaaaa-0000-0000-0000-000000000001"
B = "bbbbbbbb-0000-0000-0000-000000000002"
C = "cccccccc-0000-0000-0000-000000000003"
D = "dddddddd-0000-0000-0000-000000000004"
E = "eeeeeeee-0000-0000-0000-000000000005"


def job(ts, job_id, text):
    return (
        f"{ts}.100 | INFO     | cvetopt.core.job_manager:append_log:120 - "
        f"[job {job_id}] {ts}.100 | {text}"
    )


# Как в журнале сервера за 03.11: автозапуск вторника, потом ручные прогоны.
LOG = [
    "2026-11-03 01:06:10.396 | INFO     | cvetopt.core.scheduler:start_scheduler:184 - "
    "Автозапуск «Auto_new.xls → Голландия → Шаблон»: включён (tuesday 09:00)",
    "2026-11-03 09:00:11.028 | INFO     | cvetopt.core.scheduler:_tick:165 - "
    "Автозапуск: Auto_new.xls → Голландия → Шаблон",
    job("2026-11-03 09:00:17", A, "Playwright proxy: http://127.0.0.1:10809"),
    job("2026-11-03 09:00:18", A, "Открываю https://ec.term.biflorica.ru/marketPlaceNew/balance"),
    "2026-11-03 09:00:20.063 | ERROR    | cvetopt.core.job_manager:run_coro_logged:202 - "
    f"Job {A} failed",
    "Traceback (most recent call last):",
    job("2026-11-03 09:00:20", A, "Ошибка: RuntimeError: Нет доступа к сайту. Подробности в журнале: C:\\x.log"),
    job("2026-11-03 09:00:20", B, "Шаг 1 из 3: забираю вложения из почты…"),
    job("2026-11-03 09:00:21", B, "Шаблон: это строка вложенного шага"),
    job("2026-11-03 09:01:11", B, "Готово: Голландия (почта → обработка → перевод)."),
    job("2026-11-03 09:01:12", C, "Шаблон: папка склада C:\\Инвойсы склад"),
    job("2026-11-03 09:01:13", C, "Готово."),
    job("2026-11-03 17:31:00", D, "Архив Biflorica: C:\\Invoice\\3\\архив"),
    job("2026-11-03 17:31:05", D, "Эквадор: отчёт BiFlorica-1.xlsx"),
    job("2026-11-03 17:40:00", E, "Миксы: папка склада C:\\Инвойсы склад"),
    job("2026-11-03 17:40:01", E, "Прогон прерван."),
]


class ParseLogTests(unittest.TestCase):
    def test_kinds_status_and_autorun(self):
        runs = {r["id"]: r for r in run_history.parse_log(LOG)}
        self.assertEqual(runs[A]["kind"], "balance_auto+delmir")
        self.assertEqual(runs[A]["status"], "failed")
        self.assertEqual(runs[A]["error"], "Нет доступа к сайту.")
        self.assertEqual(runs[B]["kind"], "holland_full_cycle")  # не «Шаблон» из вложенного шага
        self.assertEqual(runs[B]["status"], "completed")
        self.assertEqual(runs[D]["kind"], "biflorica")  # скачивание само создаёт Эквадор
        self.assertEqual(runs[E]["status"], "cancelled")
        self.assertEqual([runs[i]["auto"] for i in (A, B, C, D, E)], [True, True, True, False, False])
        self.assertEqual(runs[A]["started"], "2026-11-03T09:00:17")

    def test_manual_run_of_chain_kind_is_not_auto(self):
        late = [LOG[1], job("2026-11-03 10:30:00", A, "Шаг 1 из 3: забираю вложения из почты…")]
        self.assertFalse(run_history.parse_log(late)[0]["auto"])


class BackfillTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.logs = self.root / "data" / "logs"
        self.logs.mkdir(parents=True)
        self.day = date.today() - timedelta(days=1)
        stamp = self.day.isoformat()
        lines = [line.replace("2026-11-03", stamp) for line in LOG]
        (self.logs / f"{stamp}.log").write_text("\n".join(lines) + "\n", encoding="utf-8")

    def test_backfill_is_idempotent_and_keeps_live_records(self):
        live = {"id": C, "kind": "sklad_template_copy", "auto": True, "status": "completed",
                "started": f"{self.day.isoformat()}T09:01:12", "error": ""}
        day_log.merge_runs(self.root, self.day, [live])
        self.assertEqual(run_history.backfill_from_logs(self.root, self.logs), 4)
        self.assertEqual(run_history.backfill_from_logs(self.root, self.logs), 0)
        runs = day_log.load_runs(self.root, self.day)
        self.assertEqual(len(runs), 5)
        self.assertNotIn("from_log", next(r for r in runs if r["id"] == C))

    def test_last_runs_skip_today(self):
        run_history.backfill_from_logs(self.root, self.logs)
        last = run_history.last_runs(self.root, date.today())
        self.assertEqual(last["mix_separation"]["id"], E)
        self.assertEqual(run_history.last_runs(self.root, self.day), {})

    def test_autorun_overview_last_and_missed(self):
        run_history.backfill_from_logs(self.root, self.logs)
        week_ago = self.day - timedelta(days=7)
        # Неделю назад сервер работал, но цепочка не запускалась.
        (self.root / "data" / "state" / "runs" / f"{week_ago.isoformat()}.json").write_text("[]")
        schedule = ScheduleConfig(
            enabled=True,
            weekly_chain=ScheduleTaskConfig(weekday=_weekday_name(self.day), hour=9, minute=0),
            biflorica=ScheduleTaskConfig(enabled=False),
        )
        now = datetime.combine(date.today(), datetime.min.time()).replace(hour=12)
        overview = {a["id"]: a for a in run_history.autorun_overview(self.root, schedule, now, week_ago)}
        chain = overview["weekly_chain"]
        self.assertEqual(chain["last"]["day"], self.day.isoformat())
        self.assertEqual(chain["last"]["status"], "failed")
        self.assertTrue(chain["last"]["error"].startswith("Заполнить Auto_new.xls"))
        self.assertEqual(chain["missed"], [week_ago.isoformat()])
        self.assertEqual(chain["next"][:10], (self.day + timedelta(days=7)).isoformat())
        self.assertFalse(overview["biflorica"]["enabled"])
        self.assertEqual(overview["biflorica"]["missed"], [])


def _weekday_name(day):
    return ["monday", "tuesday", "wednesday", "thursday", "friday", "saturday", "sunday"][day.weekday()]


if __name__ == "__main__":
    unittest.main()
