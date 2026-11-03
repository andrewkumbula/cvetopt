import json
import tempfile
import unittest
from datetime import date, datetime, timedelta
from pathlib import Path

from cvetopt.core import day_log
from cvetopt.core.models import JobState, JobStatus


class DayLogTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)

    def finished_job(self, portal_id, status, error=None):
        job = JobState(id=f"id-{portal_id}-{status.value}", portal_id=portal_id)
        job.status = status
        job.error = error
        return job

    def test_finished_run_survives_restart(self):
        job = self.finished_job("auto:balance_auto+delmir:10d", JobStatus.completed)
        day_log.record_finished_run(self.root, job)
        day_log.record_finished_run(self.root, job)  # повторный статус — та же запись
        runs = day_log.load_runs(self.root, date.today())
        self.assertEqual(len(runs), 1)
        self.assertEqual(runs[0]["kind"], "balance_auto+delmir")
        self.assertTrue(runs[0]["auto"])
        self.assertEqual(runs[0]["status"], "completed")

    def test_failed_run_keeps_short_reason(self):
        job = self.finished_job(
            "sklad_template_copy",
            JobStatus.failed,
            "RuntimeError: Нет папки склада. Подробности в журнале: C:\\log.log",
        )
        day_log.record_finished_run(self.root, job)
        self.assertEqual(day_log.load_runs(self.root, date.today())[0]["error"], "Нет папки склада.")

    def test_marks_set_and_clear(self):
        marks = day_log.set_mark(self.root, "grid-fill", day_log.STEP_DONE)
        self.assertEqual(marks["grid-fill"]["state"], "done")
        day_log.set_mark(self.root, "holland-full", day_log.STEP_UNDONE)
        marks = day_log.set_mark(self.root, "grid-fill", None)
        self.assertEqual(set(marks), {"holland-full"})
        self.assertEqual(day_log.load_marks(self.root, date.today()), marks)

    def test_rejects_unknown_step_and_state(self):
        with self.assertRaises(ValueError):
            day_log.set_mark(self.root, "../etc", day_log.STEP_DONE)
        with self.assertRaises(ValueError):
            day_log.set_mark(self.root, "grid-fill", "maybe")

    def test_old_day_files_are_pruned(self):
        folder = self.root / "data" / "state" / "steps"
        folder.mkdir(parents=True)
        old = folder / f"{(date.today() - timedelta(days=30)).isoformat()}.json"
        old.write_text(json.dumps({}), encoding="utf-8")
        day_log.set_mark(self.root, "grid-fill", day_log.STEP_DONE)
        self.assertFalse(old.exists())

    def test_run_started_time_is_local(self):
        job = self.finished_job("mix_separation", JobStatus.completed)
        started = datetime.fromisoformat(day_log.run_entry(job)["started"])
        self.assertLess(abs((started - datetime.now()).total_seconds()), 5)


if __name__ == "__main__":
    unittest.main()
