"""«Перевести инвойсы» дописывает новые слова и в уже выгруженный Голландия_1 склада."""

import asyncio
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest import mock

from cvetopt.scrapers import holland_translated as job

SKLAD = Path("/склад")


class HollandTranslatedJobTests(unittest.TestCase):
    def run_job(self, postprocess):
        lines = []

        async def fake_log(_job_id, msg):
            lines.append(msg)

        folder = SimpleNamespace(path=Path("x.xls"), roses_updated=0)
        env = mock.Mock()
        env.yaml_config.return_value = mock.MagicMock()
        with mock.patch.multiple(
            job,
            job_log=fake_log,
            load_runtime_settings=mock.Mock(),
            resolve_mail_output_layout=mock.Mock(return_value=SimpleNamespace(short_dir="1", long_dir="2")),
            resolve_holland_dictionary=mock.Mock(return_value=Path("Словарь.xls")),
            resolve_auto_new_workbook=mock.Mock(return_value=Path("Auto_new.xls")),
            resolve_holland_sklad_dir=mock.Mock(return_value=SKLAD),
            effective_holland_append_missing=mock.Mock(return_value=True),
            effective_holland_skip_rules=mock.Mock(),
            process_mail_folders=mock.Mock(return_value=(folder, folder)),
            run_auto1_pipeline=mock.Mock(),
            postprocess_holland_after_auto1=postprocess,
        ), mock.patch.object(job.sys, "platform", "darwin"):
            asyncio.run(job.run_holland_translated_job("id", env))
        return lines

    def test_translates_sklad_export(self):
        postprocess = mock.Mock(return_value=SKLAD / "Голландия_1_11.11.2026.xlsx")
        lines = self.run_job(postprocess)
        self.assertEqual(postprocess.call_args.kwargs["sklad_output_dir"], SKLAD)
        self.assertNotIn("add_row_markers", postprocess.call_args.kwargs)
        self.assertEqual(lines[-1], "Готово.")

    def test_missing_export_is_not_an_error(self):
        lines = self.run_job(mock.Mock(return_value=None))
        self.assertIn("Переведено: Голландия_1 в папке склада нет — пропускаю.", lines)
        self.assertEqual(lines[-1], "Готово.")

    def test_open_file_gives_clear_reason(self):
        with self.assertRaisesRegex(RuntimeError, "закройте файл"):
            self.run_job(mock.Mock(side_effect=PermissionError("locked")))


if __name__ == "__main__":
    unittest.main()
