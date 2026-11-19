"""Архив Эквадора после «Создать файл»: старые файлы уходят, только что созданный остаётся."""

import os
import tempfile
import unittest
from pathlib import Path

from cvetopt.invoice.ecuador_create import archive_stale_ecuador_exports


class EcuadorArchiveTests(unittest.TestCase):
    def setUp(self):
        temp = tempfile.TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        self.root = Path(temp.name)
        self.out = self.root / "Инвойсы склад"
        self.archive = self.out / "Архивы Инвойсы"
        self.out.mkdir()
        self.old = self.out / "Эквадор 12.11.26 09.01.xlsm"
        self.new = self.out / "Эквадор 19.11.26 09.01.xlsm"
        self.old.write_bytes(b"old")
        self.new.write_bytes(b"new")
        os.utime(self.old, (1, 1))

    def run_archive(self, keep):
        lines = []
        archive_stale_ecuador_exports(self.out, self.archive, keep_path=keep, log=lines.append)
        return lines

    def test_new_file_stays_when_folder_is_seen_by_another_path(self):
        # Как 19.11 на сервере: путь к новому файлу не совпал с путём папки выгрузки.
        other_path = self.root / "D" / "Инвойсы Склад" / self.new.name
        lines = self.run_archive(other_path)
        self.assertTrue(self.new.is_file())
        self.assertFalse(self.old.exists())
        self.assertTrue((self.archive / self.old.name).is_file())
        self.assertIn(f"Эквадор: в архив 1 → {self.archive.resolve()}: {self.old.name}", lines)

    def test_new_file_stays_by_same_path(self):
        self.run_archive(self.new)
        self.assertTrue(self.new.is_file())
        self.assertEqual([p.name for p in self.archive.iterdir()], [self.old.name])


if __name__ == "__main__":
    unittest.main()
