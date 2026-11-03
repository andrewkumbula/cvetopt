"""Те же кейсы, что запускает кнопка «Проверить программу» (логика и миксы)."""

import tempfile
import unittest
from pathlib import Path

from cvetopt.selfcheck import logic_cases, mix_cases
from cvetopt.selfcheck.cases import CHECK_FAIL, run_cases


class SelfcheckCasesTests(unittest.TestCase):
    def _run(self, cases):
        failed = []
        with tempfile.TemporaryDirectory() as base:
            run_cases(
                "test",
                cases,
                Path(base),
                on_result=lambda r: failed.append(f"{r.name}: {r.detail}")
                if r.status == CHECK_FAIL
                else None,
            )
        self.assertEqual(failed, [])

    def test_logic_cases(self):
        self._run(logic_cases.TESTS)

    def test_mix_cases(self):
        self._run(mix_cases.TESTS)


if __name__ == "__main__":
    unittest.main()
