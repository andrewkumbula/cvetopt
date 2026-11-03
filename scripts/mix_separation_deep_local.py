#!/usr/bin/env python3
"""
Глубокие тесты разбора миксов (сами кейсы — в src/cvetopt/selfcheck/mix_cases.py,
их же запускает кнопка «Проверить программу»). Запуск из корня репо:

  .venv/bin/python scripts/mix_separation_deep_local.py
"""

from __future__ import annotations

import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from cvetopt.selfcheck import mix_cases  # noqa: E402
from cvetopt.selfcheck.cases import CHECK_FAIL, CheckResult, run_cases  # noqa: E402


def main() -> int:
    results: list[CheckResult] = []

    def _print(result: CheckResult) -> None:
        results.append(result)
        print(f"{result.status.upper():4}  {result.name}")
        if result.detail:
            print(f"      {result.detail}")

    with tempfile.TemporaryDirectory(prefix="cvetopt-mix-tests-") as base:
        run_cases("Миксы", mix_cases.TESTS, Path(base), on_result=_print)

    failed = [r.name for r in results if r.status == CHECK_FAIL]
    print()
    print(f"Итого: {len(results) - len(failed)}/{len(results)} прошли")
    return 1 if failed else 0


if __name__ == "__main__":
    raise SystemExit(main())
