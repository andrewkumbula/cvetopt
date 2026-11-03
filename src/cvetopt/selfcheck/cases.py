from __future__ import annotations

import shutil
from collections.abc import Callable
from dataclasses import asdict, dataclass
from pathlib import Path

CHECK_OK = "ok"
CHECK_FAIL = "fail"
CHECK_SKIP = "skip"


class SkipCheck(Exception):
    """Проверку нельзя выполнить на этой машине (нет файла, не Windows и т.п.)."""


@dataclass
class CheckResult:
    group: str
    name: str
    status: str
    detail: str = ""

    def as_dict(self) -> dict[str, str]:
        return asdict(self)


CaseFn = Callable[[Path], None]


def run_cases(
    group: str,
    cases: list[tuple[str, CaseFn]],
    work_dir: Path,
    *,
    on_result: Callable[[CheckResult], None],
) -> None:
    """Каждый кейс получает свою пустую папку внутри work_dir."""
    for i, (name, fn) in enumerate(cases):
        tmp = work_dir / f"case_{i:02d}"
        shutil.rmtree(tmp, ignore_errors=True)
        tmp.mkdir(parents=True)
        try:
            fn(tmp)
        except SkipCheck as e:
            result = CheckResult(group, name, CHECK_SKIP, str(e))
        except Exception as e:  # noqa: BLE001 — любая ошибка кейса = «не прошла»
            result = CheckResult(group, name, CHECK_FAIL, f"{type(e).__name__}: {e}")
        else:
            result = CheckResult(group, name, CHECK_OK)
        on_result(result)
