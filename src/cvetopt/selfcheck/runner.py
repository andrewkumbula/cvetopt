from __future__ import annotations

import json
import shutil
from collections.abc import Callable
from datetime import datetime
from pathlib import Path

from cvetopt.core.runtime_settings import load_runtime_settings
from cvetopt.core.settings import EnvSettings
from cvetopt.selfcheck import logic_cases, mix_cases
from cvetopt.selfcheck.cases import CHECK_FAIL, CHECK_OK, CheckResult, run_cases
from cvetopt.selfcheck.environment import Context, run_environment_checks


def selftest_root(env: EnvSettings) -> Path:
    """Своя папка проверки: data/selftest. Вне её проверка ничего не пишет."""
    return env.project_root / "data" / "selftest"


def run_selfcheck(
    env: EnvSettings,
    *,
    on_result: Callable[[CheckResult], None],
) -> list[CheckResult]:
    root = selftest_root(env)
    work_dir = root / "run"
    # Файлы прошлой проверки удаляем, текущие оставляем — их можно открыть и посмотреть.
    shutil.rmtree(work_dir, ignore_errors=True)
    work_dir.mkdir(parents=True)

    results: list[CheckResult] = []

    def _collect(result: CheckResult) -> None:
        results.append(result)
        on_result(result)

    ctx = Context(env=env, runtime=load_runtime_settings(env), yaml=env.yaml_config())
    run_environment_checks(ctx, on_result=_collect)
    run_cases("Логика программы", logic_cases.TESTS, work_dir / "logic", on_result=_collect)
    run_cases("Миксы (подробно)", mix_cases.TESTS, work_dir / "mixes", on_result=_collect)

    save_last_result(root, results)
    return results


def summarize(results: list[CheckResult]) -> dict[str, int]:
    return {
        "total": len(results),
        "ok": sum(r.status == CHECK_OK for r in results),
        "failed": sum(r.status == CHECK_FAIL for r in results),
    }


def save_last_result(root: Path, results: list[CheckResult]) -> None:
    payload = {
        "finished_at": datetime.now().isoformat(timespec="seconds"),
        **summarize(results),
        "checks": [r.as_dict() for r in results],
    }
    (root / "last.json").write_text(
        json.dumps(payload, ensure_ascii=False, indent=1), encoding="utf-8"
    )


def load_last_result(env: EnvSettings) -> dict | None:
    path = selftest_root(env) / "last.json"
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return None
