from __future__ import annotations

import asyncio

from cvetopt.core.job_manager import job_manager, job_step
from cvetopt.core.settings import EnvSettings
from cvetopt.selfcheck.cases import CHECK_FAIL, CHECK_OK, CheckResult
from cvetopt.selfcheck.runner import run_selfcheck, selftest_root, summarize

_ICONS = {CHECK_OK: "✓", CHECK_FAIL: "✗"}


async def run_selfcheck_job(job_id: str, env: EnvSettings) -> None:
    """«Проверить программу»: окружение (только чтение) + логика на тестовых файлах."""
    await job_step(job_id, f"Проверка: тестовые файлы — в {selftest_root(env) / 'run'}")

    loop = asyncio.get_running_loop()

    async def _report(result: CheckResult) -> None:
        await job_manager.add_check(job_id, result.as_dict())
        line = f"{_ICONS.get(result.status, '–')} {result.group}: {result.name}"
        if result.detail:
            line += f" — {result.detail}"
        await job_manager.append_ui_log(job_id, line)

    def _on_result(result: CheckResult) -> None:
        asyncio.run_coroutine_threadsafe(_report(result), loop).result(timeout=60)

    results = await asyncio.to_thread(run_selfcheck, env, on_result=_on_result)
    summary = summarize(results)
    await job_step(
        job_id,
        f"Проверка: прошло {summary['ok']} из {summary['total']}, "
        f"не прошло {summary['failed']}",
    )
    if summary["failed"]:
        raise RuntimeError(
            f"Проверка: не прошло {summary['failed']} из {summary['total']} — "
            "смотрите красные пункты в списке"
        )
