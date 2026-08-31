from __future__ import annotations

import asyncio
from pathlib import Path

from cvetopt.core.job_manager import job_log
from cvetopt.core.runtime_settings import (
    load_runtime_settings,
    resolve_biflorica_download_dir,
)
from cvetopt.core.settings import EnvSettings
from cvetopt.invoice.biflorica_split import run_gypsophila_split


async def run_gypsophila_split_job(
    job_id: str,
    env: EnvSettings,
    *,
    source_path: str | None = None,
) -> None:
    """
    Аналог «открыть Гипсофила → выбрать Biflorica → Да»:
    делит свежий BiFlorica-*.xlsx на «… Гипсофила.xlsx» и «… Роза.xlsx».
    """
    runtime = load_runtime_settings(env)
    download_dir = resolve_biflorica_download_dir(env, runtime.biflorica_download_dir)
    await job_log(job_id, f"Гипсофила: папка Biflorica {download_dir}")

    source: Path | None = None
    if source_path and source_path.strip():
        source = Path(source_path.strip())

    loop = asyncio.get_running_loop()

    def _thread_log(msg: str) -> None:
        fut = asyncio.run_coroutine_threadsafe(job_log(job_id, msg), loop)
        try:
            fut.result(timeout=60)
        except Exception:
            pass

    try:
        result = await asyncio.to_thread(
            run_gypsophila_split,
            download_dir,
            source=source,
            log=_thread_log,
        )
    except FileNotFoundError as e:
        await job_log(job_id, f"Гипсофила: {e}")
        raise

    await job_log(
        job_id,
        "Гипсофила: готово — "
        + ", ".join(f"{label} {n} стр." for label, n in result.counts.items())
        + f"; полный файл: {result.source.name}",
    )
    await job_log(job_id, "Готово.")
