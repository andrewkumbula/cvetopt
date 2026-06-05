from __future__ import annotations

import asyncio
import queue
import subprocess
import sys

from cvetopt.core.job_manager import job_log
from cvetopt.core.runtime_settings import (
    load_runtime_settings,
    resolve_biflorica_download_dir,
    resolve_ecuador_output_dir,
    resolve_ecuador_template,
)
from cvetopt.core.settings import EnvSettings
from cvetopt.invoice.ecuador_create import (
    create_ecuador_file_from_biflorica,
    find_latest_biflorica_report,
)


async def run_ecuador_create_job(job_id: str, env: EnvSettings) -> None:
    yaml_cfg = env.yaml_config()
    cfg = yaml_cfg.ecuador_create

    if not cfg.enabled:
        await job_log(job_id, "ecuador_create отключён в config.yaml")
        return

    if sys.platform != "win32":
        await job_log(job_id, "Эквадор: нужен Windows + Excel")
        return

    runtime = load_runtime_settings(env)
    download_dir = resolve_biflorica_download_dir(env, runtime.biflorica_download_dir)
    biflorica_path = find_latest_biflorica_report(download_dir)
    if biflorica_path is None:
        raise FileNotFoundError(
            f"В {download_dir} нет отчёта BiFlorica-*.xlsx. "
            "Сначала скачайте Biflorica или положите xlsx в папку скачивания."
        )

    template = resolve_ecuador_template(env, runtime.ecuador_template_path)
    out_dir = resolve_ecuador_output_dir(env, runtime.ecuador_output_dir)

    await job_log(job_id, f"Эквадор: отчёт {biflorica_path.name}")
    await job_log(job_id, f"Эквадор: шаблон {template}")
    await job_log(job_id, f"Эквадор: выгрузка {out_dir}")

    ecuador_log_q: queue.Queue[str] = queue.Queue()

    def _ecuador_log(msg: str) -> None:
        ecuador_log_q.put(msg)

    async def _drain_logs() -> None:
        while True:
            try:
                while True:
                    await job_log(job_id, ecuador_log_q.get_nowait())
            except queue.Empty:
                pass
            await asyncio.sleep(0.2)

    drain_task = asyncio.create_task(_drain_logs())
    timeout_sec = 600
    try:
        out = await asyncio.wait_for(
            asyncio.to_thread(
                create_ecuador_file_from_biflorica,
                biflorica_path,
                env,
                log=_ecuador_log,
            ),
            timeout=timeout_sec,
        )
    except TimeoutError:
        subprocess.run(
            ["taskkill", "/im", "EXCEL.EXE", "/f"],
            capture_output=True,
            check=False,
        )
        raise RuntimeError(
            f"Эквадор: таймаут {timeout_sec} с — Excel, вероятно, ждёт диалог. "
            "Задайте ECUADOR_EXCEL_VISIBLE=1 в .env и перезапустите cvetopt.bat."
        ) from None
    finally:
        drain_task.cancel()
        try:
            await drain_task
        except asyncio.CancelledError:
            pass
        try:
            while True:
                await job_log(job_id, ecuador_log_q.get_nowait())
        except queue.Empty:
            pass

    await job_log(job_id, f"Готово: {out}")
