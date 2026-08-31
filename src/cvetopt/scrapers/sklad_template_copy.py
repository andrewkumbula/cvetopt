from __future__ import annotations

import asyncio
from datetime import date

from cvetopt.core.job_manager import job_log
from cvetopt.core.runtime_settings import (
    effective_holland_sklad_dir_raw,
    load_runtime_settings,
    resolve_holland_sklad_dir,
)
from cvetopt.core.settings import EnvSettings
from cvetopt.invoice.sklad_template import copy_sklad_template_to_date


async def run_sklad_template_copy_job(job_id: str, env: EnvSettings) -> None:
    """Копия «шаблон» → «шаблон ДД.ММ.ГГГГ» в папке Инвойсы склад (ручное заполнение сетки дальше)."""
    yaml_cfg = env.yaml_config()
    runtime = load_runtime_settings(env)
    sklad_dir = resolve_holland_sklad_dir(
        env,
        effective_holland_sklad_dir_raw(
            runtime,
            yaml_dir=yaml_cfg.holland_translate.sklad_output_dir,
        ),
    )
    await job_log(job_id, f"Шаблон: папка склада {sklad_dir}")

    loop = asyncio.get_running_loop()

    def _thread_log(msg: str) -> None:
        fut = asyncio.run_coroutine_threadsafe(job_log(job_id, msg), loop)
        try:
            fut.result(timeout=60)
        except Exception:
            pass

    try:
        dest = await asyncio.to_thread(
            copy_sklad_template_to_date,
            sklad_dir,
            on_date=date.today(),
            overwrite=False,
            log=_thread_log,
        )
    except FileExistsError as e:
        await job_log(job_id, f"Шаблон: {e}")
        raise
    except FileNotFoundError as e:
        await job_log(job_id, f"Шаблон: {e}")
        raise

    await job_log(
        job_id,
        f"Шаблон: готово → {dest.name}. "
        "Дальше сотрудник заполняет сетку вручную; оригинал «шаблон» не трогали.",
    )
    await job_log(job_id, "Готово.")
