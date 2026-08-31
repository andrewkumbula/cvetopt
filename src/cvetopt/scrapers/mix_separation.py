from __future__ import annotations

import asyncio
from pathlib import Path

from cvetopt.core.job_manager import job_log
from cvetopt.core.runtime_settings import (
    effective_holland_sklad_dir_raw,
    load_runtime_settings,
    resolve_biflorica_download_dir,
    resolve_holland_sklad_dir,
)
from cvetopt.core.settings import EnvSettings
from cvetopt.invoice.biflorica_mixes import run_mix_separation_from_dirs


async def run_mix_separation_job(
    job_id: str,
    env: EnvSettings,
    *,
    template_path: str | None = None,
    biflorica_path: str | None = None,
) -> None:
    """
    Заполненный «Шаблон ДД.ММ.ГГ» + Biflorica Mix (50/60) →
    отдельные позиции со средневзвешенной ценой без плантации.
    """
    runtime = load_runtime_settings(env)
    sklad_dir = resolve_holland_sklad_dir(
        env,
        effective_holland_sklad_dir_raw(env, runtime),
    )
    download_dir = resolve_biflorica_download_dir(env, runtime.biflorica_download_dir)
    await job_log(job_id, f"Миксы: папка склада {sklad_dir}")
    await job_log(job_id, f"Миксы: папка Biflorica {download_dir}")

    tpl: Path | None = Path(template_path.strip()) if template_path and template_path.strip() else None
    bif: Path | None = (
        Path(biflorica_path.strip()) if biflorica_path and biflorica_path.strip() else None
    )

    loop = asyncio.get_running_loop()

    def _thread_log(msg: str) -> None:
        fut = asyncio.run_coroutine_threadsafe(job_log(job_id, msg), loop)
        try:
            fut.result(timeout=60)
        except Exception:
            pass

    try:
        tpl_out, bif_out, plans = await asyncio.to_thread(
            run_mix_separation_from_dirs,
            sklad_dir,
            download_dir,
            template_path=tpl,
            biflorica_path=bif,
            log=_thread_log,
        )
    except FileNotFoundError as e:
        await job_log(job_id, f"Миксы: {e}")
        raise
    except RuntimeError as e:
        await job_log(job_id, f"Миксы: {e}")
        raise

    summary = ", ".join(
        f"{p.length}см → {len(p.lines)} поз. @ {p.avg_price:.4f}" for p in plans
    )
    await job_log(
        job_id,
        f"Миксы: готово — шаблон {tpl_out.name}, Biflorica {bif_out.name}; {summary}",
    )
    await job_log(job_id, "Готово.")
