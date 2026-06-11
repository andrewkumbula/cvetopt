from __future__ import annotations

import asyncio

from cvetopt.core.job_manager import job_log
from cvetopt.core.runtime_settings import (
    effective_auto_new_workbook_raw,
    effective_holland_sklad_dir_raw,
    load_runtime_settings,
    resolve_auto_new_workbook,
    resolve_biflorica_archive_dir,
    resolve_ecuador_template,
    resolve_holland_sklad_dir,
)
from cvetopt.core.settings import EnvSettings
from cvetopt.invoice.auto1_pipeline import run_auto1_pipeline
from cvetopt.invoice.holland_translate import (
    archive_stale_holland_exports,
    find_holland_export_file,
)


async def run_auto1_pipeline_job(job_id: str, env: EnvSettings) -> None:
    yaml_cfg = env.yaml_config()
    cfg = yaml_cfg.auto1_pipeline

    if not cfg.enabled:
        await job_log(job_id, "auto1_pipeline отключён в config.yaml")
        return

    runtime = load_runtime_settings(env)
    wb_raw = effective_auto_new_workbook_raw(
        runtime,
        yaml_auto1=cfg.workbook_path,
        yaml_balance=yaml_cfg.balance_auto.workbook_path,
    )
    wb_path = resolve_auto_new_workbook(env, wb_raw)
    await job_log(job_id, f"Книга Auto_new: {wb_path}")

    async def lg(msg: str) -> None:
        await job_log(job_id, msg)

    await lg(
        "Auto1: Scan → Import → Calculate → Sort → for sklad "
        f"(лист {cfg.sheet_name!r})"
    )

    loop = asyncio.get_running_loop()

    sklad_dir = resolve_holland_sklad_dir(
        env,
        effective_holland_sklad_dir_raw(
            runtime,
            yaml_dir=yaml_cfg.holland_translate.sklad_output_dir,
        ),
    )

    def _thread_log(msg: str) -> None:
        fut = asyncio.run_coroutine_threadsafe(job_log(job_id, msg), loop)
        try:
            fut.result(timeout=30)
        except Exception:
            pass

    try:
        await asyncio.to_thread(
            run_auto1_pipeline,
            wb_path,
            cfg,
            sklad_export_dir=sklad_dir,
            log=_thread_log,
        )
    except Exception as e:
        await job_log(job_id, f"Ошибка auto1: {e}")
        raise

    holland_cfg = yaml_cfg.holland_translate
    newest = find_holland_export_file(sklad_dir)

    if holland_cfg.archive_previous_on_auto1:
        archive_dir = resolve_biflorica_archive_dir(
            env,
            runtime.biflorica_archive_dir,
            runtime=runtime,
        )
        archive_dir.mkdir(parents=True, exist_ok=True)
        if newest is None:
            await lg(f"Голландия: файл не найден в {sklad_dir} — архив не нужен")
        else:
            await lg(f"Голландия: оставляем {newest.name}")
            try:
                await asyncio.to_thread(
                    archive_stale_holland_exports,
                    sklad_dir,
                    archive_dir,
                    keep_path=newest,
                    log=_thread_log,
                )
            except Exception as e:
                await lg(f"Голландия: архив пропущен — {e}")

    if holland_cfg.add_row_markers and newest is not None:
        from cvetopt.invoice.holland_markers import add_holland_row_markers

        assets_dir = resolve_ecuador_template(env, runtime.ecuador_template_path).parent
        try:
            marked = await asyncio.to_thread(
                add_holland_row_markers,
                newest,
                assets_dir,
                log=_thread_log,
            )
            await lg(f"Голландия: маркеры → {marked.name}")
        except Exception as e:
            await lg(f"Голландия: маркеры пропущены — {e}")

    await lg("Готово: цепочка auto1 выполнена.")
