from __future__ import annotations

import asyncio
import sys

from cvetopt.core.job_manager import job_log
from cvetopt.core.runtime_settings import (
    effective_auto_new_workbook_raw,
    effective_holland_append_missing,
    effective_holland_dictionary_raw,
    load_runtime_settings,
    resolve_auto_new_workbook,
    resolve_holland_dictionary,
    resolve_mail_output_layout,
)
from cvetopt.core.settings import EnvSettings
from cvetopt.invoice.auto1_pipeline import run_auto1_pipeline
from cvetopt.invoice.holland_mail_continue import process_mail_folders


async def run_holland_translated_job(job_id: str, env: EnvSettings) -> None:
    """
    «Переведено»: словарь уже заполнен вручную → перевод папок 1 и 2,
    ростовка у роз → Auto1 Scan…Sort (без Group / For sklad).
    """
    yaml_cfg = env.yaml_config()
    translate_cfg = yaml_cfg.holland_translate
    auto1_cfg = yaml_cfg.auto1_pipeline

    if not translate_cfg.enabled:
        await job_log(job_id, "holland_translate отключён в config.yaml")
        return
    if not auto1_cfg.enabled:
        await job_log(job_id, "auto1_pipeline отключён в config.yaml")
        return

    runtime = load_runtime_settings(env)
    layout = resolve_mail_output_layout(env, runtime, yaml_cfg.mail)
    dict_path = resolve_holland_dictionary(
        env,
        effective_holland_dictionary_raw(
            runtime, yaml_path=translate_cfg.dictionary_path
        ),
    )
    wb_path = resolve_auto_new_workbook(
        env,
        effective_auto_new_workbook_raw(
            runtime,
            yaml_auto1=auto1_cfg.workbook_path,
            yaml_balance=yaml_cfg.balance_auto.workbook_path,
        ),
    )

    await job_log(job_id, f"Переведено: папка 1 = {layout.short_dir}")
    await job_log(job_id, f"Переведено: папка 2 = {layout.long_dir}")
    await job_log(job_id, f"Переведено: словарь {dict_path}")

    loop = asyncio.get_running_loop()

    def _thread_log(msg: str) -> None:
        fut = asyncio.run_coroutine_threadsafe(job_log(job_id, msg), loop)
        try:
            fut.result(timeout=120)
        except Exception:
            pass

    await job_log(job_id, "Переведено: перевод Description + ростовка у роз…")
    r1, r2 = await asyncio.to_thread(
        process_mail_folders,
        layout.short_dir,
        layout.long_dir,
        dict_path,
        append_missing_to_dictionary=effective_holland_append_missing(runtime),
        log=_thread_log,
    )
    await job_log(
        job_id,
        f"Переведено: готово {r1.path.name} (роз {r1.roses_updated}) и "
        f"{r2.path.name} (роз {r2.roses_updated})",
    )

    if sys.platform != "win32":
        await job_log(
            job_id,
            "Переведено: Auto1 пропущен (нужен Windows + Excel). "
            "Файлы в папках 1 и 2 уже переведены.",
        )
        await job_log(job_id, "Готово.")
        return

    await job_log(job_id, f"Переведено: Auto1 Scan → … → Sort в {wb_path.name}")
    await asyncio.to_thread(
        run_auto1_pipeline,
        wb_path,
        auto1_cfg,
        stop_after="Sort",
        add_holland_row_markers=False,
        log=_thread_log,
    )
    await job_log(job_id, "Переведено: Auto1 до Sort завершён (Group и For sklad не трогали).")
    await job_log(job_id, "Готово.")
