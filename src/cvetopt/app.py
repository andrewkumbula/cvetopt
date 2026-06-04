from __future__ import annotations

import asyncio
import os
import subprocess
import sys
from pathlib import Path

import uvicorn
from fastapi import FastAPI, Request
from fastapi.responses import HTMLResponse, JSONResponse, RedirectResponse
from fastapi.templating import Jinja2Templates

from cvetopt.core.job_manager import job_manager, run_coro_logged
from cvetopt.core.models import JobStatus
from cvetopt.core.logging_setup import configure_logging
from cvetopt.core.runtime_settings import (
    RuntimeSettings,
    load_runtime_settings,
    save_runtime_settings,
    validate_biflorica_archive_dir,
    validate_biflorica_download_dir,
    validate_auto_new_workbook_path,
    validate_ecuador_paths,
    validate_mail_output_dirs,
)
from cvetopt.core.settings import EnvSettings, SelectionOverride
from cvetopt.core.testing_reset import reset_testing_state
from cvetopt.scrapers.auto1_pipeline import run_auto1_pipeline_job
from cvetopt.scrapers.balance_auto import run_balance_auto_job
from cvetopt.scrapers.biflorica import run_biflorica_job
from cvetopt.scrapers.delmir import run_delmir_transport_job
from cvetopt.scrapers.mail_attachments import run_mail_attachments_job

PROJECT_ROOT = Path(__file__).resolve().parents[2]
RESTART_EXIT_CODE = 42  # лаунчер cvetopt.bat подхватывает этот код и делает git pull + перезапуск.

_TEMPLATES = Jinja2Templates(
    directory=str(Path(__file__).resolve().parent / "ui" / "templates"),
)

app = FastAPI(title="cvetopt", version="0.1.0")
configure_logging(EnvSettings())


def _git_version() -> dict[str, str]:
    """Возвращает (commit, date, branch) текущего worktree, или пустые поля если git недоступен."""
    info: dict[str, str] = {"commit": "", "date": "", "branch": ""}
    if not (PROJECT_ROOT / ".git").exists():
        return info
    try:
        info["commit"] = subprocess.check_output(
            ["git", "rev-parse", "--short", "HEAD"], cwd=PROJECT_ROOT, text=True
        ).strip()
    except Exception:
        pass
    try:
        info["date"] = subprocess.check_output(
            ["git", "log", "-1", "--format=%cd", "--date=short"], cwd=PROJECT_ROOT, text=True
        ).strip()
    except Exception:
        pass
    try:
        info["branch"] = subprocess.check_output(
            ["git", "rev-parse", "--abbrev-ref", "HEAD"], cwd=PROJECT_ROOT, text=True
        ).strip()
    except Exception:
        pass
    return info


@app.get("/", response_class=HTMLResponse)
async def index(request: Request) -> HTMLResponse:
    env = EnvSettings()
    runtime_settings = load_runtime_settings(env)
    return _TEMPLATES.TemplateResponse(
        request,
        "index.html",
        {
            "version": _git_version(),
            "git_available": (PROJECT_ROOT / ".git").exists(),
            "runtime_settings": runtime_settings.model_dump(),
        },
    )


@app.get("/api/version")
async def api_version() -> JSONResponse:
    return JSONResponse({**_git_version(), "active_job": job_manager.has_active_job()})


@app.get("/api/state")
async def api_state() -> JSONResponse:
    return JSONResponse(
        {
            "active_job": job_manager.has_active_job(),
            "recent": [
                {"id": j.id, "portal_id": j.portal_id, "status": j.status.value}
                for j in job_manager.list_recent(5)
            ],
        }
    )


@app.get("/api/runtime-settings")
async def api_runtime_settings() -> JSONResponse:
    env = EnvSettings()
    return JSONResponse(load_runtime_settings(env).model_dump())


@app.post("/api/runtime-settings")
async def api_runtime_settings_update(request: Request) -> JSONResponse:
    env = EnvSettings()
    try:
        data = await request.json()
        current = load_runtime_settings(env).model_dump()
        if not isinstance(data, dict):
            return JSONResponse({"error": "Ожидался JSON-объект."}, status_code=422)
        merged = {**current, **data}
        try:
            settings = RuntimeSettings.model_validate(merged)
        except Exception as e:
            return JSONResponse({"error": f"Некорректные значения: {e}"}, status_code=422)

        if settings.biflorica_min_age_days < 0 or settings.biflorica_max_age_days < 0:
            return JSONResponse({"error": "Biflorica: период не может быть отрицательным."}, status_code=422)
        if settings.biflorica_min_age_days > settings.biflorica_max_age_days:
            return JSONResponse({"error": "Biflorica: min не может быть больше max."}, status_code=422)
        if settings.biflorica_max_age_days > 365:
            return JSONResponse({"error": "Biflorica: максимум 365 дней."}, status_code=422)
        if settings.delmir_lookback_days < 1 or settings.delmir_lookback_days > 365:
            return JSONResponse({"error": "del-mir: диапазон 1..365 дней."}, status_code=422)
        if settings.mail_lookback_days < 1 or settings.mail_lookback_days > 365:
            return JSONResponse({"error": "Почта: диапазон 1..365 дней."}, status_code=422)
        dir_err = validate_biflorica_download_dir(env, settings.biflorica_download_dir)
        if dir_err:
            return JSONResponse({"error": dir_err}, status_code=422)
        arch_err = validate_biflorica_archive_dir(
            env,
            settings.biflorica_archive_dir,
            settings.biflorica_download_dir,
        )
        if arch_err:
            return JSONResponse({"error": arch_err}, status_code=422)
        auto_err = validate_auto_new_workbook_path(env, settings.auto_new_workbook_path)
        if auto_err:
            return JSONResponse({"error": auto_err}, status_code=422)
        ecu_err = validate_ecuador_paths(
            env,
            settings.ecuador_template_path,
            settings.ecuador_output_dir,
        )
        if ecu_err:
            return JSONResponse({"error": ecu_err}, status_code=422)
        mail_err = validate_mail_output_dirs(env, settings)
        if mail_err:
            return JSONResponse({"error": mail_err}, status_code=422)

        save_runtime_settings(env, settings)
        return JSONResponse({"ok": True, "settings": settings.model_dump()})
    except Exception as e:
        return JSONResponse(
            {"error": f"Ошибка сохранения настроек: {e}"},
            status_code=500,
        )


def _pick_folder_native() -> str | None:
    try:
        import tkinter as tk
        from tkinter import filedialog
    except ImportError:
        return None
    root = tk.Tk()
    root.withdraw()
    try:
        root.attributes("-topmost", True)
    except tk.TclError:
        pass
    try:
        return filedialog.askdirectory() or None
    finally:
        root.destroy()


@app.post("/api/pick-folder")
async def api_pick_folder() -> JSONResponse:
    """Системный диалог выбора папки (tkinter), только для локального UI."""
    path = await asyncio.to_thread(_pick_folder_native)
    if path is None:
        return JSONResponse(
            {"error": "Диалог недоступен или папка не выбрана. Введите путь вручную."},
            status_code=503,
        )
    return JSONResponse({"path": path})


@app.post("/api/testing/reset")
async def api_testing_reset(request: Request) -> JSONResponse:
    """Сброс реестров (и опционально файлов) для повторного тестового скачивания."""
    busy = _reject_if_busy()
    if busy is not None:
        return busy
    try:
        data = await request.json()
    except Exception:
        data = {}
    if not isinstance(data, dict):
        data = {}

    lines = reset_testing_state(
        EnvSettings(),
        biflorica_registry=bool(data.get("biflorica_registry", True)),
        mail_registry=bool(data.get("mail_registry", True)),
        biflorica_files=bool(data.get("biflorica_files", False)),
        mail_files=bool(data.get("mail_files", False)),
    )
    return JSONResponse({"ok": True, "messages": lines})


def _reject_if_busy() -> JSONResponse | None:
    if job_manager.has_active_job():
        return JSONResponse(
            {"error": "Уже выполняется другой прогон. Дождитесь его завершения."},
            status_code=409,
        )
    return None


@app.post("/run/balance-auto")
async def run_balance_auto(
    request: Request,
    delmir_lookback_days: int | None = None,
):
    """
    Объединённый шаг 2+3: сначала balance_auto (Biflorica → перелёты),
    затем delmir_transport (Транспорт трак). del-mir стартует ТОЛЬКО при успехе balance_auto.
    """
    busy = _reject_if_busy()
    if busy is not None:
        return busy
    env = EnvSettings()
    runtime_settings = load_runtime_settings(env)
    effective_delmir_lookback = (
        delmir_lookback_days
        if delmir_lookback_days is not None
        else runtime_settings.delmir_lookback_days
    )
    if effective_delmir_lookback < 1 or effective_delmir_lookback > 365:
        return JSONResponse(
            {"error": "Период del-mir должен быть в диапазоне 1..365 дней."},
            status_code=422,
        )
    job = job_manager.create_job(f"balance_auto+delmir:{effective_delmir_lookback}d")

    async def _chain() -> None:
        from cvetopt.core.job_manager import job_log
        from cvetopt.core.models import JobStatus

        await run_coro_logged(job.id, run_balance_auto_job(job.id, env))
        current = job_manager.get(job.id)
        if current is None or current.status != JobStatus.completed:
            await job_log(
                job.id,
                "Шаг 2 (баланс Biflorica) завершился неудачно — Транспорт трак с del-mir пропущен.",
            )
            return
        await job_log(
            job.id,
            f"Шаг 2 завершён успешно. Запускаю Транспорт трак с del-mir.com ({effective_delmir_lookback} дн.)…",
        )
        await run_coro_logged(
            job.id,
            run_delmir_transport_job(
                job.id,
                env,
                lookback_days_override=effective_delmir_lookback,
            ),
        )

    job_manager.schedule(job.id, _chain())
    return RedirectResponse(url=f"/job/{job.id}", status_code=303)


@app.post("/run/auto1-pipeline")
async def run_auto1_pipeline_route(request: Request):
    """Лист auto1: Scan → Import → Calculate → Sort → for sklad (VBA через Excel, только Windows)."""
    busy = _reject_if_busy()
    if busy is not None:
        return busy
    env = EnvSettings()
    job = job_manager.create_job("auto1_pipeline")
    job_manager.schedule(job.id, run_auto1_pipeline_job(job.id, env))
    return RedirectResponse(url=f"/job/{job.id}", status_code=303)


@app.post("/run/biflorica")
async def run_biflorica(
    request: Request,
    min_age_days: int | None = None,
    max_age_days: int | None = None,
):
    busy = _reject_if_busy()
    if busy is not None:
        return busy
    env = EnvSettings()
    runtime_settings = load_runtime_settings(env)
    effective_min_age = (
        min_age_days if min_age_days is not None else runtime_settings.biflorica_min_age_days
    )
    effective_max_age = (
        max_age_days if max_age_days is not None else runtime_settings.biflorica_max_age_days
    )
    if effective_min_age < 0 or effective_max_age < 0:
        return JSONResponse(
            {"error": "Период Biflorica не может быть отрицательным."},
            status_code=422,
        )
    if effective_min_age > effective_max_age:
        return JSONResponse(
            {"error": "Для Biflorica min_age_days не может быть больше max_age_days."},
            status_code=422,
        )
    if effective_max_age > 365:
        return JSONResponse(
            {"error": "Период Biflorica должен быть в диапазоне 0..365 дней."},
            status_code=422,
        )
    bif_selection = SelectionOverride(
        min_age_days=effective_min_age,
        max_age_days=effective_max_age,
    )
    job = job_manager.create_job(f"biflorica:{effective_min_age}-{effective_max_age}d")

    job_manager.schedule(
        job.id,
        run_biflorica_job(job.id, env, selection_override=bif_selection),
    )
    return RedirectResponse(url=f"/job/{job.id}", status_code=303)


@app.post("/run/mail-attachments")
async def run_mail_attachments(
    request: Request,
    lookback_days: int | None = None,
):
    busy = _reject_if_busy()
    if busy is not None:
        return busy
    env = EnvSettings()
    runtime_settings = load_runtime_settings(env)
    effective_lookback = (
        lookback_days if lookback_days is not None else runtime_settings.mail_lookback_days
    )
    if effective_lookback < 1 or effective_lookback > 365:
        return JSONResponse(
            {"error": "Период должен быть в диапазоне 1..365 дней."},
            status_code=422,
        )
    job = job_manager.create_job(f"mail_attachments:{effective_lookback}d")

    job_manager.schedule(
        job.id,
        run_mail_attachments_job(
            job.id,
            env,
            lookback_days_override=effective_lookback,
        ),
    )
    return RedirectResponse(url=f"/job/{job.id}", status_code=303)


@app.post("/run/delmir-transport")
async def run_delmir(
    request: Request,
    lookback_days: int = 14,
):
    """Оставлено для отладки — кнопки в UI больше нет, см. /run/balance-auto."""
    busy = _reject_if_busy()
    if busy is not None:
        return busy
    if lookback_days < 1 or lookback_days > 365:
        return JSONResponse(
            {"error": "Период del-mir должен быть в диапазоне 1..365 дней."},
            status_code=422,
        )
    env = EnvSettings()
    job = job_manager.create_job(f"delmir_transport:{lookback_days}d")

    job_manager.schedule(
        job.id,
        run_delmir_transport_job(job.id, env, lookback_days_override=lookback_days),
    )
    return RedirectResponse(url=f"/job/{job.id}", status_code=303)


@app.post("/admin/update")
async def admin_update():
    """
    Сигналит лаунчеру cvetopt.bat выйти с кодом 42 → лаунчер делает git pull + uv sync
    и поднимает сервер снова. На запросе мы успеваем отдать 202, а уже потом грохаемся.
    """
    if job_manager.has_active_job():
        return JSONResponse(
            {"error": "Сейчас выполняется прогон — обновление отложено. Подождите окончания."},
            status_code=409,
        )

    async def _kill_self() -> None:
        await asyncio.sleep(1.0)
        os._exit(RESTART_EXIT_CODE)

    asyncio.create_task(_kill_self())
    return JSONResponse(
        {
            "ok": True,
            "message": "Обновление запущено. Сервер перезапустится через несколько секунд.",
        },
        status_code=202,
    )


@app.get("/job/{job_id}", response_class=HTMLResponse)
async def job_page(request: Request, job_id: str) -> HTMLResponse:
    return _TEMPLATES.TemplateResponse(request, "job.html", {"job_id": job_id})


@app.get("/api/job/{job_id}")
async def job_api(job_id: str) -> JSONResponse:
    job = job_manager.get(job_id)
    if not job:
        return JSONResponse({"error": "not found"}, status_code=404)
    return JSONResponse(
        {
            "id": job.id,
            "portal_id": job.portal_id,
            "status": job.status.value,
            "logs": job.logs,
            "downloaded_paths": job.downloaded_paths,
            "error": job.error,
            "cancellable": job.status in (JobStatus.pending, JobStatus.running),
        }
    )


@app.post("/api/job/{job_id}/cancel")
async def job_cancel(job_id: str) -> JSONResponse:
    ok, message = await job_manager.cancel_job(job_id)
    if not ok:
        status = 404 if message == "Прогон не найден." else 409
        return JSONResponse({"error": message}, status_code=status)
    return JSONResponse({"ok": True, "status": JobStatus.cancelled.value})


def main() -> None:
    env = EnvSettings()
    uvicorn.run(
        "cvetopt.app:app",
        host=env.app_host,
        port=env.app_port,
        reload=False,
    )


if __name__ == "__main__":
    main()
