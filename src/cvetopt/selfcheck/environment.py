"""
Проверки окружения сервера. Только чтение: ничего не создают, не меняют и не
запускают Excel. Почта открывается в режиме «только просмотр» (письма не
помечаются прочитанными).
"""

from __future__ import annotations

import imaplib
import os
import shutil
import socket
import sys
import urllib.error
import urllib.request
from collections.abc import Callable
from dataclasses import dataclass
from datetime import date, timedelta
from pathlib import Path
from urllib.parse import urlparse

import xlrd

from cvetopt.core.playwright_proxy import playwright_proxy_dict
from cvetopt.core.runtime_settings import (
    RuntimeSettings,
    effective_auto_new_workbook_raw,
    effective_holland_dictionary_raw,
    effective_holland_sklad_dir_raw,
    resolve_auto_new_workbook,
    resolve_biflorica_download_dir,
    resolve_holland_dictionary,
    resolve_holland_sklad_dir,
    resolve_mail_output_layout,
)
from cvetopt.core.settings import AppYamlConfig, EnvSettings
from cvetopt.invoice.description_dictionary import load_description_dictionary
from cvetopt.mail.attachments import _imap_search_criteria
from cvetopt.selfcheck.cases import CHECK_FAIL, CHECK_OK, CHECK_SKIP, CheckResult, SkipCheck

GROUP = "Окружение сервера"

_NETWORK_TIMEOUT_SEC = 15
_MIN_FREE_DISK_GB = 1.0


@dataclass
class Context:
    env: EnvSettings
    runtime: RuntimeSettings
    yaml: AppYamlConfig


CHECKS: list[tuple[str, Callable[[Context], str]]] = []


def check(name: str) -> Callable[[Callable[[Context], str]], Callable[[Context], str]]:
    def deco(fn: Callable[[Context], str]) -> Callable[[Context], str]:
        CHECKS.append((name, fn))
        return fn

    return deco


def run_environment_checks(
    ctx: Context, *, on_result: Callable[[CheckResult], None]
) -> None:
    for name, fn in CHECKS:
        try:
            detail = fn(ctx)
        except SkipCheck as e:
            result = CheckResult(GROUP, name, CHECK_SKIP, str(e))
        except Exception as e:  # noqa: BLE001 — любая ошибка = «не прошла», с текстом
            result = CheckResult(GROUP, name, CHECK_FAIL, str(e) or type(e).__name__)
        else:
            result = CheckResult(GROUP, name, CHECK_OK, detail)
        on_result(result)


# --- Пути из Настроек ---------------------------------------------------------------


def _sklad_dir(ctx: Context) -> Path:
    raw = effective_holland_sklad_dir_raw(
        ctx.runtime, yaml_dir=ctx.yaml.holland_translate.sklad_output_dir
    )
    return resolve_holland_sklad_dir(ctx.env, raw)


def _dictionary_path(ctx: Context) -> Path:
    raw = effective_holland_dictionary_raw(
        ctx.runtime, yaml_path=ctx.yaml.holland_translate.dictionary_path
    )
    return resolve_holland_dictionary(ctx.env, raw)


def _auto_new_path(ctx: Context) -> Path:
    raw = effective_auto_new_workbook_raw(
        ctx.runtime,
        yaml_auto1=ctx.yaml.auto1_pipeline.workbook_path,
        yaml_balance=ctx.yaml.balance_auto.workbook_path,
    )
    return resolve_auto_new_workbook(ctx.env, raw)


def _require_dir(path: Path) -> str:
    if not path.is_dir():
        raise RuntimeError(f"папки {path} нет — проверьте путь в Настройках")
    if not os.access(path, os.R_OK | os.W_OK):
        raise RuntimeError(f"{path}: нет прав на чтение/запись для пользователя программы")
    return str(path)


@check("Папка 1 почты (инвойсы)")
def _(ctx: Context) -> str:
    return _require_dir(resolve_mail_output_layout(ctx.env, ctx.runtime, ctx.yaml.mail).short_dir)


@check("Папка 2 почты (файлы цен)")
def _(ctx: Context) -> str:
    return _require_dir(resolve_mail_output_layout(ctx.env, ctx.runtime, ctx.yaml.mail).long_dir)


@check("Папка склада (Голландия_1_*, шаблоны)")
def _(ctx: Context) -> str:
    return _require_dir(_sklad_dir(ctx))


@check("Папка Biflorica")
def _(ctx: Context) -> str:
    return _require_dir(resolve_biflorica_download_dir(ctx.env, ctx.runtime.biflorica_download_dir))


@check("Свободное место на диске")
def _(ctx: Context) -> str:
    free_gb = shutil.disk_usage(ctx.env.project_root).free / 1024**3
    if free_gb < _MIN_FREE_DISK_GB:
        raise RuntimeError(f"свободно всего {free_gb:.1f} ГБ — освободите место на диске")
    return f"свободно {free_gb:.0f} ГБ"


# --- Файлы ---------------------------------------------------------------------------


def _check_excel_signature(path: Path) -> None:
    head = path.read_bytes()[:8]
    if not any(head):
        raise RuntimeError(f"{path.name} испорчен: файл состоит из нулей — восстановите из копии")
    if not (head.startswith(b"\xd0\xcf\x11\xe0") or head.startswith(b"PK")):
        raise RuntimeError(f"{path.name} — не файл Excel (неизвестное начало файла)")


@check("Словарь.xls открывается и не пустой")
def _(ctx: Context) -> str:
    path = _dictionary_path(ctx)
    if not path.is_file():
        raise RuntimeError(f"словаря нет: {path} — проверьте путь в Настройках")
    _check_excel_signature(path)
    dictionary = load_description_dictionary(path)
    if not dictionary:
        raise RuntimeError(f"в {path.name} нет ни одного перевода")
    return f"{path} — переводов: {len(dictionary)}"


@check("Auto_new.xls на месте (лист auto1)")
def _(ctx: Context) -> str:
    path = _auto_new_path(ctx)
    if not path.is_file():
        raise RuntimeError(f"нет {path} — проверьте путь в Настройках")
    _check_excel_signature(path)
    book = xlrd.open_workbook(str(path), on_demand=True)
    try:
        sheet = ctx.yaml.auto1_pipeline.sheet_name
        if sheet not in book.sheet_names():
            raise RuntimeError(f"в {path.name} нет листа «{sheet}»")
    finally:
        book.release_resources()
    return f"{path}"


@check("Excel установлен")
def _(ctx: Context) -> str:
    if sys.platform != "win32":
        raise SkipCheck("не Windows")
    import winreg

    try:
        winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"Excel.Application\CLSID").Close()
    except OSError:
        raise RuntimeError("Excel не найден в системе — без него Auto1 и Эквадор не работают") from None
    return "Excel.Application зарегистрирован"


# --- Сеть ----------------------------------------------------------------------------


@check("Почта: вход и письма Enigma")
def _(ctx: Context) -> str:
    cfg = ctx.yaml.mail
    if not cfg.enabled:
        raise SkipCheck("почта отключена в config.yaml")
    if not ctx.env.mail_email or not ctx.env.mail_password:
        raise SkipCheck("MAIL_EMAIL / MAIL_PASSWORD не заданы в .env")
    host = ctx.env.mail_imap_host or cfg.imap_host
    imap_cls = imaplib.IMAP4_SSL if cfg.use_ssl else imaplib.IMAP4
    try:
        imap = imap_cls(host, cfg.imap_port, timeout=_NETWORK_TIMEOUT_SEC)
    except OSError as e:
        raise RuntimeError(f"нет связи с {host}: {e} — проверьте интернет/DNS") from None
    try:
        try:
            imap.login(ctx.env.mail_email, ctx.env.mail_password)
        except imaplib.IMAP4.error:
            raise RuntimeError(
                f"{ctx.env.mail_email}: Яндекс не принял пароль — нужен пароль приложения"
            ) from None
        status, _ = imap.select(cfg.folder, readonly=True)
        if status != "OK":
            raise RuntimeError(f"не открылась папка {cfg.folder!r}")
        since = (date.today() - timedelta(days=cfg.lookback_days)).strftime("%d-%b-%Y")
        status, data = imap.search(None, _imap_search_criteria(cfg, since))
        found = len(data[0].split()) if status == "OK" and data and data[0] else 0
    finally:
        try:
            imap.logout()
        except Exception:  # noqa: BLE001
            pass
    return f"{ctx.env.mail_email}: писем за {cfg.lookback_days} дн.: {found}"


def _proxy_server(ctx: Context) -> str:
    proxy = playwright_proxy_dict(ctx.env)
    if proxy is None:
        raise SkipCheck("PLAYWRIGHT_PROXY не задан в .env")
    return proxy["server"]


@check("Прокси Happ отвечает")
def _(ctx: Context) -> str:
    server = _proxy_server(ctx)
    parsed = urlparse(server)
    try:
        socket.create_connection((parsed.hostname, parsed.port or 80), timeout=5).close()
    except OSError:
        raise RuntimeError(f"{server} не отвечает — включите Happ на сервере") from None
    return server


@check("Сайт Biflorica открывается через прокси")
def _(ctx: Context) -> str:
    server = _proxy_server(ctx)
    if not server.startswith("http"):
        raise SkipCheck(f"проверка умеет только HTTP-прокси ({server})")
    url = ctx.yaml.portals.biflorica.login_url
    opener = urllib.request.build_opener(
        urllib.request.ProxyHandler({"http": server, "https": server})
    )
    try:
        with opener.open(url, timeout=_NETWORK_TIMEOUT_SEC) as resp:
            code = resp.status
    except urllib.error.HTTPError as e:
        code = e.code
    except OSError as e:
        raise RuntimeError(f"{urlparse(url).hostname} не открылся через {server}: {e}") from None
    if code >= 500:
        raise RuntimeError(f"сайт отвечает ошибкой {code} — проблема на стороне Biflorica")
    return f"{urlparse(url).hostname}: ответ {code}"
