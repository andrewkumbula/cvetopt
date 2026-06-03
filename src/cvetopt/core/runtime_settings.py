from __future__ import annotations

import json
import os
import re
import shutil
import stat
import sys
import time
from datetime import date, datetime
from pathlib import Path

from pydantic import BaseModel, Field

from cvetopt.core.settings import EnvSettings, _resolve_selection

DEFAULT_BIFLORICA_DOWNLOAD_DIR = "data/downloads/biflorica"
DEFAULT_BIFLORICA_ARCHIVE_DIR = "data/downloads/biflorica/архив"
DEFAULT_ECUADOR_TEMPLATE = "Invoice/3/Обработка/Прием товара Эквадор-4.xlsm"
ECUADOR_TEMPLATE_FILENAME = "Прием товара Эквадор-4.xlsm"
DEFAULT_ECUADOR_OUTPUT_DIR = r"D:\Склад ОБмен\Инвойсы Склад"
BIFLORICA_ARCHIVE_LEGACY_NAMES = frozenset({"архив", "archive"})
BIFLORICA_DOWNLOAD_PREFIX = "BiFlorica-"
# Скрипт: BiFlorica-<order_id>__<YYYY-MM-DD>.xlsx; старые без префикса тоже в архив.
_BIFLORICA_REPORT_STEM_RE = re.compile(
    r"^(?:BiFlorica-)?\d+__\d{4}-\d{2}-\d{2}$",
    re.IGNORECASE,
)


def biflorica_download_filename(order_id: str, flight_date: date) -> str:
    """Имя xlsx отчёта в папке скачивания (префикс BiFlorica- для отличия от прочих файлов)."""
    return f"{BIFLORICA_DOWNLOAD_PREFIX}{order_id}__{flight_date.isoformat()}.xlsx"


class RuntimeSettings(BaseModel):
    biflorica_min_age_days: int
    biflorica_max_age_days: int
    biflorica_download_dir: str = Field(default=DEFAULT_BIFLORICA_DOWNLOAD_DIR)
    biflorica_archive_dir: str = Field(default=DEFAULT_BIFLORICA_ARCHIVE_DIR)
    ecuador_template_path: str = Field(default=DEFAULT_ECUADOR_TEMPLATE)
    ecuador_output_dir: str = Field(default=DEFAULT_ECUADOR_OUTPUT_DIR)
    delmir_lookback_days: int
    mail_lookback_days: int


def _settings_path(env: EnvSettings) -> Path:
    return env.project_root / "data" / "state" / "runtime_settings.json"


def default_runtime_settings(env: EnvSettings) -> RuntimeSettings:
    yaml_cfg = env.yaml_config()
    bif_sel = _resolve_selection(yaml_cfg.selection, yaml_cfg.portals.biflorica.selection)
    return RuntimeSettings(
        biflorica_min_age_days=bif_sel.min_age_days,
        biflorica_max_age_days=bif_sel.max_age_days,
        biflorica_download_dir=DEFAULT_BIFLORICA_DOWNLOAD_DIR,
        biflorica_archive_dir=DEFAULT_BIFLORICA_ARCHIVE_DIR,
        ecuador_template_path=DEFAULT_ECUADOR_TEMPLATE,
        ecuador_output_dir=DEFAULT_ECUADOR_OUTPUT_DIR,
        delmir_lookback_days=yaml_cfg.delmir.lookback_days,
        mail_lookback_days=yaml_cfg.mail.lookback_days,
    )


def load_runtime_settings(env: EnvSettings) -> RuntimeSettings:
    defaults = default_runtime_settings(env)
    path = _settings_path(env)
    if not path.exists():
        return defaults
    try:
        raw = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return defaults
    if not isinstance(raw, dict):
        return defaults
    return defaults.model_copy(update=raw)


def save_runtime_settings(env: EnvSettings, settings: RuntimeSettings) -> RuntimeSettings:
    path = _settings_path(env)
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(".json.tmp")
    tmp.write_text(settings.model_dump_json(indent=2), encoding="utf-8")
    tmp.replace(path)
    return settings


_WIN_DRIVE_RE = re.compile(r"^[A-Za-z]:[\\/]")


def is_windows_drive_path(text: str) -> bool:
    """Пути вида C:\\... или D:/... (в т.ч. когда на Mac они не absolute для Path)."""
    return bool(_WIN_DRIVE_RE.match((text or "").strip()))


def _resolve_dir(env: EnvSettings, raw_dir: str, default: str) -> Path:
    text = (raw_dir or "").strip() or default
    if is_windows_drive_path(text):
        return Path(text)
    path = Path(text).expanduser()
    if not path.is_absolute():
        path = env.project_root / path
    return path.resolve()


def _skip_local_path_check(raw_dir: str) -> bool:
    """На Mac/Linux пути диска Windows только сохраняем в JSON, не трогаем диск."""
    return sys.platform != "win32" and is_windows_drive_path(raw_dir)


def resolve_biflorica_download_dir(env: EnvSettings, raw_dir: str) -> Path:
    """Абсолютный или относительный (к корню проекта) каталог для xlsx Biflorica."""
    return _resolve_dir(env, raw_dir, DEFAULT_BIFLORICA_DOWNLOAD_DIR)


def resolve_biflorica_archive_dir(
    env: EnvSettings,
    raw_dir: str,
    download_dir: Path | None = None,
) -> Path:
    """Каталог архива Biflorica; пустое значение → <папка скачивания>/архив."""
    text = (raw_dir or "").strip()
    if not text:
        base = download_dir or resolve_biflorica_download_dir(env, "")
        return (base / "архив").resolve()
    return _resolve_dir(env, raw_dir, DEFAULT_BIFLORICA_ARCHIVE_DIR)


def _is_biflorica_report_to_archive(entry: Path) -> bool:
    """Файлы отчётов Biflorica в корне папки (не каталоги Обработка/архив/Задачи)."""
    if not entry.is_file():
        return False
    ext = entry.suffix.lower()
    if ext not in (".xlsx", ".xls"):
        return False
    name_lower = entry.name.lower()
    if "biflorica-deals" in name_lower or name_lower.startswith("biflorica"):
        return True
    if _BIFLORICA_REPORT_STEM_RE.match(entry.stem):
        return True
    return False


def _is_archive_entry(entry: Path, archive_dir: Path, download_dir: Path) -> bool:
    try:
        if entry.resolve() == archive_dir:
            return True
    except OSError:
        pass
    if entry.name in BIFLORICA_ARCHIVE_LEGACY_NAMES:
        try:
            if archive_dir.is_relative_to(download_dir):
                return True
        except ValueError:
            pass
    return False


def _archive_target_path(dest_dir: Path, src: Path, stamp: str) -> Path:
    target = dest_dir / src.name
    if target.exists():
        suffix = src.suffix if src.is_file() else ""
        stem = src.stem if src.is_file() else src.name
        target = dest_dir / f"{stem}_{stamp}{suffix}"
    return target


def _is_access_denied(exc: BaseException) -> bool:
    if isinstance(exc, PermissionError):
        return True
    return getattr(exc, "winerror", None) == 5


def _clear_readonly_windows(path: Path) -> None:
    if sys.platform != "win32":
        return
    if path.is_file():
        os.chmod(path, stat.S_IWRITE | stat.S_IREAD)
        return
    for root, _dirs, files in os.walk(path):
        for name in files:
            os.chmod(os.path.join(root, name), stat.S_IWRITE | stat.S_IREAD)


def _copy_with_retries(src: Path, target: Path, attempts: int = 4) -> None:
    last_err: OSError | None = None
    for attempt in range(attempts):
        try:
            shutil.copy2(src, target)
            return
        except OSError as e:
            last_err = e
            if attempt + 1 < attempts:
                time.sleep(0.35)
    if last_err is not None:
        raise last_err


def _archive_one_entry(src: Path, target: Path) -> str | None:
    """
    Переносит элемент в архив. На Windows — копирование + удаление (надёжнее move).
    При WinError 5 оставляет копию в архиве и пишет предупреждение.
    """
    _clear_readonly_windows(src)

    if src.is_dir():
        shutil.copytree(src, target, dirs_exist_ok=True)
        shutil.rmtree(src, ignore_errors=True)
        if src.exists():
            return f"{src.name}: папка скопирована в архив, исходник частично остался"
        return None

    if sys.platform == "win32":
        _copy_with_retries(src, target)
        try:
            src.unlink()
            return None
        except OSError as e:
            if _is_access_denied(e):
                return (
                    f"{src.name}: копия в архиве есть ({target}), исходник не удалён — "
                    "файл занят другим процессом или нет прав на удаление"
                )
            raise

    try:
        shutil.move(str(src), str(target))
        return None
    except OSError as e:
        if not _is_access_denied(e):
            raise
    _copy_with_retries(src, target)
    try:
        src.unlink()
        return f"{src.name}: скопирован в архив (переместить не удалось)"
    except OSError:
        return f"{src.name}: копия в архиве есть, исходник не удалён"


def archive_biflorica_download_dir(
    download_dir: Path,
    archive_dir: Path,
) -> tuple[Path | None, list[str], list[str]]:
    """
    Переносит в <папка архива>/<YYYY-MM-DD_HHMMSS>/ только xlsx/xls отчётов Biflorica
    из корня папки скачивания (имя начинается с BiFlorica- или старый <order_id>__<дата>.xlsx).
    Подпапки (Обработка, старые архивы, Задачи) и .xlsm/.lnk не трогает.
    """
    if not download_dir.is_dir():
        return None, [], []

    archive_dir = archive_dir.resolve()
    download_dir = download_dir.resolve()
    if archive_dir == download_dir:
        raise ValueError("Папка архива не может совпадать с папкой скачивания.")

    to_move: list[Path] = []
    for entry in download_dir.iterdir():
        if _is_archive_entry(entry, archive_dir, download_dir):
            continue
        if _is_biflorica_report_to_archive(entry):
            to_move.append(entry)

    if not to_move:
        return None, [], []

    stamp = datetime.now().strftime("%Y-%m-%d_%H%M%S")
    dest_dir = archive_dir / stamp
    dest_dir.mkdir(parents=True, exist_ok=True)

    moved: list[str] = []
    warnings: list[str] = []
    access_hint_shown = False
    for src in sorted(to_move, key=lambda p: p.name.lower()):
        target = _archive_target_path(dest_dir, src, stamp)
        if not os.access(src, os.R_OK):
            if not access_hint_shown:
                import getpass

                warnings.append(
                    f"Нет прав на чтение файлов в {download_dir} для пользователя "
                    f"«{getpass.getuser()}» — запустите cvetopt.bat под той же учёткой, "
                    "что владеет C:\\Invoice, или выдайте права на папку."
                )
                access_hint_shown = True
            warnings.append(f"{src.name}: пропуск (нет чтения)")
            continue
        try:
            warn = _archive_one_entry(src, target)
            moved.append(src.name)
            if warn:
                warnings.append(warn)
        except OSError as e:
            warnings.append(f"{src.name}: не удалось архивировать — {e}")

    return dest_dir, moved, warnings


def validate_biflorica_download_dir(env: EnvSettings, raw_dir: str) -> str | None:
    """Возвращает текст ошибки или None, если путь допустим."""
    text = (raw_dir or "").strip()
    if not text:
        return "Укажите папку для скачивания."
    if _skip_local_path_check(text):
        return None
    try:
        resolved = resolve_biflorica_download_dir(env, text)
    except (OSError, ValueError) as e:
        return f"Некорректный путь: {e}"
    if resolved.exists() and not resolved.is_dir():
        return "Путь существует, но это не папка."
    try:
        resolved.mkdir(parents=True, exist_ok=True)
    except OSError as e:
        return f"Не удалось создать папку: {e}"
    return None


def validate_biflorica_archive_dir(
    env: EnvSettings,
    raw_archive_dir: str,
    raw_download_dir: str,
) -> str | None:
    download = validate_biflorica_download_dir(env, raw_download_dir)
    if download:
        return download
    if _skip_local_path_check(raw_archive_dir) or _skip_local_path_check(raw_download_dir):
        return None
    try:
        dl_path = resolve_biflorica_download_dir(env, raw_download_dir)
        arch_path = resolve_biflorica_archive_dir(env, raw_archive_dir, dl_path)
    except (OSError, ValueError) as e:
        return f"архива: некорректный путь — {e}"
    if arch_path.exists() and not arch_path.is_dir():
        return "архива: путь существует, но это не папка."
    try:
        arch_path.mkdir(parents=True, exist_ok=True)
    except OSError as e:
        return f"архива: не удалось создать папку — {e}"
    if arch_path.resolve() == dl_path.resolve():
        return "Папка архива не может совпадать с папкой скачивания."
    return None


def _default_ecuador_template(env: EnvSettings) -> Path:
    candidates = [
        env.project_root / "Invoice" / "3" / "Обработка" / "Прием товара Эквадор-4.xlsm",
        Path(r"C:\Invoice\3\Обработка\Прием товара Эквадор-4.xlsm"),
    ]
    for path in candidates:
        if path.is_file() and path.stat().st_size > 1024 and path.read_bytes()[:2] == b"PK":
            return path.resolve()
    return candidates[0].resolve()


def resolve_ecuador_template(env: EnvSettings, raw: str) -> Path:
    text = (raw or "").strip()
    if not text:
        return _default_ecuador_template(env)
    path = _resolve_dir(env, text, DEFAULT_ECUADOR_TEMPLATE)
    if path.is_dir():
        path = path / ECUADOR_TEMPLATE_FILENAME
    elif not path.is_file() and path.suffix.lower() not in (".xlsm", ".xls"):
        maybe = path / ECUADOR_TEMPLATE_FILENAME
        if maybe.is_file():
            path = maybe
    return path


def resolve_ecuador_output_dir(env: EnvSettings, raw: str) -> Path:
    text = (raw or "").strip()
    default = (
        DEFAULT_ECUADOR_OUTPUT_DIR
        if sys.platform == "win32"
        else str(env.project_root / "data" / "output" / "ecuador")
    )
    return _resolve_dir(env, text or default, default)


def validate_ecuador_paths(env: EnvSettings, raw_template: str, raw_output: str) -> str | None:
    if sys.platform != "win32":
        return None
    try:
        template = resolve_ecuador_template(env, raw_template)
        output = resolve_ecuador_output_dir(env, raw_output)
    except (OSError, ValueError) as e:
        return f"Эквадор: некорректный путь — {e}"
    if not template.is_file():
        return (
            f"Эквадор: файл шаблона не найден — {template}. "
            f"Укажите полный путь к {ECUADOR_TEMPLATE_FILENAME} "
            f"(например C:\\Invoice\\3\\Обработка\\{ECUADOR_TEMPLATE_FILENAME})."
        )
    try:
        with template.open("rb") as fh:
            if template.stat().st_size < 1024 or fh.read(2) != b"PK":
                return f"Эквадор: шаблон пустой или повреждён — {template}"
    except OSError as e:
        return f"Эквадор: не удалось прочитать шаблон — {e}"
    try:
        output.mkdir(parents=True, exist_ok=True)
    except OSError as e:
        return f"Эквадор: не удалось создать папку выгрузки — {e}"
    return None
