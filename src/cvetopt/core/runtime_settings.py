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

from dataclasses import dataclass

from pydantic import BaseModel, Field

from cvetopt.core.settings import EnvSettings, MailConfig, _resolve_selection

DEFAULT_BIFLORICA_DOWNLOAD_DIR = "data/downloads/biflorica"
DEFAULT_BIFLORICA_ARCHIVE_DIR = (
    r"C:\Инвойсы склад\архив"
    if sys.platform == "win32"
    else "data/downloads/biflorica/архив"
)
BIFLORICA_ARCHIVE_SUBDIR = "архив"
DEFAULT_ECUADOR_TEMPLATE = "Invoice/3/Обработка/Прием товара Эквадор-4.xlsm"
ECUADOR_TEMPLATE_FILENAME = "Прием товара Эквадор-4.xlsm"
DEFAULT_ECUADOR_OUTPUT_DIR = r"D:\Склад ОБмен\Инвойсы Склад"
ECUADOR_ARCHIVE_SUBDIR = "Архивы Инвойсы"
DEFAULT_ECUADOR_ARCHIVE_DIR = (
    r"C:\Инвойсы склад\Архивы Инвойсы"
    if sys.platform == "win32"
    else "data/output/ecuador/Архивы Инвойсы"
)
DEFAULT_MAIL_OUTPUT_DIR_SHORT = "data/downloads/mail/1"
DEFAULT_MAIL_OUTPUT_DIR_LONG = "data/downloads/mail/2"
DEFAULT_MAIL_FILENAME_SHORT_MAX_LEN = 35
DEFAULT_AUTO_NEW_WORKBOOK = "Auto_new.xls"
DEFAULT_HOLLAND_DICTIONARY = r"C:\Invoice\Словарь.xls"
DEFAULT_HOLLAND_SKLAD_DIR = r"C:\Инвойсы склад"
BIFLORICA_ARCHIVE_LEGACY_NAMES = frozenset({"архив", "archive"})
BIFLORICA_DOWNLOAD_PREFIX = "BiFlorica-"
# Скрипт: BiFlorica-<order_id>__<YYYY-MM-DD>.xlsx; старые без префикса тоже в архив.
_BIFLORICA_REPORT_STEM_RE = re.compile(
    r"^(?:BiFlorica-)?\d+__\d{4}-\d{2}-\d{2}$",
    re.IGNORECASE,
)
_BIFLORICA_ORDER_ID_FROM_FILE_RE = re.compile(
    r"^(?:BiFlorica-)?(\d+)__\d{4}-\d{2}-\d{2}$",
    re.IGNORECASE,
)


def biflorica_download_filename(order_id: str, flight_date: date) -> str:
    """Имя xlsx отчёта в папке скачивания (префикс BiFlorica- для отличия от прочих файлов)."""
    return f"{BIFLORICA_DOWNLOAD_PREFIX}{order_id}__{flight_date.isoformat()}.xlsx"


def order_id_from_biflorica_report(path: Path) -> str | None:
    """Номер заказа из имени BiFlorica-<id>__<дата>.xlsx или <id>__<дата>.xlsx."""
    m = _BIFLORICA_ORDER_ID_FROM_FILE_RE.match(path.stem)
    return m.group(1) if m else None


class RuntimeSettings(BaseModel):
    biflorica_min_age_days: int
    biflorica_max_age_days: int
    biflorica_download_dir: str = Field(default=DEFAULT_BIFLORICA_DOWNLOAD_DIR)
    biflorica_archive_dir: str = Field(default=DEFAULT_BIFLORICA_ARCHIVE_DIR)
    ecuador_template_path: str = Field(default=DEFAULT_ECUADOR_TEMPLATE)
    ecuador_output_dir: str = Field(default=DEFAULT_ECUADOR_OUTPUT_DIR)
    ecuador_archive_dir: str = Field(default=DEFAULT_ECUADOR_ARCHIVE_DIR)
    delmir_lookback_days: int
    mail_lookback_days: int
    mail_output_dir_short: str = Field(default=DEFAULT_MAIL_OUTPUT_DIR_SHORT)
    mail_output_dir_long: str = Field(default=DEFAULT_MAIL_OUTPUT_DIR_LONG)
    mail_filename_short_max_len: int = Field(default=DEFAULT_MAIL_FILENAME_SHORT_MAX_LEN)
    mail_archive_dir: str = Field(default="")
    auto_new_workbook_path: str = Field(default=DEFAULT_AUTO_NEW_WORKBOOK)
    holland_dictionary_path: str = Field(default=DEFAULT_HOLLAND_DICTIONARY)
    holland_sklad_output_dir: str = Field(default=DEFAULT_HOLLAND_SKLAD_DIR)


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
        ecuador_archive_dir=DEFAULT_ECUADOR_ARCHIVE_DIR,
        delmir_lookback_days=yaml_cfg.delmir.lookback_days,
        mail_lookback_days=yaml_cfg.mail.lookback_days,
        mail_output_dir_short=yaml_cfg.mail.output_dir_short,
        mail_output_dir_long=yaml_cfg.mail.output_dir_long,
        mail_filename_short_max_len=yaml_cfg.mail.filename_short_max_len,
        auto_new_workbook_path=yaml_cfg.balance_auto.workbook_path,
        holland_dictionary_path=yaml_cfg.holland_translate.dictionary_path,
        holland_sklad_output_dir=yaml_cfg.holland_translate.sklad_output_dir,
    )


@dataclass(frozen=True)
class MailOutputLayout:
    short_dir: Path
    long_dir: Path
    short_max_len: int


def is_short_mail_filename(filename: str, max_len: int) -> bool:
    """Короткое имя (обычно дата + короткий суффикс) → папка 1."""
    return len(filename) <= max_len


def resolve_mail_output_layout(
    env: EnvSettings,
    runtime: RuntimeSettings,
    mail_cfg: MailConfig | None = None,
) -> MailOutputLayout:
    yaml_mail = (mail_cfg or env.yaml_config().mail)
    short_raw = (runtime.mail_output_dir_short or "").strip() or yaml_mail.output_dir_short
    long_raw = (runtime.mail_output_dir_long or "").strip() or yaml_mail.output_dir_long
    max_len = runtime.mail_filename_short_max_len
    if max_len < 1:
        max_len = yaml_mail.filename_short_max_len
    return MailOutputLayout(
        short_dir=_resolve_dir(env, short_raw, DEFAULT_MAIL_OUTPUT_DIR_SHORT),
        long_dir=_resolve_dir(env, long_raw, DEFAULT_MAIL_OUTPUT_DIR_LONG),
        short_max_len=max_len,
    )


def mail_destination_dir(filename: str, layout: MailOutputLayout) -> Path:
    if is_short_mail_filename(filename, layout.short_max_len):
        return layout.short_dir
    return layout.long_dir


def resolve_mail_archive_dir(
    env: EnvSettings,
    runtime: RuntimeSettings,
    layout: MailOutputLayout,
    mail_cfg: MailConfig | None = None,
) -> Path:
    """Пусто в UI/config → <родитель папки 1>/архив."""
    yaml_mail = mail_cfg or env.yaml_config().mail
    raw = (
        (runtime.mail_archive_dir or "").strip()
        or (yaml_mail.archive_dir or "").strip()
    )
    default = layout.short_dir.parent / BIFLORICA_ARCHIVE_SUBDIR
    if not raw:
        return default.resolve()
    return _resolve_dir(env, raw, str(default))


def validate_mail_archive_dir(
    env: EnvSettings,
    runtime: RuntimeSettings,
    *,
    raw_archive_dir: str = "",
) -> str | None:
    try:
        layout = resolve_mail_output_layout(env, runtime)
        arch = resolve_mail_archive_dir(
            env,
            runtime.model_copy(update={"mail_archive_dir": raw_archive_dir}),
            layout,
        )
    except (OSError, ValueError) as e:
        return f"Почта, архив: некорректный путь — {e}"
    if arch.exists() and not arch.is_dir():
        return "Почта, архив: путь существует, но это не папка."
    try:
        arch.mkdir(parents=True, exist_ok=True)
    except OSError as e:
        return f"Почта, архив: не удалось создать папку — {e}"
    short = layout.short_dir.resolve()
    long = layout.long_dir.resolve()
    arch_r = arch.resolve()
    if arch_r in (short, long):
        return "Почта: папка архива не может совпадать с папками 1 или 2."
    return None


def validate_mail_output_dirs(env: EnvSettings, runtime: RuntimeSettings) -> str | None:
    try:
        layout = resolve_mail_output_layout(env, runtime)
    except (OSError, ValueError) as e:
        return f"Почта: некорректный путь — {e}"
    if runtime.mail_filename_short_max_len < 1 or runtime.mail_filename_short_max_len > 500:
        return "Почта: длина «короткого» имени — от 1 до 500 символов."
    for label, path in (("1 (короткие имена)", layout.short_dir), ("2 (длинные имена)", layout.long_dir)):
        if path.exists() and not path.is_dir():
            return f"Почта, папка {label}: путь существует, но это не папка."
        try:
            path.mkdir(parents=True, exist_ok=True)
        except OSError as e:
            return f"Почта, папка {label}: не удалось создать — {e}"
    if layout.short_dir.resolve() == layout.long_dir.resolve():
        return "Почта: папки 1 и 2 не должны совпадать."
    return None


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


def resolve_auto_new_workbook(env: EnvSettings, raw: str) -> Path:
    """Путь к Auto_new.xls: относительный к корню проекта или абсолютный (C:\\...)."""
    return _resolve_dir(env, raw, DEFAULT_AUTO_NEW_WORKBOOK)


def effective_auto_new_workbook_raw(
    runtime: RuntimeSettings,
    *,
    yaml_auto1: str = "",
    yaml_balance: str = "",
) -> str:
    text = (runtime.auto_new_workbook_path or "").strip()
    if text:
        return text
    return (yaml_auto1 or "").strip() or (yaml_balance or "").strip() or DEFAULT_AUTO_NEW_WORKBOOK


def validate_auto_new_workbook_path(env: EnvSettings, raw: str) -> str | None:
    if _skip_local_path_check(raw):
        return None
    try:
        path = resolve_auto_new_workbook(env, raw)
    except (OSError, ValueError) as e:
        return f"Auto_new.xls: некорректный путь — {e}"
    if not path.is_file():
        return (
            f"Auto_new.xls: файл не найден — {path}. "
            "Укажите полный путь (например C:\\Apps\\cvetopt\\Auto_new.xls) "
            "или относительный к корню проекта."
        )
    if path.stat().st_size < 1024:
        return f"Auto_new.xls: файл слишком маленький или пустой — {path}"
    return None


def resolve_holland_dictionary(env: EnvSettings, raw: str) -> Path:
    """Путь к Словарь.xls (файл): абсолютный или относительно корня проекта."""
    return _resolve_dir(env, raw, DEFAULT_HOLLAND_DICTIONARY)


def resolve_holland_sklad_dir(env: EnvSettings, raw: str) -> Path:
    return _resolve_dir(env, raw, DEFAULT_HOLLAND_SKLAD_DIR)


def effective_holland_dictionary_raw(
    runtime: RuntimeSettings,
    *,
    yaml_path: str = "",
) -> str:
    text = (runtime.holland_dictionary_path or "").strip()
    if text:
        return text
    return (yaml_path or "").strip() or DEFAULT_HOLLAND_DICTIONARY


def effective_holland_sklad_dir_raw(
    runtime: RuntimeSettings,
    *,
    yaml_dir: str = "",
) -> str:
    text = (runtime.holland_sklad_output_dir or "").strip()
    if text:
        return text
    return (yaml_dir or "").strip() or DEFAULT_HOLLAND_SKLAD_DIR


def validate_holland_translate_paths(
    env: EnvSettings,
    raw_dictionary: str,
    raw_sklad_dir: str,
) -> str | None:
    if _skip_local_path_check(raw_dictionary) and _skip_local_path_check(raw_sklad_dir):
        return None
    try:
        dictionary = resolve_holland_dictionary(env, raw_dictionary)
        sklad = resolve_holland_sklad_dir(env, raw_sklad_dir)
    except (OSError, ValueError) as e:
        return f"Перевод: некорректный путь — {e}"
    if not _skip_local_path_check(raw_dictionary):
        if not dictionary.is_file():
            return (
                f"Перевод: словарь не найден — {dictionary}. "
                f"Укажите полный путь (например {DEFAULT_HOLLAND_DICTIONARY})."
            )
        if dictionary.stat().st_size < 32:
            return f"Перевод: словарь пустой или повреждён — {dictionary}"
    if not _skip_local_path_check(raw_sklad_dir):
        if sklad.exists() and not sklad.is_dir():
            return f"Перевод: папка склада — не каталог: {sklad}"
    return None


def resolve_biflorica_download_dir(env: EnvSettings, raw_dir: str) -> Path:
    """Абсолютный или относительный (к корню проекта) каталог для xlsx Biflorica."""
    return _resolve_dir(env, raw_dir, DEFAULT_BIFLORICA_DOWNLOAD_DIR)


def effective_biflorica_archive_dir_raw(
    runtime: RuntimeSettings,
    *,
    yaml_default: str = "",
) -> str:
    """Пустое поле в UI → архив рядом с папкой Эквадор/склад."""
    text = (runtime.biflorica_archive_dir or "").strip()
    if text:
        return text
    ecu = (runtime.ecuador_output_dir or "").strip()
    if ecu:
        return str(Path(ecu) / BIFLORICA_ARCHIVE_SUBDIR)
    sklad = (runtime.holland_sklad_output_dir or "").strip()
    if sklad:
        return str(Path(sklad) / BIFLORICA_ARCHIVE_SUBDIR)
    yaml_text = (yaml_default or "").strip()
    if yaml_text:
        return yaml_text
    return DEFAULT_BIFLORICA_ARCHIVE_DIR


def resolve_biflorica_archive_dir(
    env: EnvSettings,
    raw_dir: str,
    download_dir: Path | None = None,
    *,
    runtime: RuntimeSettings | None = None,
    yaml_default: str = "",
) -> Path:
    """Каталог архива Biflorica (настраивается в UI; по умолчанию — Инвойсы склад/архив)."""
    text = (raw_dir or "").strip()
    if not text and runtime is not None:
        text = effective_biflorica_archive_dir_raw(runtime, yaml_default=yaml_default)
    if not text:
        base = download_dir or resolve_biflorica_download_dir(env, "")
        return (base / BIFLORICA_ARCHIVE_SUBDIR).resolve()
    return _resolve_dir(env, text, DEFAULT_BIFLORICA_ARCHIVE_DIR)


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


def _format_archive_oserror(src: Path, target: Path, exc: OSError, phase: str) -> str:
    """Понятное сообщение для WinError 5 при архивировании Biflorica."""
    import getpass

    user = getpass.getuser()
    base = f"{src.name}: {phase} — {exc}"
    if not _is_access_denied(exc):
        return base
    hints: list[str] = []
    if phase.startswith("копирование"):
        if not os.access(src, os.R_OK):
            hints.append(f"нет чтения файла «{src}»")
        if not os.access(target.parent, os.W_OK):
            hints.append(f"нет записи в папку архива «{target.parent}»")
        if not hints:
            hints.append("файл может быть открыт в Excel — закройте книгу")
    elif phase.startswith("удаление"):
        hints.append("копия уже в архиве; исходник не удалён")
        if not os.access(src.parent, os.W_OK):
            hints.append(f"нет прав на изменение «{src.parent}»")
    detail = "; ".join(hints) if hints else "отказано в доступе (WinError 5)"
    return (
        f"{src.name}: {detail}. Процесс: «{user}». "
        "Запускайте cvetopt.bat под той же учёткой, что создаёт файлы в C:\\Invoice, "
        "или выдайте Modify на папку скачивания и архив (см. README_WIN.md §4.4)."
    )


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
        try:
            _copy_with_retries(src, target)
        except OSError as e:
            raise OSError(_format_archive_oserror(src, target, e, "копирование в архив")) from e
        try:
            src.unlink()
            return None
        except OSError as e:
            if _is_access_denied(e):
                return _format_archive_oserror(src, target, e, "удаление исходника")
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


def _biflorica_archive_candidate(
    entry: Path,
    *,
    policy: str,
    keep_order_ids: set[str],
    keep_paths: set[Path],
) -> bool:
    order_id = order_id_from_biflorica_report(entry)
    resolved = entry.resolve()
    if policy == "unregistered_only":
        return not (order_id and order_id in keep_order_ids)
    if policy == "stale_registered":
        # В архив всё, кроме отчётов, скачанных в текущей сессии (keep_paths).
        return resolved not in keep_paths
    raise ValueError(f"Неизвестная политика архива Biflorica: {policy}")


def archive_biflorica_download_dir(
    download_dir: Path,
    archive_dir: Path,
    *,
    keep_order_ids: set[str] | None = None,
    keep_paths: set[Path] | None = None,
    policy: str = "unregistered_only",
) -> tuple[Path | None, list[str], list[str], list[str]]:
    """
    Переносит xlsx/xls отчётов Biflorica из корня папки скачивания прямо в папку архива.

    policy=unregistered_only (до скачивания): в архив всё, чего нет в реестре keep_order_ids.
    policy=stale_registered (после скачивания): в архив все отчёты Biflorica, кроме keep_paths (сессия).
    """
    if not download_dir.is_dir():
        return None, [], [], []

    archive_dir = archive_dir.resolve()
    download_dir = download_dir.resolve()
    if archive_dir == download_dir:
        raise ValueError("Папка архива не может совпадать с папкой скачивания.")

    keep_ids = keep_order_ids or set()
    keep_resolved = {p.resolve() for p in (keep_paths or set())}
    to_move: list[Path] = []
    kept: list[str] = []
    for entry in download_dir.iterdir():
        if _is_archive_entry(entry, archive_dir, download_dir):
            continue
        if not _is_biflorica_report_to_archive(entry):
            continue
        if not _biflorica_archive_candidate(
            entry,
            policy=policy,
            keep_order_ids=keep_ids,
            keep_paths=keep_resolved,
        ):
            kept.append(entry.name)
            continue
        to_move.append(entry)

    if not to_move:
        return None, [], [], kept

    archive_dir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y-%m-%d_%H%M%S")

    moved: list[str] = []
    warnings: list[str] = []
    if sys.platform == "win32":
        import getpass

        user = getpass.getuser()
        if not os.access(download_dir, os.W_OK | os.R_OK):
            warnings.append(
                f"Папка скачивания {download_dir}: нет чтения/записи для «{user}»."
            )
        if not os.access(archive_dir, os.W_OK):
            warnings.append(
                f"Папка архива {archive_dir}: нет записи для «{user}» — "
                "выдайте Modify (M) на эту папку."
            )
    access_hint_shown = False
    for src in sorted(to_move, key=lambda p: p.name.lower()):
        target = _archive_target_path(archive_dir, src, stamp)
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

    return archive_dir, moved, warnings, kept


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
    *,
    runtime: RuntimeSettings | None = None,
) -> str | None:
    download = validate_biflorica_download_dir(env, raw_download_dir)
    if download:
        return download
    if _skip_local_path_check(raw_archive_dir) or _skip_local_path_check(raw_download_dir):
        return None
    try:
        dl_path = resolve_biflorica_download_dir(env, raw_download_dir)
        arch_path = resolve_biflorica_archive_dir(
            env,
            raw_archive_dir,
            dl_path,
            runtime=runtime,
        )
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


def effective_ecuador_archive_dir_raw(
    runtime: RuntimeSettings,
    *,
    output_dir: Path | None = None,
) -> str:
    """Пустое поле в UI → <папка выгрузки Эквадор>/Архивы Инвойсы."""
    text = (runtime.ecuador_archive_dir or "").strip()
    if text:
        return text
    out_raw = (runtime.ecuador_output_dir or "").strip()
    if out_raw:
        return str(Path(out_raw) / ECUADOR_ARCHIVE_SUBDIR)
    if output_dir is not None:
        return str(output_dir / ECUADOR_ARCHIVE_SUBDIR)
    return DEFAULT_ECUADOR_ARCHIVE_DIR


def resolve_ecuador_archive_dir(
    env: EnvSettings,
    raw_dir: str,
    output_dir: Path,
    *,
    runtime: RuntimeSettings | None = None,
) -> Path:
    """Каталог архива Эквадор (настраивается в UI)."""
    text = (raw_dir or "").strip()
    if not text and runtime is not None:
        text = effective_ecuador_archive_dir_raw(runtime, output_dir=output_dir)
    if not text:
        return (output_dir / ECUADOR_ARCHIVE_SUBDIR).resolve()
    return _resolve_dir(env, text, DEFAULT_ECUADOR_ARCHIVE_DIR)


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


def validate_ecuador_archive_dir(
    env: EnvSettings,
    raw_archive_dir: str,
    raw_output_dir: str,
    *,
    runtime: RuntimeSettings | None = None,
) -> str | None:
    if sys.platform != "win32":
        return None
    if _skip_local_path_check(raw_archive_dir) or _skip_local_path_check(raw_output_dir):
        return None
    try:
        output = resolve_ecuador_output_dir(env, raw_output_dir)
        arch = resolve_ecuador_archive_dir(
            env,
            raw_archive_dir,
            output,
            runtime=runtime,
        )
    except (OSError, ValueError) as e:
        return f"Эквадор, архив: некорректный путь — {e}"
    if arch.exists() and not arch.is_dir():
        return "Эквадор, архив: путь существует, но это не папка."
    try:
        arch.mkdir(parents=True, exist_ok=True)
    except OSError as e:
        return f"Эквадор, архив: не удалось создать папку — {e}"
    if arch.resolve() == output.resolve():
        return "Папка архива Эквадор не может совпадать с папкой выгрузки."
    return None
