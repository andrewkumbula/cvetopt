from __future__ import annotations

import json
import shutil
from datetime import datetime
from pathlib import Path

from pydantic import BaseModel, Field

from cvetopt.core.settings import EnvSettings, _resolve_selection

DEFAULT_BIFLORICA_DOWNLOAD_DIR = "data/downloads/biflorica"
DEFAULT_BIFLORICA_ARCHIVE_DIR = "data/downloads/biflorica/архив"
DEFAULT_ECUADOR_TEMPLATE = "Invoice/3/Обработка/Прием товара Эквадор-4.xlsm"
DEFAULT_ECUADOR_OUTPUT_DIR = r"D:\Склад ОБмен\Инвойсы Склад"
BIFLORICA_ARCHIVE_LEGACY_NAMES = frozenset({"архив", "archive"})


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


def _resolve_dir(env: EnvSettings, raw_dir: str, default: str) -> Path:
    text = (raw_dir or "").strip() or default
    path = Path(text).expanduser()
    if not path.is_absolute():
        path = env.project_root / path
    return path.resolve()


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


def archive_biflorica_download_dir(
    download_dir: Path,
    archive_dir: Path,
) -> tuple[Path | None, list[str]]:
    """
    Переносит файлы и подпапки из папки скачивания в
    <папка архива>/<YYYY-MM-DD_HHMMSS>/.
    """
    if not download_dir.is_dir():
        return None, []

    archive_dir = archive_dir.resolve()
    download_dir = download_dir.resolve()
    if archive_dir == download_dir:
        raise ValueError("Папка архива не может совпадать с папкой скачивания.")

    to_move: list[Path] = []
    for entry in download_dir.iterdir():
        if _is_archive_entry(entry, archive_dir, download_dir):
            continue
        if entry.is_file() or entry.is_dir():
            to_move.append(entry)

    if not to_move:
        return None, []

    stamp = datetime.now().strftime("%Y-%m-%d_%H%M%S")
    dest_dir = archive_dir / stamp
    dest_dir.mkdir(parents=True, exist_ok=True)

    moved: list[str] = []
    for src in sorted(to_move, key=lambda p: p.name.lower()):
        target = dest_dir / src.name
        if target.exists():
            suffix = src.suffix if src.is_file() else ""
            stem = src.stem if src.is_file() else src.name
            target = dest_dir / f"{stem}_{stamp}{suffix}"
        shutil.move(str(src), str(target))
        moved.append(src.name)

    return dest_dir, moved


def validate_biflorica_download_dir(env: EnvSettings, raw_dir: str) -> str | None:
    """Возвращает текст ошибки или None, если путь допустим."""
    text = (raw_dir or "").strip()
    if not text:
        return "Укажите папку для скачивания."
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


def validate_ecuador_paths(env: EnvSettings, raw_template: str, raw_output: str) -> str | None:
    from cvetopt.invoice.ecuador_create import resolve_ecuador_output_dir, resolve_ecuador_template

    try:
        template = resolve_ecuador_template(env, raw_template)
        output = resolve_ecuador_output_dir(env, raw_output)
    except (OSError, ValueError) as e:
        return f"Эквадор: некорректный путь — {e}"
    if not template.is_file():
        return f"Эквадор: шаблон не найден — {template}"
    if template.stat().st_size < 1024 or template.read_bytes()[:2] != b"PK":
        return f"Эквадор: шаблон пустой или повреждён — {template}"
    try:
        output.mkdir(parents=True, exist_ok=True)
    except OSError as e:
        return f"Эквадор: не удалось создать папку выгрузки — {e}"
    return None
