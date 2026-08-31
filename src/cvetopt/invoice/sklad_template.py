"""Копия файла «шаблон» в папке Инвойсы склад → имя = текущая дата."""

from __future__ import annotations

import shutil
from collections.abc import Callable
from datetime import date
from pathlib import Path

LogFn = Callable[[str], None]

_TEMPLATE_STEM = "шаблон"
_EXCEL_SUFFIXES = (".xls", ".xlsx", ".xlsm", ".XLS", ".XLSX", ".XLSM")


def _default_log(_msg: str) -> None:
    pass


def dated_template_name(on_date: date, *, suffix: str) -> str:
    """Имя копии: «шаблон ДД.ММ.ГГГГ» + расширение исходного файла."""
    ext = suffix if suffix.startswith(".") else f".{suffix}"
    return f"шаблон {on_date.strftime('%d.%m.%Y')}{ext.lower()}"


def find_sklad_template(sklad_dir: Path) -> Path:
    """
    Ищет файл с именем «шаблон» (+ расширение Excel) в корне папки склада.
    Регистр имени не важен; подпапки не смотрим.
    """
    if not sklad_dir.is_dir():
        raise FileNotFoundError(f"Папка склада не найдена: {sklad_dir}")

    matches: list[Path] = []
    want = _TEMPLATE_STEM.casefold()
    for path in sklad_dir.iterdir():
        if not path.is_file():
            continue
        if path.suffix.lower() not in {s.lower() for s in _EXCEL_SUFFIXES}:
            continue
        if path.stem.casefold() == want:
            matches.append(path)

    if not matches:
        raise FileNotFoundError(
            f"В {sklad_dir} нет файла «{_TEMPLATE_STEM}» "
            f"(.xls / .xlsx / .xlsm). Положите шаблон в эту папку."
        )
    if len(matches) > 1:
        names = ", ".join(p.name for p in matches)
        raise RuntimeError(
            f"В {sklad_dir} несколько файлов «{_TEMPLATE_STEM}»: {names}. "
            "Оставьте один."
        )
    return matches[0].resolve()


def copy_sklad_template_to_date(
    sklad_dir: Path,
    *,
    on_date: date | None = None,
    overwrite: bool = False,
    log: LogFn | None = None,
) -> Path:
    """
    Копирует «шаблон» → «шаблон ДД.ММ.ГГГГ.<ext>» в той же папке.
    Оригинал не трогает. Если файл на дату уже есть — ошибка (если не overwrite).
    """
    _lg = log or _default_log
    day = on_date or date.today()
    sklad_dir = sklad_dir.resolve()
    template = find_sklad_template(sklad_dir)
    dest_name = dated_template_name(day, suffix=template.suffix)
    dest = (sklad_dir / dest_name).resolve()

    if dest == template:
        raise RuntimeError(f"Имя копии совпало с шаблоном: {dest.name}")

    if dest.exists() and not overwrite:
        raise FileExistsError(
            f"Файл уже есть: {dest.name}. "
            "Удалите или переименуйте его, либо повторите с заменой "
            "(если сотрудник ещё не заполнял сетку)."
        )

    _lg(f"Шаблон: источник {template.name}")
    shutil.copy2(template, dest)
    _lg(f"Шаблон: создана копия → {dest.name}")
    # Убедимся, что оригинал на месте
    if not template.is_file():
        raise RuntimeError(f"После копирования пропал оригинал: {template}")
    return dest
