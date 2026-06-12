from __future__ import annotations

import os
import re
import sys
from collections.abc import Callable
from datetime import date, datetime
from pathlib import Path

from cvetopt.core.runtime_settings import _archive_one_entry, _archive_target_path

from cvetopt.invoice.description_dictionary import (
    append_missing_descriptions,
    load_description_dictionary,
    lookup_translation,
)
from cvetopt.invoice.xlsx_patch import patch_xlsx_cell_values
from cvetopt.invoice.xlsx_read import grid_by_row, read_xlsx_grid

LogFn = Callable[[str], None]


def _default_log(_msg: str) -> None:
    pass


def _next_col_letter(col: str) -> str:
    n = 0
    for ch in col.upper():
        n = n * 26 + (ord(ch) - 64)
    n += 1
    out: list[str] = []
    while n:
        n, rem = divmod(n - 1, 26)
        out.append(chr(65 + rem))
    return "".join(reversed(out))


def holland_export_candidates(output_dir: Path, on_date: date | None = None) -> list[Path]:
    if not output_dir.is_dir():
        return []
    names: list[str] = ["Голландия_1_*.xlsx", "Голландия_1_*.xlsm"]
    if on_date is not None:
        names.insert(0, f"Голландия_1_{on_date.strftime('%d.%m.%Y')}.xlsx")
        names.insert(1, f"Голландия_1_{on_date.day}.{on_date.month}.{on_date.year}.xlsx")
    seen: set[Path] = set()
    found: list[Path] = []
    for pattern in names:
        for p in output_dir.glob(pattern):
            rp = p.resolve()
            if rp not in seen and rp.is_file():
                seen.add(rp)
                found.append(rp)
    if found:
        return sorted(found, key=lambda p: p.stat().st_mtime, reverse=True)
    found = sorted(
        list(output_dir.glob("Голландия_1_*.xlsx"))
        + list(output_dir.glob("Голландия_1_*.xlsm")),
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )
    return found


def find_holland_export_file(
    output_dir: Path,
    *,
    on_date: date | None = None,
) -> Path | None:
    items = holland_export_candidates(output_dir, on_date=on_date)
    return items[0] if items else None


def archive_stale_holland_exports(
    sklad_dir: Path,
    archive_dir: Path,
    *,
    keep_path: Path | None = None,
    log: LogFn | None = None,
) -> tuple[Path | None, list[str], list[str]]:
    """
    Оставляет один файл Голландия_1_*.xlsx (самый свежий или keep_path),
    остальные переносит прямо в папку архива.
    """
    _lg = log or _default_log
    if not sklad_dir.is_dir():
        return None, [], []

    archive_dir = archive_dir.resolve()
    sklad_dir = sklad_dir.resolve()
    if archive_dir == sklad_dir:
        raise ValueError("Папка архива не может совпадать с папкой Голландия.")

    candidates: list[Path] = []
    for path in holland_export_candidates(sklad_dir):
        try:
            path.resolve().relative_to(archive_dir)
        except ValueError:
            candidates.append(path)

    if not candidates:
        return None, [], []

    keep = (keep_path or candidates[0]).resolve()
    to_move = [path for path in candidates if path.resolve() != keep]
    if not to_move:
        _lg("Голландия: старых файлов для архива нет")
        return None, [], []

    archive_dir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y-%m-%d_%H%M%S")

    moved: list[str] = []
    warnings: list[str] = []
    if sys.platform == "win32" and not os.access(archive_dir, os.W_OK):
        import getpass

        warnings.append(
            f"Папка архива {archive_dir}: нет записи для «{getpass.getuser()}»."
        )

    for src in sorted(to_move, key=lambda p: p.name.lower()):
        target = _archive_target_path(archive_dir, src, stamp)
        if not os.access(src, os.R_OK):
            warnings.append(f"{src.name}: пропуск (нет чтения)")
            continue
        try:
            warn = _archive_one_entry(src, target)
            moved.append(src.name)
            if warn:
                warnings.append(warn)
        except OSError as e:
            warnings.append(f"{src.name}: не удалось архивировать — {e}")

    if moved:
        _lg(f"Голландия: в архив {len(moved)} → {archive_dir}")
    for warn in warnings:
        _lg(f"Голландия (архив): {warn}")

    return (archive_dir if moved else None), moved, warnings


def _find_description_columns(header: dict[str, str]) -> tuple[str, str] | None:
    for col, val in header.items():
        if str(val).strip().casefold() == "description":
            return col, _next_col_letter(col)
    return None


def _build_translation_plan(
    export_path: Path,
    dictionary: dict[str, str],
    *,
    log: LogFn,
) -> tuple[str, str, dict[str, str | None], list[str], int, int, int]:
    grid = read_xlsx_grid(export_path)
    rows = grid_by_row(grid)
    header = rows.get(1, {})
    cols = _find_description_columns(header)
    if cols is None:
        raise RuntimeError(
            f"В {export_path.name} нет заголовка Description в первой строке."
        )
    desc_col, trans_col = cols
    log(f"Description → колонка {trans_col} (рядом с {desc_col})")

    updates: dict[str, str | None] = {}
    missing_texts: list[str] = []
    translated = 0
    missing = 0
    total = 0
    max_row = max(rows) if rows else 1
    for row_n in range(2, max_row + 1):
        row = rows.get(row_n, {})
        text = _norm_cell(row.get(desc_col, ""))
        if not text:
            continue
        total += 1
        tr = lookup_translation(dictionary, text)
        ref = f"{trans_col}{row_n}"
        if tr:
            updates[ref] = tr
            translated += 1
        else:
            updates[ref] = None
            missing += 1
            missing_texts.append(text)
    return desc_col, trans_col, updates, missing_texts, translated, missing, total


def _translate_via_xlwings(
    export_path: Path,
    trans_col: str,
    updates: dict[str, str | None],
    *,
    log: LogFn,
) -> None:
    import xlwings as xw

    rows_sorted = sorted(
        (int(m.group(2)), value)
        for ref, value in updates.items()
        if (m := re.match(r"^([A-Z]+)(\d+)$", ref))
    )
    if not rows_sorted:
        return

    app: object | None = None
    wb: object | None = None
    try:
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        app.screen_updating = False
        wb = app.books.open(str(export_path), update_links=False)
        ws = wb.sheets[0]
        first_row = rows_sorted[0][0]
        last_row = rows_sorted[-1][0]
        values: list[list[object]] = []
        row_map = dict(rows_sorted)
        for row_n in range(first_row, last_row + 1):
            val = row_map.get(row_n)
            values.append([val if val else None])
        ws.range(f"{trans_col}{first_row}").resize(len(values), 1).value = values
        wb.save()
        log("Сохранено через Excel (xlwings) — структура файла не меняется.")
    finally:
        if wb is not None:
            try:
                wb.close()
            except Exception:
                pass
        if app is not None:
            try:
                app.quit()
            except Exception:
                pass


def translate_holland_export(
    export_path: Path,
    dictionary_path: Path,
    *,
    append_missing_to_dictionary: bool = True,
    log: LogFn | None = None,
) -> tuple[int, int, int]:
    """
    Заполняет столбец справа от Description переводом из словаря.
    На Windows — через Excel (xlwings), иначе точечный XML-патч без ElementTree.
    """
    _lg = log or _default_log
    export_path = export_path.resolve()
    dictionary = load_description_dictionary(dictionary_path)
    if not dictionary:
        raise RuntimeError(f"Словарь пуст: {dictionary_path}")

    _lg(f"Словарь: {len(dictionary)} записей из {dictionary_path.name}")

    _desc_col, trans_col, updates, missing_texts, translated, missing, total = (
        _build_translation_plan(export_path, dictionary, log=_lg)
    )

    if append_missing_to_dictionary and missing_texts:
        try:
            append_missing_descriptions(
                dictionary_path,
                missing_texts,
                dictionary=dictionary,
                log=_lg,
            )
        except Exception as e:
            _lg(f"Словарь: не удалось дописать без перевода — {e}")

    if sys.platform == "win32":
        try:
            _translate_via_xlwings(export_path, trans_col, updates, log=_lg)
            _lg(
                f"Перевод в {export_path.name}: строк {total}, "
                f"переведено {translated}, без перевода {missing}."
            )
            return translated, missing, total
        except Exception as e:
            _lg(f"xlwings: {e} — пробую XML-патч…")

    patch_xlsx_cell_values(export_path, updates)
    _lg(
        f"Перевод в {export_path.name}: строк {total}, "
        f"переведено {translated}, без перевода {missing} "
        "(XML-патч, чекбоксы и разметка сохранены)."
    )
    return translated, missing, total


def _norm_cell(val: str) -> str:
    return str(val or "").strip()


def postprocess_holland_after_auto1(
    *,
    sklad_output_dir: Path,
    dictionary_path: Path,
    on_date: date | None = None,
    append_missing_to_dictionary: bool = True,
    add_row_markers: bool = False,
    marker_assets_dir: Path | None = None,
    log: LogFn | None = None,
) -> Path | None:
    """Ищет свежий Голландия_1_*.xlsx в папке склада и переводит Description."""
    _lg = log or _default_log
    export_file = find_holland_export_file(sklad_output_dir, on_date=on_date)
    if export_file is None:
        _lg(f"Файл Голландия_1_*.xlsx не найден в {sklad_output_dir}")
        return None
    translate_holland_export(
        export_file,
        dictionary_path,
        append_missing_to_dictionary=append_missing_to_dictionary,
        log=_lg,
    )
    if add_row_markers and marker_assets_dir is not None:
        from cvetopt.invoice.holland_markers import add_holland_row_markers

        try:
            export_file = add_holland_row_markers(
                export_file,
                marker_assets_dir,
                log=_lg,
            )
        except Exception as e:
            _lg(f"Голландия: маркеры пропущены — {e}")
    return export_file
