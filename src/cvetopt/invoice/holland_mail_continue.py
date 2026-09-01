"""
Кнопка «Переведено»: перевод Description в почтовых файлах папок 1 и 2,
ростовка у всех роз, затем Auto1 до Sort (без Group / For sklad).
"""

from __future__ import annotations

import re
import sys
from collections.abc import Callable
from dataclasses import dataclass
from pathlib import Path

from cvetopt.invoice.description_dictionary import (
    SkipDescriptionRules,
    SKIP_DESCRIPTION_RULES_EMPTY,
    append_missing_descriptions,
    is_holland_product_description,
    load_description_dictionary,
    translate_description,
)
from cvetopt.invoice.xlsx_patch import patch_xlsx_cell_values
from cvetopt.invoice.xlsx_read import grid_by_row, read_xlsx_grid
from cvetopt.mail.short_postprocess import _copy_workbook_palette, _open_xls_workbook

LogFn = Callable[[str], None]

_EXCEL_SUFFIXES = {".xls", ".xlsx", ".xlsm"}
_CM_RE = re.compile(r"(?i)\s*(cm|см)\s*$")


def _default_log(_msg: str) -> None:
    pass


def _norm(text: object) -> str:
    return " ".join(str(text or "").split()).strip()


def _col_letter(index0: int) -> str:
    n = index0 + 1
    out: list[str] = []
    while n:
        n, rem = divmod(n - 1, 26)
        out.append(chr(65 + rem))
    return "".join(reversed(out))


def _col_index(letter: str) -> int:
    n = 0
    for ch in letter.upper():
        n = n * 26 + (ord(ch) - 64)
    return n - 1


def find_single_mail_workbook(folder: Path, *, log: LogFn) -> Path:
    """Один .xls/.xlsx в корне папки (без архива). При нескольких — самый свежий."""
    if not folder.is_dir():
        raise FileNotFoundError(f"Папка не найдена: {folder}")
    files = [
        p
        for p in folder.iterdir()
        if p.is_file() and p.suffix.lower() in _EXCEL_SUFFIXES
    ]
    if not files:
        raise FileNotFoundError(f"В {folder} нет Excel-файлов (.xls/.xlsx)")
    files.sort(key=lambda p: p.stat().st_mtime, reverse=True)
    if len(files) > 1:
        names = ", ".join(p.name for p in files[:5])
        log(f"В {folder.name}: файлов {len(files)}, беру самый свежий ({files[0].name}). Остальные: {names}")
    return files[0].resolve()


def length_without_cm(value: object) -> str:
    """«70 Cm» / 70.0 / «70» → «70»."""
    if isinstance(value, float):
        if value == int(value):
            return str(int(value))
        return str(value).rstrip("0").rstrip(".")
    if isinstance(value, int):
        return str(value)
    text = _norm(value)
    if not text:
        return ""
    text = _CM_RE.sub("", text).strip()
    try:
        num = float(text.replace(",", "."))
        if num == int(num):
            return str(int(num))
    except ValueError:
        pass
    return text


def is_rose_description(text: str) -> bool:
    """
    Роза Enigma: первый токен «R» / «Rose» / «Роза».
    Не путать с сортом, где «Rose» в середине («Eus G Rosi Rose Pink»).
    """
    tokens = _norm(text).split()
    if not tokens:
        return False
    first = tokens[0].casefold()
    return first in {"r", "rose", "роза"}


def _already_has_rostovka(name: str, length: str) -> bool:
    if not length:
        return False
    tokens = _norm(name).split()
    return bool(tokens) and tokens[-1] == length


def _should_process_description(
    text: str,
    skip_rules: SkipDescriptionRules,
) -> bool:
    return is_holland_product_description(text, extra=skip_rules)


@dataclass(frozen=True)
class MailFileLayout:
    """Колонки почтового Excel (0-based)."""

    desc_col: int
    length_col: int
    header_rows: frozenset[int]


def detect_mail_layout(sheet: object) -> MailFileLayout:
    """Ищет строки с заголовком Description и колонку S1 рядом."""
    import xlrd

    header_rows: list[int] = []
    desc_col: int | None = None
    length_col: int | None = None
    for row in range(min(40, sheet.nrows)):
        found_desc: int | None = None
        found_s1: int | None = None
        for col in range(sheet.ncols):
            header = _norm(sheet.cell_value(row, col)).casefold()
            if header == "description":
                found_desc = col
            elif header in {"s1", "s 1"}:
                found_s1 = col
        if found_desc is not None:
            header_rows.append(row)
            if desc_col is None:
                desc_col = found_desc
            if found_s1 is not None and length_col is None:
                length_col = found_s1
    if desc_col is None:
        raise RuntimeError("Не найден заголовок Description")
    if length_col is None:
        # Папка 1: L=Description, T=S1; папка 2: F=Description, K=S1
        if desc_col == _col_index("L"):
            length_col = _col_index("T")
        elif desc_col == _col_index("F"):
            length_col = _col_index("K")
        else:
            raise RuntimeError(
                f"Не найден заголовок S1 рядом с Description (колонка {_col_letter(desc_col)})"
            )
    if not header_rows:
        raise RuntimeError("Не найдена строка заголовка Description")
    return MailFileLayout(
        desc_col=desc_col,
        length_col=length_col,
        header_rows=frozenset(header_rows),
    )


@dataclass
class MailProcessResult:
    path: Path
    translated: int
    roses_updated: int
    total_rows: int
    missing: list[str]


def _cell_str(sheet: object, row: int, col: int) -> str:
    import xlrd

    if col >= sheet.ncols:
        return ""
    val = sheet.cell_value(row, col)
    ctype = sheet.cell_type(row, col)
    if ctype == xlrd.XL_CELL_NUMBER and isinstance(val, float) and val == int(val):
        return str(int(val))
    return _norm(val)


def _plan_updates_xls(
    sheet: object,
    layout: MailFileLayout,
    dictionary: dict[str, str],
    *,
    skip_rules: SkipDescriptionRules = SKIP_DESCRIPTION_RULES_EMPTY,
) -> tuple[dict[int, str], list[str], int, int, int]:
    """row → новое Description; missing texts; translated; roses; total."""
    updates: dict[int, str] = {}
    missing: list[str] = []
    translated = 0
    roses = 0
    total = 0
    for row in range(sheet.nrows):
        if row in layout.header_rows:
            continue
        raw = _cell_str(sheet, row, layout.desc_col)
        if not _should_process_description(raw, skip_rules):
            continue
        total += 1
        new_text, exact, any_hit = translate_description(dictionary, raw)
        if exact or any_hit:
            translated += 1
        if not exact:
            missing.append(raw)

        length = length_without_cm(sheet.cell_value(row, layout.length_col))
        # Роза — по исходному (R …) или уже переведённому (Роза …).
        rose = is_rose_description(raw) or is_rose_description(new_text)
        if rose and length and not _already_has_rostovka(new_text, length):
            new_text = f"{new_text} {length}".strip()
            roses += 1
        elif rose and length and _already_has_rostovka(new_text, length):
            pass

        if new_text != raw:
            updates[row] = new_text
    return updates, missing, translated, roses, total


def _apply_xls_updates(path: Path, desc_col: int, updates: dict[int, str]) -> None:
    from xlutils.copy import copy as xl_copy

    raw = path.read_bytes()
    rb = _open_xls_workbook(raw)
    wb = xl_copy(rb)
    _copy_workbook_palette(rb, wb)
    ws = wb.get_sheet(0)
    sheet_in = rb.sheet_by_index(0)
    # Сохраняем ширину колонки Description.
    try:
        ws.col(desc_col).width = sheet_in.computed_column_width(desc_col)
    except Exception:
        pass
    for row, value in updates.items():
        ws.write(row, desc_col, value)
    wb.save(str(path))


def _plan_updates_xlsx(
    path: Path,
    dictionary: dict[str, str],
    *,
    skip_rules: SkipDescriptionRules = SKIP_DESCRIPTION_RULES_EMPTY,
) -> tuple[str, str, dict[str, str | None], list[str], int, int, int]:
    """desc_col, length_col letters, updates A1→value, missing, translated, roses, total."""
    grid = read_xlsx_grid(path)
    rows = grid_by_row(grid)
    header_rows: list[int] = []
    desc_col: str | None = None
    length_col: str | None = None
    for row_n, cells in rows.items():
        found_desc = found_s1 = None
        for col, val in cells.items():
            h = _norm(val).casefold()
            if h == "description":
                found_desc = col
            elif h in {"s1", "s 1"}:
                found_s1 = col
        if found_desc:
            header_rows.append(row_n)
            desc_col = desc_col or found_desc
            if found_s1 and length_col is None:
                length_col = found_s1
    if desc_col is None:
        raise RuntimeError(f"В {path.name} нет заголовка Description")
    if length_col is None:
        if desc_col == "L":
            length_col = "T"
        elif desc_col == "F":
            length_col = "K"
        else:
            raise RuntimeError(f"В {path.name} нет колонки S1")

    updates: dict[str, str | None] = {}
    missing: list[str] = []
    translated = roses = total = 0
    max_row = max(rows) if rows else 1
    header_set = set(header_rows)
    for row_n in range(1, max_row + 1):
        if row_n in header_set:
            continue
        row = rows.get(row_n, {})
        raw = _norm(row.get(desc_col, ""))
        if not _should_process_description(raw, skip_rules):
            continue
        total += 1
        new_text, exact, any_hit = translate_description(dictionary, raw)
        if exact or any_hit:
            translated += 1
        if not exact:
            missing.append(raw)
        length = length_without_cm(row.get(length_col, ""))
        rose = is_rose_description(raw) or is_rose_description(new_text)
        if rose and length and not _already_has_rostovka(new_text, length):
            new_text = f"{new_text} {length}".strip()
            roses += 1
        if new_text != raw:
            updates[f"{desc_col}{row_n}"] = new_text
    return desc_col, length_col, updates, missing, translated, roses, total


def process_mail_workbook(
    path: Path,
    dictionary_path: Path,
    *,
    append_missing_to_dictionary: bool = True,
    skip_rules: SkipDescriptionRules = SKIP_DESCRIPTION_RULES_EMPTY,
    log: LogFn | None = None,
) -> MailProcessResult:
    """Переводит Description in-place и добавляет ростовку всем розам."""
    _lg = log or _default_log
    path = path.resolve()
    dictionary = load_description_dictionary(dictionary_path)
    if not dictionary:
        raise RuntimeError(f"Словарь пуст: {dictionary_path}")

    suffix = path.suffix.lower()
    if suffix == ".xls":
        raw = path.read_bytes()
        # Иногда .xls — на самом деле zip/xlsx.
        if raw[:2] == b"PK":
            suffix = ".xlsx"
        else:
            rb = _open_xls_workbook(raw)
            sheet = rb.sheet_by_index(0)
            layout = detect_mail_layout(sheet)
            _lg(
                f"{path.name}: Description={_col_letter(layout.desc_col)}, "
                f"S1={_col_letter(layout.length_col)}"
            )
            updates, missing, translated, roses, total = _plan_updates_xls(
                sheet, layout, dictionary, skip_rules=skip_rules
            )
            if append_missing_to_dictionary and missing:
                try:
                    append_missing_descriptions(
                        dictionary_path,
                        missing,
                        dictionary=dictionary,
                        extra=skip_rules,
                        log=_lg,
                    )
                except Exception as e:
                    _lg(f"Словарь: не удалось дописать — {e}")
            if updates:
                _apply_xls_updates(path, layout.desc_col, updates)
            _lg(
                f"{path.name}: строк {total}, переведено {translated}, "
                f"ростовка у роз {roses}, изменено ячеек {len(updates)}"
            )
            return MailProcessResult(path, translated, roses, total, missing)

    if suffix in {".xlsx", ".xlsm"}:
        desc_col, length_col, updates, missing, translated, roses, total = (
            _plan_updates_xlsx(path, dictionary, skip_rules=skip_rules)
        )
        _lg(f"{path.name}: Description={desc_col}, S1={length_col}")
        if append_missing_to_dictionary and missing:
            try:
                append_missing_descriptions(
                    dictionary_path,
                    missing,
                    dictionary=dictionary,
                    extra=skip_rules,
                    log=_lg,
                )
            except Exception as e:
                _lg(f"Словарь: не удалось дописать — {e}")
        if updates:
            if sys.platform == "win32":
                try:
                    _apply_xlsx_via_xlwings(path, updates, log=_lg)
                except Exception as e:
                    _lg(f"xlwings: {e} — XML-патч…")
                    patch_xlsx_cell_values(path, updates)
            else:
                patch_xlsx_cell_values(path, updates)
        _lg(
            f"{path.name}: строк {total}, переведено {translated}, "
            f"ростовка у роз {roses}, изменено ячеек {len(updates)}"
        )
        return MailProcessResult(path, translated, roses, total, missing)

    raise ValueError(f"Неподдерживаемый формат: {path.suffix}")


def _apply_xlsx_via_xlwings(
    path: Path,
    updates: dict[str, str | None],
    *,
    log: LogFn,
) -> None:
    import re as _re

    import xlwings as xw

    by_col: dict[str, list[tuple[int, str | None]]] = {}
    for ref, value in updates.items():
        m = _re.match(r"^([A-Z]+)(\d+)$", ref)
        if not m:
            continue
        by_col.setdefault(m.group(1), []).append((int(m.group(2)), value))

    app = wb = None
    try:
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        app.screen_updating = False
        wb = app.books.open(str(path), update_links=False)
        ws = wb.sheets[0]
        for col, pairs in by_col.items():
            pairs.sort(key=lambda x: x[0])
            first, last = pairs[0][0], pairs[-1][0]
            row_map = dict(pairs)
            values = [[row_map.get(r)] for r in range(first, last + 1)]
            ws.range(f"{col}{first}").resize(len(values), 1).value = values
        wb.save()
        log("Сохранено через Excel (xlwings).")
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


def process_mail_folders(
    short_dir: Path,
    long_dir: Path,
    dictionary_path: Path,
    *,
    append_missing_to_dictionary: bool = True,
    skip_rules: SkipDescriptionRules = SKIP_DESCRIPTION_RULES_EMPTY,
    log: LogFn | None = None,
) -> tuple[MailProcessResult, MailProcessResult]:
    """Обрабатывает по одному файлу в папке 1 и папке 2."""
    _lg = log or _default_log
    f1 = find_single_mail_workbook(short_dir, log=_lg)
    f2 = find_single_mail_workbook(long_dir, log=_lg)
    _lg(f"Папка 1: {f1.name}")
    r1 = process_mail_workbook(
        f1,
        dictionary_path,
        append_missing_to_dictionary=append_missing_to_dictionary,
        skip_rules=skip_rules,
        log=_lg,
    )
    _lg(f"Папка 2: {f2.name}")
    r2 = process_mail_workbook(
        f2,
        dictionary_path,
        append_missing_to_dictionary=append_missing_to_dictionary,
        skip_rules=skip_rules,
        log=_lg,
    )
    return r1, r2
