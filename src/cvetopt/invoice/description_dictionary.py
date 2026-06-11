from __future__ import annotations

import shutil
import sys
from collections.abc import Callable, Iterable
from pathlib import Path

import xlrd
import xlwt

from cvetopt.invoice.xlsx_read import _COL_RE, read_xlsx_grid

LogFn = Callable[[str], None]

# Словарь.xls: B — английское название, C — перевод.
_DICT_COL_EN = 1  # B, 0-based
_DICT_COL_RU = 2  # C, 0-based
_DICT_COL_EN_LETTER = "B"
_DICT_COL_RU_LETTER = "C"

_HEADER_HINTS = frozenset(
    {
        "description",
        "название",
        "наименование",
        "англ",
        "english",
        "исходный",
        "исход",
        "перевод",
        "рус",
        "русский",
        "original",
        "translation",
    }
)


def _norm_key(text: str) -> str:
    return " ".join(str(text or "").split()).strip()


def _is_header_row(a: str, b: str) -> bool:
    al = _norm_key(a).lower()
    bl = _norm_key(b).lower()
    if not al and not bl:
        return True
    return al in _HEADER_HINTS or bl in _HEADER_HINTS or "описан" in al or "описан" in bl


def _load_from_xls(path: Path) -> dict[str, str]:
    book = xlrd.open_workbook(str(path))
    sheet = book.sheet_by_index(0)
    out: dict[str, str] = {}
    for r in range(sheet.nrows):
        if sheet.ncols <= _DICT_COL_RU:
            continue
        src = sheet.cell_value(r, _DICT_COL_EN)
        dst = sheet.cell_value(r, _DICT_COL_RU)
        if isinstance(src, float) and src == int(src):
            src = str(int(src))
        if isinstance(dst, float) and dst == int(dst):
            dst = str(int(dst))
        key = _norm_key(str(src))
        val = _norm_key(str(dst))
        if not key or not val:
            continue
        if r == 0 and _is_header_row(key, val):
            continue
        out[key] = val
    return out


def _load_from_xlsx(path: Path) -> dict[str, str]:
    grid = read_xlsx_grid(path)
    if not grid:
        return {}
    rows: dict[int, dict[str, str]] = {}
    for ref, val in grid.items():
        m = _COL_RE.match(ref)
        if not m:
            continue
        col, row_no = m.group(1), int(m.group(2))
        if col not in (_DICT_COL_EN_LETTER, _DICT_COL_RU_LETTER):
            continue
        rows.setdefault(row_no, {})[col] = str(val)
    out: dict[str, str] = {}
    for row_no in sorted(rows):
        cells = rows[row_no]
        if _DICT_COL_EN_LETTER not in cells or _DICT_COL_RU_LETTER not in cells:
            continue
        key = _norm_key(cells[_DICT_COL_EN_LETTER])
        val = _norm_key(cells[_DICT_COL_RU_LETTER])
        if not key or not val:
            continue
        if row_no == 1 and _is_header_row(key, val):
            continue
        out[key] = val
    return out


def load_description_dictionary(path: Path) -> dict[str, str]:
    """Словарь: колонка B (англ.) → C (перевод), первый лист."""
    p = path.resolve()
    if not p.is_file():
        raise FileNotFoundError(f"Словарь не найден: {p}")
    if p.stat().st_size < 32:
        raise RuntimeError(f"Словарь пустой или повреждён: {p}")

    if p.suffix.lower() == ".xlsx":
        return _load_from_xlsx(p)

    if p.suffix.lower() == ".xls":
        try:
            return _load_from_xls(p)
        except xlrd.XLRDError:
            # Иногда .xls на диске — на самом деле xlsx/zip.
            with p.open("rb") as fh:
                if fh.read(2) == b"PK":
                    return _load_from_xlsx(p)
            raise

    raise ValueError(f"Неподдерживаемый формат словаря: {p.suffix}")


def lookup_translation(dictionary: dict[str, str], description: str) -> str | None:
    key = _norm_key(description)
    if not key:
        return None
    if key in dictionary:
        return dictionary[key]
    folded = key.casefold()
    for src, dst in dictionary.items():
        if src.casefold() == folded:
            return dst
    return None


def _english_keys_from_xls(path: Path) -> set[str]:
    book = xlrd.open_workbook(str(path))
    sheet = book.sheet_by_index(0)
    keys: set[str] = set()
    for r in range(sheet.nrows):
        if sheet.ncols <= _DICT_COL_EN:
            continue
        src = sheet.cell_value(r, _DICT_COL_EN)
        if isinstance(src, float) and src == int(src):
            src = str(int(src))
        key = _norm_key(str(src))
        if not key:
            continue
        dst = ""
        if sheet.ncols > _DICT_COL_RU:
            raw_dst = sheet.cell_value(r, _DICT_COL_RU)
            if isinstance(raw_dst, float) and raw_dst == int(raw_dst):
                raw_dst = str(int(raw_dst))
            dst = _norm_key(str(raw_dst))
        if r == 0 and _is_header_row(key, dst):
            continue
        keys.add(key.casefold())
    return keys


def _english_keys_from_xlsx(path: Path) -> set[str]:
    grid = read_xlsx_grid(path)
    if not grid:
        return set()
    rows: dict[int, dict[str, str]] = {}
    for ref, val in grid.items():
        m = _COL_RE.match(ref)
        if not m:
            continue
        col, row_no = m.group(1), int(m.group(2))
        if col not in (_DICT_COL_EN_LETTER, _DICT_COL_RU_LETTER):
            continue
        rows.setdefault(row_no, {})[col] = str(val)
    keys: set[str] = set()
    for row_no in sorted(rows):
        cells = rows[row_no]
        if _DICT_COL_EN_LETTER not in cells:
            continue
        key = _norm_key(cells[_DICT_COL_EN_LETTER])
        if not key:
            continue
        val = _norm_key(cells.get(_DICT_COL_RU_LETTER, ""))
        if row_no == 1 and _is_header_row(key, val):
            continue
        keys.add(key.casefold())
    return keys


def load_dictionary_english_keys(path: Path) -> set[str]:
    p = path.resolve()
    if p.suffix.lower() == ".xlsx":
        return _english_keys_from_xlsx(p)
    if p.suffix.lower() == ".xls":
        try:
            return _english_keys_from_xls(p)
        except xlrd.XLRDError:
            with p.open("rb") as fh:
                if fh.read(2) == b"PK":
                    return _english_keys_from_xlsx(p)
            raise
    raise ValueError(f"Неподдерживаемый формат словаря: {p.suffix}")


def _unique_missing_descriptions(
    descriptions: Iterable[str],
    dictionary: dict[str, str],
    existing_keys: set[str],
) -> list[str]:
    translated_folded = {k.casefold() for k in dictionary}
    out: list[str] = []
    seen: set[str] = set()
    for raw in descriptions:
        key = _norm_key(raw)
        if not key:
            continue
        folded = key.casefold()
        if folded in seen or folded in translated_folded or folded in existing_keys:
            continue
        seen.add(folded)
        out.append(key)
    return sorted(out, key=str.casefold)


def _dictionary_next_row(path: Path) -> int:
    suffix = path.suffix.lower()
    if suffix == ".xlsx":
        grid = read_xlsx_grid(path)
        max_row = 1
        for ref in grid:
            m = _COL_RE.match(ref)
            if m and m.group(1) in (_DICT_COL_EN_LETTER, _DICT_COL_RU_LETTER):
                max_row = max(max_row, int(m.group(2)))
        return max_row + 1
    book = xlrd.open_workbook(str(path))
    return book.sheet_by_index(0).nrows + 1


def _append_rows_xlwings(path: Path, rows: list[tuple[str, str]]) -> None:
    import xlwings as xw

    start = _dictionary_next_row(path)
    app: object | None = None
    wb: object | None = None
    try:
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        wb = app.books.open(str(path), update_links=False)
        sht = wb.sheets[0]
        values = [[en, ru] for en, ru in rows]
        sht.range(f"B{start}").resize(len(values), 2).value = values
        wb.save()
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


def _append_rows_xlwt(path: Path, rows: list[tuple[str, str]]) -> None:
    rb = xlrd.open_workbook(str(path))
    sheet_in = rb.sheet_by_index(0)
    wb_out = xlwt.Workbook()
    ws_out = wb_out.add_sheet(rb.sheet_names()[0] or "Sheet1")
    for r in range(sheet_in.nrows):
        for c in range(sheet_in.ncols):
            ws_out.write(r, c, sheet_in.cell_value(r, c))
    start = sheet_in.nrows
    for offset, (en, ru) in enumerate(rows):
        row = start + offset
        ws_out.write(row, _DICT_COL_EN, en)
        ws_out.write(row, _DICT_COL_RU, ru)
    tmp = path.with_suffix(path.suffix + ".dict_append.tmp")
    wb_out.save(str(tmp))
    tmp.replace(path)


def append_missing_descriptions(
    dictionary_path: Path,
    descriptions: Iterable[str],
    *,
    dictionary: dict[str, str] | None = None,
    log: LogFn | None = None,
) -> int:
    """
    Дописывает в словарь (B — англ., C — пусто) Description без перевода.
    Уже существующие ключи в колонке B не дублируются.
    """
    path = dictionary_path.resolve()
    if not path.is_file():
        raise FileNotFoundError(f"Словарь не найден: {path}")

    dict_map = dictionary if dictionary is not None else load_description_dictionary(path)
    existing_keys = load_dictionary_english_keys(path)
    to_add = _unique_missing_descriptions(descriptions, dict_map, existing_keys)
    if not to_add:
        return 0

    rows = [(text, "") for text in to_add]
    backup = path.with_suffix(path.suffix + ".bak")
    shutil.copy2(path, backup)

    try:
        if path.suffix.lower() == ".xls":
            if sys.platform == "win32":
                try:
                    _append_rows_xlwings(path, rows)
                except Exception:
                    _append_rows_xlwt(path, rows)
            else:
                _append_rows_xlwt(path, rows)
        elif path.suffix.lower() == ".xlsx":
            if sys.platform != "win32":
                raise RuntimeError(
                    "Дополнение .xlsx словаря без Excel доступно только на Windows."
                )
            _append_rows_xlwings(path, rows)
        else:
            raise ValueError(f"Неподдерживаемый формат словаря: {path.suffix}")
    except Exception:
        shutil.copy2(backup, path)
        raise

    if log is not None:
        log(
            f"Словарь: добавлено {len(to_add)} без перевода "
            f"(колонка B, C пустая) → {path.name}"
        )
        preview = ", ".join(to_add[:5])
        if len(to_add) > 5:
            preview += f" … (+{len(to_add) - 5})"
        log(f"Словарь: новые: {preview}")
    return len(to_add)
