from __future__ import annotations

from pathlib import Path

import xlrd

from cvetopt.invoice.xlsx_read import _COL_RE, read_xlsx_grid

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
