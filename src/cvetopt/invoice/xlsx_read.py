from __future__ import annotations

import re
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

import xlrd

_NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
_COL_RE = re.compile(r"^([A-Z]+)(\d+)$")
_OLE_MAGIC = b"\xd0\xcf\x11\xe0"


def _col_index_to_letter(idx: int) -> str:
    n = idx + 1
    letters = ""
    while n:
        n, rem = divmod(n - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def _cell_to_text(value: object) -> str:
    if value is None or value == "":
        return ""
    if isinstance(value, float):
        if value == int(value):
            return str(int(value))
        return str(value)
    return " ".join(str(value).split()).strip()


def excel_file_kind(path: Path) -> str:
    """xlsx | xls | invalid"""
    try:
        head = path.read_bytes()[:4]
    except OSError:
        return "invalid"
    if head[:2] == b"PK":
        return "xlsx"
    if head == _OLE_MAGIC:
        return "xls"
    return "invalid"


def read_xls_grid(path: Path) -> dict[str, str]:
    """Первый лист .xls → «A1» → значение (как read_xlsx_grid)."""
    book = xlrd.open_workbook(str(path))
    sheet = book.sheet_by_index(0)
    grid: dict[str, str] = {}
    for row_idx in range(sheet.nrows):
        for col_idx in range(sheet.ncols):
            text = _cell_to_text(sheet.cell_value(row_idx, col_idx))
            if not text:
                continue
            grid[f"{_col_index_to_letter(col_idx)}{row_idx + 1}"] = text
    return grid


def read_xlsx_grid(path: Path) -> dict[str, str]:
    """Возвращает словарь «A1» → значение (строка или число как str)."""
    with zipfile.ZipFile(path) as zf:
        shared: list[str] = []
        if "xl/sharedStrings.xml" in zf.namelist():
            root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
            for si in root.findall(".//m:si", _NS):
                shared.append("".join((n.text or "") for n in si.iter()))

        sheet_name = next(
            (n for n in zf.namelist() if n.startswith("xl/worksheets/sheet") and n.endswith(".xml")),
            None,
        )
        if not sheet_name:
            return {}
        root = ET.fromstring(zf.read(sheet_name))
        grid: dict[str, str] = {}
        for cell in root.findall(".//m:sheetData//m:c", _NS):
            ref = cell.get("r")
            if not ref:
                continue
            v = cell.find("m:v", _NS)
            if v is not None and v.text is not None:
                if cell.get("t") == "s":
                    grid[ref] = shared[int(v.text)]
                else:
                    grid[ref] = v.text
                continue
            is_node = cell.find("m:is", _NS)
            if is_node is not None:
                grid[ref] = "".join((n.text or "") for n in is_node.iter())
        return grid


def read_excel_grid(path: Path) -> dict[str, str]:
    """Читает .xlsx или старый .xls (в т.ч. .xls под именем .xlsx)."""
    kind = excel_file_kind(path)
    if kind == "xlsx":
        return read_xlsx_grid(path)
    if kind == "xls":
        return read_xls_grid(path)
    raise ValueError(f"Не Excel или повреждённый файл: {path.name}")


def ensure_xlsx_workbook(path: Path) -> Path:
    """
    Если файл — старый .xls (даже с расширением .xlsx), пересохраняет как настоящий .xlsx.
    Нужно для openpyxl (миксы, сплит).
    """
    from openpyxl import Workbook

    path = path.resolve()
    if excel_file_kind(path) != "xls":
        return path
    grid = read_xls_grid(path)
    wb = Workbook()
    ws = wb.active
    for ref, val in grid.items():
        if _COL_RE.match(ref):
            ws[ref] = val
    wb.save(path)
    wb.close()
    return path


def grid_by_row(grid: dict[str, str]) -> dict[int, dict[str, str]]:
    rows: dict[int, dict[str, str]] = {}
    for ref, val in grid.items():
        m = _COL_RE.match(ref)
        if not m:
            continue
        col, row = m.group(1), int(m.group(2))
        rows.setdefault(row, {})[col] = val
    return rows
