from __future__ import annotations

import re
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

_NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
_COL_RE = re.compile(r"^([A-Z]+)(\d+)$")


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
            if v is None or v.text is None:
                continue
            if cell.get("t") == "s":
                grid[ref] = shared[int(v.text)]
            else:
                grid[ref] = v.text
        return grid


def grid_by_row(grid: dict[str, str]) -> dict[int, dict[str, str]]:
    rows: dict[int, dict[str, str]] = {}
    for ref, val in grid.items():
        m = _COL_RE.match(ref)
        if not m:
            continue
        col, row = m.group(1), int(m.group(2))
        rows.setdefault(row, {})[col] = val
    return rows
