from __future__ import annotations

import re
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

_NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
_COL_RE = re.compile(r"^([A-Z]+)(\d+)$")


def _normalize_header(value: object) -> str:
    return re.sub(r"\s+", " ", str(value).strip().lower())


def _find_price_total_columns_xlrd(sheet: object) -> tuple[int, int, int] | None:
    """Возвращает (строка заголовка, колонка price, колонка total) или None."""
    import xlrd

    for row in range(min(30, sheet.nrows)):
        price_col: int | None = None
        total_col: int | None = None
        for col in range(sheet.ncols):
            header = _normalize_header(sheet.cell_value(row, col))
            if header == "price":
                price_col = col
            elif header == "total":
                total_col = col
        if price_col is not None and total_col is not None:
            return row, price_col, total_col
    return None


def _preserve_col_width(rb_sheet: object, ws: object, col: int) -> None:
    """Фиксирует ширину столбца до записи ячеек (xlwt иначе ставит ~11.5 символов)."""
    ws.col(col).width = rb_sheet.computed_column_width(col)


def _cell_has_content(sheet: object, row: int, col: int) -> bool:
    import xlrd

    ctype = sheet.cell_type(row, col)
    if ctype in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK):
        return False
    return bool(str(sheet.cell_value(row, col)).strip())


def _clear_xls_cell(ws: object, sheet_in: object, row: int, col: int) -> None:
    """Пустая ячейка без текста — как BLANK в оригинале, не пустая строка."""
    ws.write(row, col, None)


def _open_xls_workbook(raw: bytes):
    import xlrd

    try:
        return xlrd.open_workbook(file_contents=raw, formatting_info=True)
    except Exception:
        return xlrd.open_workbook(file_contents=raw)


def _copy_workbook_palette(rb: object, wb: object) -> None:
    """
    xlwt по умолчанию подставляет свою палитру: серые заливки шапки становятся красными.
    Копируем RGB из исходной книги (индексы 8–63).
    """
    if not getattr(rb, "formatting_info", False):
        return
    for idx in range(8, 64):
        rgb = rb.colour_map.get(idx)
        if rgb and len(rgb) == 3:
            wb.set_colour_RGB(idx, rgb[0], rgb[1], rgb[2])


def _clear_xls_price_total(path: Path) -> bool:
    import xlrd
    from xlutils.copy import copy as xl_copy

    raw = path.read_bytes()
    rb = _open_xls_workbook(raw)
    info = _find_price_total_columns_xlrd(rb.sheet_by_index(0))
    if info is None:
        return False
    header_row, price_col, total_col = info
    sheet_in = rb.sheet_by_index(0)
    wb = xl_copy(rb)
    _copy_workbook_palette(rb, wb)
    ws = wb.get_sheet(0)
    ws._cell_overwrite_ok = True

    cols = (price_col, total_col)
    if rb.formatting_info:
        for col in cols:
            _preserve_col_width(sheet_in, ws, col)

    cleared = 0
    for row in range(header_row, sheet_in.nrows):
        for col in cols:
            if row == header_row or _cell_has_content(sheet_in, row, col):
                _clear_xls_cell(ws, sheet_in, row, col)
                cleared += 1

    if cleared == 0:
        return False

    with path.open("wb") as fh:
        wb.save(fh)
    return True


def _col_letter_to_index(col: str) -> int:
    n = 0
    for ch in col.upper():
        n = n * 26 + (ord(ch) - ord("A") + 1)
    return n - 1


def _index_to_col_letter(index: int) -> str:
    index += 1
    letters = ""
    while index:
        index, rem = divmod(index - 1, 26)
        letters = chr(ord("A") + rem) + letters
    return letters


def _clear_xlsx_cell(cell: ET.Element) -> bool:
    changed = False
    for tag in ("v", "is", "f"):
        node = cell.find(f"m:{tag}", _NS)
        if node is not None:
            cell.remove(node)
            changed = True
    if cell.get("t") is not None:
        del cell.attrib["t"]
        changed = True
    return changed


def _clear_xlsx_price_total(path: Path) -> bool:
    """xlsx (Office Open XML): очистка ячеек в столбцах price и total."""
    with zipfile.ZipFile(path, "r") as zf:
        sheet_name = next(
            (
                n
                for n in zf.namelist()
                if n.startswith("xl/worksheets/sheet") and n.endswith(".xml")
            ),
            None,
        )
        if not sheet_name:
            return False
        root = ET.fromstring(zf.read(sheet_name))
        shared: list[str] = []
        if "xl/sharedStrings.xml" in zf.namelist():
            sroot = ET.fromstring(zf.read("xl/sharedStrings.xml"))
            for si in sroot.findall(".//m:si", _NS):
                shared.append("".join((n.text or "") for n in si.iter()))

        grid: dict[tuple[int, int], str] = {}
        for cell in root.findall(".//m:sheetData//m:c", _NS):
            ref = cell.get("r")
            if not ref:
                continue
            m = _COL_RE.match(ref)
            if not m:
                continue
            col_l, row_n = m.group(1), int(m.group(2))
            col_i = _col_letter_to_index(col_l)
            v = cell.find("m:v", _NS)
            is_node = cell.find("m:is", _NS)
            if v is not None and v.text is not None:
                text = shared[int(v.text)] if cell.get("t") == "s" else v.text
            elif is_node is not None:
                text = "".join((n.text or "") for n in is_node.iter())
            else:
                continue
            grid[(row_n, col_i)] = text

        header_row = price_col = total_col = None
        for (row_n, col_i), text in grid.items():
            h = _normalize_header(text)
            if h == "price":
                price_col = col_i
                if header_row is None or row_n < header_row:
                    header_row = row_n
            elif h == "total":
                total_col = col_i
                if header_row is None or row_n < header_row:
                    header_row = row_n
        if header_row is None or price_col is None or total_col is None:
            return False

        price_letter = _index_to_col_letter(price_col)
        total_letter = _index_to_col_letter(total_col)
        cleared = 0
        for cell in root.findall(".//m:sheetData//m:c", _NS):
            ref = cell.get("r")
            if not ref:
                continue
            m = _COL_RE.match(ref)
            if not m:
                continue
            col_l, row_n = m.group(1), int(m.group(2))
            if row_n < header_row:
                continue
            if col_l not in (price_letter, total_letter):
                continue
            if _clear_xlsx_cell(cell):
                cleared += 1

        if cleared == 0:
            return False

        out_xml = ET.tostring(root, encoding="utf-8", xml_declaration=True)
        tmp = path.with_suffix(path.suffix + ".tmp")
        with zipfile.ZipFile(tmp, "w") as zout:
            for name in zf.namelist():
                data = out_xml if name == sheet_name else zf.read(name)
                zout.writestr(name, data)
        tmp.replace(path)
    return True


def clear_price_total_columns(path: Path) -> bool:
    """
    Очищает столбцы price и total на первом листе (включая заголовки).
    Возвращает True, если заголовки найдены и файл изменён.
    """
    raw = path.read_bytes()
    if len(raw) >= 2 and raw[:2] == b"PK":
        return _clear_xlsx_price_total(path)
    return _clear_xls_price_total(path)
