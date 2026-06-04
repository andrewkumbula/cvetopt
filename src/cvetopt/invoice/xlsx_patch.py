from __future__ import annotations

import re
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

_NS_URI = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_NS = {"m": _NS_URI}
_COL_RE = re.compile(r"^([A-Z]+)(\d+)$")


def _col_letter_to_index(col: str) -> int:
    n = 0
    for ch in col.upper():
        n = n * 26 + (ord(ch) - ord("A") + 1)
    return n - 1


def _cell_text(cell: ET.Element, shared: list[str]) -> str:
    v = cell.find("m:v", _NS)
    if v is not None and v.text is not None:
        if cell.get("t") == "s":
            return shared[int(v.text)]
        return v.text
    is_node = cell.find("m:is", _NS)
    if is_node is not None:
        return "".join((n.text or "") for n in is_node.iter())
    return ""


def _clear_cell(cell: ET.Element) -> None:
    for tag in ("v", "is", "f"):
        node = cell.find(f"m:{tag}", _NS)
        if node is not None:
            cell.remove(node)
    if cell.get("t") is not None:
        del cell.attrib["t"]


def _set_inline_str(cell: ET.Element, text: str, *, style: str | None) -> None:
    _clear_cell(cell)
    if style:
        cell.set("s", style)
    cell.set("t", "inlineStr")
    is_el = ET.SubElement(cell, f"{{{_NS_URI}}}is")
    t_el = ET.SubElement(is_el, f"{{{_NS_URI}}}t")
    if text.startswith(" ") or text.endswith(" "):
        t_el.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t_el.text = text


def _find_row(sheet_data: ET.Element, row_num: int) -> ET.Element | None:
    for row in sheet_data.findall("m:row", _NS):
        if row.get("r") == str(row_num):
            return row
    return None


def _ensure_row(sheet_data: ET.Element, row_num: int) -> ET.Element:
    row = _find_row(sheet_data, row_num)
    if row is not None:
        return row
    row = ET.Element(f"{{{_NS_URI}}}row", {"r": str(row_num)})
    inserted = False
    for idx, existing in enumerate(sheet_data.findall("m:row", _NS)):
        er = int(existing.get("r", "0"))
        if er > row_num:
            sheet_data.insert(idx, row)
            inserted = True
            break
    if not inserted:
        sheet_data.append(row)
    return row


def _find_cell(row: ET.Element, ref: str) -> ET.Element | None:
    for cell in row.findall("m:c", _NS):
        if cell.get("r") == ref:
            return cell
    return None


def _insert_cell_sorted(row: ET.Element, cell: ET.Element, ref: str) -> None:
    new_idx = _col_letter_to_index("".join(ch for ch in ref if ch.isalpha()))
    for idx, existing in enumerate(row.findall("m:c", _NS)):
        er = existing.get("r", "")
        m = _COL_RE.match(er)
        if not m:
            continue
        if _col_letter_to_index(m.group(1)) > new_idx:
            row.insert(idx, cell)
            return
    row.append(cell)


def _first_sheet_path(names: list[str]) -> str | None:
    return next(
        (n for n in names if n.startswith("xl/worksheets/sheet") and n.endswith(".xml")),
        None,
    )


def patch_xlsx_cell_values(path: Path, updates: dict[str, str | None]) -> int:
    """
    Меняет только ячейки в sheet XML, остальные части xlsx (чекбоксы, drawing) не трогает.
    updates: «E2» → текст или None (очистить).
    """
    if not updates:
        return 0
    path = path.resolve()
    changed = 0
    with zipfile.ZipFile(path, "r") as zf:
        sheet_name = _first_sheet_path(zf.namelist())
        if not sheet_name:
            raise RuntimeError(f"В {path.name} нет листа worksheet.")
        root = ET.fromstring(zf.read(sheet_name))
        shared: list[str] = []
        if "xl/sharedStrings.xml" in zf.namelist():
            sroot = ET.fromstring(zf.read("xl/sharedStrings.xml"))
            for si in sroot.findall(".//m:si", _NS):
                shared.append("".join((n.text or "") for n in si.iter()))

        sheet_data = root.find("m:sheetData", _NS)
        if sheet_data is None:
            raise RuntimeError(f"В {path.name} нет sheetData.")

        for ref, value in updates.items():
            m = _COL_RE.match(ref)
            if not m:
                continue
            col_l, row_n = m.group(1), int(m.group(2))
            row = _ensure_row(sheet_data, row_n)
            cell = _find_cell(row, ref)
            if cell is None:
                style = None
                left_ref = f"{col_l}{row_n}"
                # стиль с ячейки слева (Description), если есть
                for c in row.findall("m:c", _NS):
                    cr = c.get("r", "")
                    cm = _COL_RE.match(cr)
                    if cm and _col_letter_to_index(cm.group(1)) < _col_letter_to_index(col_l):
                        style = c.get("s")
                cell = ET.Element(f"{{{_NS_URI}}}c", {"r": ref})
                if style:
                    cell.set("s", style)
                _insert_cell_sorted(row, cell, ref)

            if value is None or str(value).strip() == "":
                if list(cell) or cell.get("t"):
                    _clear_cell(cell)
                    changed += 1
            else:
                prev = _cell_text(cell, shared)
                text = str(value)
                if prev != text:
                    _set_inline_str(cell, text, style=cell.get("s"))
                    changed += 1

        if changed == 0:
            return 0

        ET.register_namespace("", _NS_URI)
        out_xml = ET.tostring(root, encoding="utf-8", xml_declaration=True)
        tmp = path.with_suffix(path.suffix + ".patch.tmp")
        with zipfile.ZipFile(tmp, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            for name in zf.namelist():
                data = out_xml if name == sheet_name else zf.read(name)
                zinfo = zf.getinfo(name)
                zout.writestr(zinfo, data)
        tmp.replace(path)
    return changed
