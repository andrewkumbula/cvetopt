from __future__ import annotations

import io
import os
import re
import stat
import sys
import time
import zipfile
from pathlib import Path

_COL_REF_RE = re.compile(r"^([A-Z]+)(\d+)$")


def _escape_xml_text(text: str) -> str:
    return (
        text.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


def _inline_str_inner(text: str) -> bytes:
    t = _escape_xml_text(text)
    if text != text.strip() or " " in text:
        return f'<is><t xml:space="preserve">{t}</t></is>'.encode("utf-8")
    return f"<is><t>{t}</t></is>".encode("utf-8")


def _col_letter_to_index(col: str) -> int:
    n = 0
    for ch in col.upper():
        n = n * 26 + (ord(ch) - ord("A") + 1)
    return n - 1


def _writable_win_errors(exc: BaseException) -> bool:
    if isinstance(exc, PermissionError):
        return True
    if isinstance(exc, OSError) and getattr(exc, "winerror", None) in (5, 32):
        return True
    return False


def _write_bytes_with_retry(path: Path, data: bytes, *, attempts: int = 12) -> None:
    last_err: BaseException | None = None
    for i in range(attempts):
        try:
            if path.exists():
                try:
                    os.chmod(path, stat.S_IWRITE | stat.S_IREAD)
                except OSError:
                    pass
            with path.open("wb") as fh:
                fh.write(data)
            return
        except (PermissionError, OSError) as e:
            last_err = e
            if not _writable_win_errors(e):
                raise
            time.sleep(0.35 * (i + 1))
    hint = (
        f"Не удалось записать {path}. "
        "Закройте этот файл в Excel и снова нажмите «Перевод»."
    )
    if sys.platform == "win32":
        hint += " Если файл только что создан auto1 — подождите 2–3 сек."
    raise PermissionError(hint) from last_err


def _first_sheet_path(names: list[str]) -> str | None:
    return next(
        (n for n in names if n.startswith("xl/worksheets/sheet") and n.endswith(".xml")),
        None,
    )


def _cell_pattern(ref: str) -> re.Pattern[bytes]:
    ref_b = ref.encode("ascii")
    return re.compile(
        rb'(<c r="' + re.escape(ref_b) + rb'"[^>]*>)(.*?)(</c>)',
        re.DOTALL,
    )


def _row_pattern(row_num: int) -> re.Pattern[bytes]:
    return re.compile(
        rb'(<row r="' + str(row_num).encode() + rb'"[^>]*>)(.*?)(</row>)',
        re.DOTALL,
    )


def _style_from_row(sheet: bytes, row_num: int, col_letter: str) -> bytes:
    """Берёт s=\"…\" с ячейки слева от целевой колонки в той же строке."""
    row_m = _row_pattern(row_num).search(sheet)
    if not row_m:
        return b""
    row_xml = row_m.group(0)
    target_idx = _col_letter_to_index(col_letter)
    style = b""
    for cell_m in re.finditer(rb'<c r="([A-Z]+)(\d+)"([^>]*)>', row_xml):
        col = cell_m.group(1).decode()
        if _col_letter_to_index(col) < target_idx:
            attrs = cell_m.group(3)
            sm = re.search(rb'\ss="(\d+)"', attrs)
            if sm:
                style = sm.group(0)
    return style


def _open_tag_for(ref: str, style: bytes, *, inline: bool) -> bytes:
    base = b'<c r="' + ref.encode("ascii") + b'"'
    if style:
        base += style
    if inline:
        base += b' t="inlineStr"'
    return base + b">"


def _patch_one_cell(sheet: bytes, ref: str, value: str | None) -> bytes:
    m = _COL_REF_RE.match(ref)
    if not m:
        return sheet
    col_l, row_n = m.group(1), int(m.group(2))
    pat = _cell_pattern(ref)
    style = _style_from_row(sheet, row_n, col_l)

    existing = pat.search(sheet)
    if value is None or not str(value).strip():
        if not existing:
            return sheet
        open_tag = existing.group(1)
        cleared = re.sub(rb'\s+t="[^"]*"', b"", open_tag) + b">"
        return pat.sub(cleared + b"</c>", sheet, count=1)

    inner = _inline_str_inner(str(value))
    if existing:
        open_tag = existing.group(1)
        if b't="inlineStr"' not in open_tag:
            open_tag = open_tag[:-1] + b' t="inlineStr">'
        return pat.sub(open_tag + inner + b"</c>", sheet, count=1)

    row_pat = _row_pattern(row_n)
    row_m = row_pat.search(sheet)
    if not row_m:
        return sheet
    new_cell = _open_tag_for(ref, style, inline=True) + inner + b"</c>"
    row_open, row_body, row_close = row_m.group(1), row_m.group(2), row_m.group(3)
    new_idx = _col_letter_to_index(col_l)
    insert_at = len(row_body)
    for cell_m in re.finditer(rb'<c r="([A-Z]+)\d+"', row_body):
        col = cell_m.group(1).decode()
        if _col_letter_to_index(col) > new_idx:
            insert_at = cell_m.start()
            break
    new_body = row_body[:insert_at] + new_cell + row_body[insert_at:]
    new_row = row_open + new_body + row_close
    return sheet[: row_m.start()] + new_row + sheet[row_m.end() :]


def patch_xlsx_cell_values(path: Path, updates: dict[str, str | None]) -> int:
    """
    Патчит только фрагменты <c> в sheet1.xml (без ElementTree),
    чтобы не ломать namespace, hyperlinks и чекбоксы.
    """
    if not updates:
        return 0
    path = path.resolve()
    changed = 0
    with zipfile.ZipFile(path, "r") as zf:
        sheet_name = _first_sheet_path(zf.namelist())
        if not sheet_name:
            raise RuntimeError(f"В {path.name} нет листа worksheet.")
        sheet_xml = zf.read(sheet_name)
        new_xml = sheet_xml
        for ref, value in updates.items():
            patched = _patch_one_cell(new_xml, ref, value)
            if patched != new_xml:
                changed += 1
                new_xml = patched
        if changed == 0:
            return 0
        buffer = io.BytesIO()
        with zipfile.ZipFile(buffer, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            for name in zf.namelist():
                part = new_xml if name == sheet_name else zf.read(name)
                zout.writestr(zf.getinfo(name), part)
        _write_bytes_with_retry(path, buffer.getvalue())
    return changed
