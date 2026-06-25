from __future__ import annotations

import re
import shutil
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import xlrd
from xlrd.sheet import Sheet
from xlutils.copy import copy as xlutils_copy
from xlwt import Worksheet

from cvetopt.core.settings import BalanceAutoBlockConfig, BalanceAutoClearRange, BalanceAutoConfig

_XL_FORMULA = 2


@dataclass(frozen=True)
class FlightFillRow:
    platform: str  # "ecuador" | "colombia"
    weight: float
    awb: str
    price: float


_CELL_REF = re.compile(r"^([A-Za-z]+)(\d+)$")


def excel_col_letters_to_index(letters: str) -> int:
    n = 0
    for ch in letters.strip().upper():
        if not ("A" <= ch <= "Z"):
            raise ValueError(f"Bad column letters: {letters!r}")
        n = n * 26 + (ord(ch) - ord("A") + 1)
    return n - 1


def parse_excel_cell(ref: str) -> tuple[int, int]:
    """Excel-адрес вида K15 → (row_0based, col_0based)."""
    m = _CELL_REF.match(ref.strip())
    if not m:
        raise ValueError(f"Bad cell ref: {ref!r}")
    col = excel_col_letters_to_index(m.group(1))
    row = int(m.group(2)) - 1
    return row, col


def inclusive_rect_cells(
    top_left: str,
    bottom_right: str,
) -> tuple[int, int, int, int]:
    r0, c0 = parse_excel_cell(top_left)
    r1, c1 = parse_excel_cell(bottom_right)
    rt = min(r0, r1)
    rb = max(r0, r1)
    cl = min(c0, c1)
    cr = max(c0, c1)
    return rt, rb, cl, cr


def _is_formula(rb_sheet: Sheet, r: int, c: int) -> bool:
    try:
        return rb_sheet.cell(r, c).ctype == _XL_FORMULA
    except IndexError:
        return False


def clear_ranges_preserve_formulas(
    rb_sheet: Sheet,
    ws: Worksheet,
    ranges: list[BalanceAutoClearRange],
) -> int:
    """Очищает только не-формульные ячейки в заданных прямоугольниках. Возвращает число очищенных."""
    cleared = 0
    for rect in ranges:
        rt, rb, cl, cr = inclusive_rect_cells(rect.top_left, rect.bottom_right)
        for r in range(rt, rb + 1):
            for c in range(cl, cr + 1):
                if _is_formula(rb_sheet, r, c):
                    continue
                ws.write(r, c, "")
                cleared += 1
    return cleared


def _force_clear_block_data_cols_xlwt(ws: Worksheet, block: BalanceAutoBlockConfig) -> int:
    """В строках данных блока затирает указанные колонки, в т.ч. с формулами (xlwt)."""
    if not block.force_clear_cols:
        return 0
    r0 = block.first_data_row_excel - 1
    r1 = block.last_data_row_excel - 1
    n = 0
    for letters in block.force_clear_cols:
        c = excel_col_letters_to_index(letters.strip().upper())
        for r in range(r0, r1 + 1):
            ws.write(r, c, "")
            n += 1
    return n


def _force_clear_blocks_xlwt(ws: Worksheet, cfg: BalanceAutoConfig) -> int:
    return _force_clear_block_data_cols_xlwt(ws, cfg.ecuador) + _force_clear_block_data_cols_xlwt(
        ws,
        cfg.colombia,
    )


def _col_to_idx(block: BalanceAutoBlockConfig, which: str) -> int:
    if which == "w":
        return excel_col_letters_to_index(block.weight_col)
    if which == "a":
        return excel_col_letters_to_index(block.awb_col)
    if which == "p":
        return excel_col_letters_to_index(block.price_col)
    raise ValueError(which)


def _write_block_xlwt(
    ws: Worksheet,
    block: BalanceAutoBlockConfig,
    rows: list[FlightFillRow],
) -> tuple[int, int]:
    """Пишет строки в блок [first..last]. Возвращает (записано, пропущено_переполнение)."""
    r_first = block.first_data_row_excel - 1
    r_last = block.last_data_row_excel - 1
    cw = _col_to_idx(block, "w")
    ca = _col_to_idx(block, "a")
    cp = _col_to_idx(block, "p")
    cap = r_last - r_first + 1
    written = 0
    for i, fr in enumerate(rows[:cap]):
        r = r_first + i
        ws.write(r, cw, fr.weight)
        ws.write(r, ca, _format_awb_for_excel(fr.awb))
        ws.write(r, cp, fr.price)
        written += 1
    skipped = max(0, len(rows) - cap)
    return written, skipped


def _apply_balance_flights_xlwt(
    path: Path,
    cfg: BalanceAutoConfig,
    flights: list[FlightFillRow],
    notes: list[str],
) -> tuple[int, int, int, int]:
    ec = [f for f in flights if f.platform == "ecuador"]
    co = [f for f in flights if f.platform == "colombia"]

    rb = xlrd.open_workbook(str(path), formatting_info=True)
    try:
        sh = rb.sheet_by_name(cfg.sheet_name)
    except xlrd.XLRDError as e:
        raise ValueError(f"Лист {cfg.sheet_name!r} не найден: {e}") from e
    sheet_idx = sh.number
    wb = xlutils_copy(rb)
    ws = wb.get_sheet(sheet_idx)

    cleared = clear_ranges_preserve_formulas(sh, ws, cfg.clear_ranges)

    forced = _force_clear_blocks_xlwt(ws, cfg)
    if forced:
        notes.append(
            f"Дополнительно очищено {forced} ячеек в колонках force_clear_cols "
            f"(строки данных; формулы в этих ячейках сброшены, итоги пересчитаются).",
        )

    we, oe = _write_block_xlwt(ws, cfg.ecuador, ec)
    wc, oc = _write_block_xlwt(ws, cfg.colombia, co)
    overflow = oe + oc

    wb.save(str(path))
    return cleared, we, wc, overflow


def _flatten_formula_value(f: Any) -> str:
    """xlwings для одной ячейки даёт str, для диапазона — вложенные list/tuple."""
    while isinstance(f, (list, tuple)):
        if not f:
            return ""
        f = f[0]
    if f is None:
        return ""
    return str(f)


def _xlwings_cell_has_formula(sheet: Any, row_excel: int, col_excel: int) -> bool:
    f = _flatten_formula_value(sheet.range((row_excel, col_excel)).formula)
    return bool(f.startswith("="))


def _force_clear_block_data_cols_xlwings(sheet: Any, block: BalanceAutoBlockConfig) -> int:
    if not block.force_clear_cols:
        return 0
    r_first = block.first_data_row_excel
    r_last = block.last_data_row_excel
    n = 0
    for letters in block.force_clear_cols:
        c_excel = excel_col_letters_to_index(letters.strip().upper()) + 1
        for r in range(r_first, r_last + 1):
            sheet.range((r, c_excel)).value = None
            n += 1
    return n


def _force_clear_blocks_xlwings(sheet: Any, cfg: BalanceAutoConfig) -> int:
    return _force_clear_block_data_cols_xlwings(sheet, cfg.ecuador) + _force_clear_block_data_cols_xlwings(
        sheet,
        cfg.colombia,
    )


@dataclass(frozen=True)
class TransportTarget:
    """Куда вписать «Транспорт трак» в одной строке Excel."""

    row_excel: int  # 1-based
    transport_col_excel: int  # 1-based
    awb_digits: str
    weight: float


def _digits_only(s: Any) -> str:
    if isinstance(s, float) and s == int(s):
        return str(int(s))
    return re.sub(r"\D+", "", str(s or ""))


def _format_awb_for_excel(awb: str) -> str:
    """AWB в Excel как текст с ведущими нулями (префикс 074… и т.п.)."""
    d = _digits_only(awb)
    if not d:
        return awb.strip()
    if len(d) <= 11:
        return d.zfill(11)
    return d


def _awb_match_keys(digits: str) -> frozenset[str]:
    """Набор ключей для сопоставления AWB: Excel часто съедает ведущий 0 (074… → 74…)."""
    d = _digits_only(digits)
    if not d:
        return frozenset()
    keys: set[str] = {d, d.lstrip("0") or "0"}
    if len(d) >= 2:
        keys.add(d[:-1])
        keys.add((d.lstrip("0") or "0")[:-1] if len(d.lstrip("0") or "0") >= 2 else "")
    for width in (10, 11, 12):
        if len(d) <= width:
            keys.add(d.zfill(width))
        stripped = d.lstrip("0") or "0"
        if len(stripped) <= width:
            keys.add(stripped.zfill(width))
        if len(d) > width:
            keys.add(d[-width:])
    keys.discard("")
    return frozenset(keys)


def _awb_variants(digits: str) -> list[str]:
    """Совместимость: все варианты ключей AWB для поиска."""
    return sorted(_awb_match_keys(digits), key=len, reverse=True)


def _read_transport_targets(
    workbook_path: Path,
    cfg: BalanceAutoConfig,
) -> list[TransportTarget]:
    """Считывает строки данных каждого блока, у которого задан transport_col, и собирает (AWB, вес)."""
    rb = xlrd.open_workbook(str(workbook_path))
    sh = rb.sheet_by_name(cfg.sheet_name)
    out: list[TransportTarget] = []
    for block in (cfg.ecuador, cfg.colombia):
        if not block.transport_col:
            continue
        cw = excel_col_letters_to_index(block.weight_col)
        ca = excel_col_letters_to_index(block.awb_col)
        ct = excel_col_letters_to_index(block.transport_col)
        for r in range(block.first_data_row_excel - 1, block.last_data_row_excel):
            try:
                w_raw = sh.cell_value(r, cw)
                a_raw = sh.cell_value(r, ca)
            except IndexError:
                continue
            digits = _digits_only(a_raw)
            try:
                w = float(w_raw) if w_raw not in ("", None) else 0.0
            except (TypeError, ValueError):
                w = 0.0
            if not digits or w <= 0:
                continue
            out.append(
                TransportTarget(
                    row_excel=r + 1,
                    transport_col_excel=ct + 1,
                    awb_digits=digits,
                    weight=w,
                )
            )
    return out


def _lookup_awb_cost(awb: str, awb_to_cost: dict[str, float]) -> tuple[str | None, float | None]:
    """Ищет cost: учитывает ведущие нули, check-digit и Excel-числа без нуля."""
    want = _awb_match_keys(awb)
    if not want:
        return None, None
    for k, v in awb_to_cost.items():
        if _awb_match_keys(k) & want:
            return k, v
    return None, None


def compute_transport_writes(
    targets: list[TransportTarget],
    awb_to_cost: dict[str, float],
) -> tuple[dict[tuple[int, int], float], list[str], list[str]]:
    """
    Считает значения «Транспорт трак»: для каждого AWB — одна запись в первой
    строке с этим номером (полная стоимость с del-mir). Если в Auto_new два
    одинаковых AWB — в «Транспорт трак» подставляется только одно значение.
    Возвращает (writes, missing_awbs, notes).
    """
    by_awb: dict[str, list[TransportTarget]] = {}
    for t in targets:
        by_awb.setdefault(t.awb_digits, []).append(t)

    writes: dict[tuple[int, int], float] = {}
    missing: list[str] = []
    notes: list[str] = []

    for awb, items in by_awb.items():
        matched_key, cost = _lookup_awb_cost(awb, awb_to_cost)
        if cost is None:
            missing.append(awb)
            continue
        if matched_key and matched_key != awb:
            notes.append(
                f"AWB {awb}: сопоставлен с del-mir {matched_key}"
                f"{' (ведущий 0 / check-digit)' if _awb_match_keys(awb) != _awb_match_keys(matched_key) else ''}.",
            )
        if not items:
            continue
        first = min(items, key=lambda t: t.row_excel)
        writes[(first.row_excel, first.transport_col_excel)] = round(cost, 2)
        if len(items) > 1:
            skipped = sorted(t.row_excel for t in items if t.row_excel != first.row_excel)
            notes.append(
                f"AWB {awb}: {len(items)} строк(и) с одинаковым номером — "
                f"транспорт {cost:g} только в строку {first.row_excel}"
                + (f" (пропуск: {', '.join(map(str, skipped))})" if skipped else "")
                + ".",
            )
    return writes, missing, notes


def _weight_col_excel_for_row(cfg: BalanceAutoConfig, row_excel: int) -> int | None:
    for block in (cfg.ecuador, cfg.colombia):
        if block.first_data_row_excel <= row_excel <= block.last_data_row_excel:
            return excel_col_letters_to_index(block.weight_col) + 1
    return None


def _match_transport_cell_format_xlwings(
    sheet: Any,
    row: int,
    col: int,
    ref_col: int,
) -> None:
    """Шрифт/формат как в «Вес» — в шаблоне M часто крупнее остальных колонок."""
    try:
        ref = sheet.range((row, ref_col))
        dst = sheet.range((row, col))
        ref_font = ref.api.Font
        dst_font = dst.api.Font
        dst_font.Size = ref_font.Size
        dst_font.Name = ref_font.Name
        dst_font.Bold = ref_font.Bold
        dst_font.Italic = ref_font.Italic
        dst.api.NumberFormat = ref.api.NumberFormat
    except Exception:
        pass


def _apply_transport_writes_xlwt(
    path: Path,
    cfg: BalanceAutoConfig,
    writes: dict[tuple[int, int], float],
) -> int:
    rb = xlrd.open_workbook(str(path), formatting_info=True)
    sh = rb.sheet_by_name(cfg.sheet_name)
    sheet_idx = sh.number
    wb = xlutils_copy(rb)
    ws = wb.get_sheet(sheet_idx)
    for (r_excel, c_excel), val in writes.items():
        r0, c0 = r_excel - 1, c_excel - 1
        wcol = _weight_col_excel_for_row(cfg, r_excel)
        if wcol:
            try:
                xf_idx = sh.cell(r0, wcol - 1).xf_index
                ws.write(r0, c0, val, rb.format_map[xf_idx])
            except Exception:
                ws.write(r0, c0, val)
        else:
            ws.write(r0, c0, val)
    wb.save(str(path))
    return len(writes)


def _apply_transport_writes_xlwings(
    path: Path,
    cfg: BalanceAutoConfig,
    writes: dict[tuple[int, int], float],
) -> int:
    import xlwings as xw

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    wb = app.books.open(str(path))
    try:
        sht = wb.sheets[cfg.sheet_name]
        for (r_excel, c_excel), val in writes.items():
            sht.range((r_excel, c_excel)).value = val
            wcol = _weight_col_excel_for_row(cfg, r_excel)
            if wcol:
                _match_transport_cell_format_xlwings(sht, r_excel, c_excel, wcol)
        wb.save()
        return len(writes)
    finally:
        try:
            wb.close()
        finally:
            app.quit()


def apply_transport_costs(
    workbook_path: Path,
    cfg: BalanceAutoConfig,
    awb_to_cost: dict[str, float],
) -> tuple[int, list[str], list[str]]:
    """
    Читает текущие AWB и веса в блоках с заданным transport_col, считает доли стоимости
    и записывает их в столбец Транспорт трак. Возвращает (записано, отсутствующие_AWB, заметки).
    """
    path = workbook_path.resolve()
    if not path.exists():
        raise FileNotFoundError(path)

    notes: list[str] = []
    if cfg.backup_before_save:
        bak = path.with_name(path.name + cfg.backup_suffix)
        shutil.copy2(path, bak)

    targets = _read_transport_targets(path, cfg)
    notes.append(f"Найдено целевых строк (AWB+вес) в файле: {len(targets)}")
    if not targets:
        return 0, [], notes

    writes, missing, calc_notes = compute_transport_writes(targets, awb_to_cost)
    notes.extend(calc_notes)

    if not writes:
        notes.append("Нет совпадений AWB с del-mir.com — записывать нечего.")
        return 0, missing, notes

    engine = cfg.excel_engine
    if engine == "xlwings":
        written = _apply_transport_writes_xlwings(path, cfg, writes)
        notes.append(f"Сохранение «Транспорт трак»: xlwings — записано {written} ячеек.")
    elif engine == "xlwt":
        written = _apply_transport_writes_xlwt(path, cfg, writes)
        notes.append(f"Сохранение «Транспорт трак»: xlwt — записано {written} ячеек (VBA пересобирается).")
    else:  # auto
        try:
            written = _apply_transport_writes_xlwings(path, cfg, writes)
            notes.append(f"Сохранение «Транспорт трак»: xlwings (auto) — записано {written} ячеек.")
        except Exception as e:
            written = _apply_transport_writes_xlwt(path, cfg, writes)
            notes.append(
                f"xlwings недоступен ({e!s}); записано через xlwt — {written} ячеек. Макросы могут пропасть.",
            )
    return written, missing, notes


def _clear_ranges_xlwings(
    sheet: Any,
    ranges: list[BalanceAutoClearRange],
) -> int:
    cleared = 0
    for rect in ranges:
        rt, rb, cl, cr = inclusive_rect_cells(rect.top_left, rect.bottom_right)
        for r0 in range(rt, rb + 1):
            for c0 in range(cl, cr + 1):
                row_excel, col_excel = r0 + 1, c0 + 1
                if _xlwings_cell_has_formula(sheet, row_excel, col_excel):
                    continue
                sheet.range((row_excel, col_excel)).value = None
                cleared += 1
    return cleared


def _write_block_xlwings(
    sheet: Any,
    block: BalanceAutoBlockConfig,
    rows: list[FlightFillRow],
) -> tuple[int, int]:
    r_first = block.first_data_row_excel
    r_last = block.last_data_row_excel
    cw = excel_col_letters_to_index(block.weight_col) + 1
    ca = excel_col_letters_to_index(block.awb_col) + 1
    cp = excel_col_letters_to_index(block.price_col) + 1
    cap = r_last - r_first + 1
    written = 0
    for i, fr in enumerate(rows[:cap]):
        r = r_first + i
        sheet.range((r, cw)).value = fr.weight
        awb_cell = sheet.range((r, ca))
        awb_cell.api.NumberFormat = "@"
        awb_cell.value = _format_awb_for_excel(fr.awb)
        sheet.range((r, cp)).value = fr.price
        written += 1
    skipped = max(0, len(rows) - cap)
    return written, skipped


def _apply_balance_flights_xlwings(
    path: Path,
    cfg: BalanceAutoConfig,
    flights: list[FlightFillRow],
    notes: list[str],
) -> tuple[int, int, int, int]:
    import xlwings as xw

    ec = [f for f in flights if f.platform == "ecuador"]
    co = [f for f in flights if f.platform == "colombia"]

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    wb = app.books.open(str(path))
    try:
        sht = wb.sheets[cfg.sheet_name]
        cleared = _clear_ranges_xlwings(sht, cfg.clear_ranges)

        forced = _force_clear_blocks_xlwings(sht, cfg)
        if forced:
            notes.append(
                f"Дополнительно очищено {forced} ячеек в колонках force_clear_cols "
                f"(строки данных; формулы в этих ячейках сброшены, итоги пересчитаются).",
            )

        we, oe = _write_block_xlwings(sht, cfg.ecuador, ec)
        wc, oc = _write_block_xlwings(sht, cfg.colombia, co)
        overflow = oe + oc
        wb.save()
        return cleared, we, wc, overflow
    finally:
        try:
            wb.close()
        finally:
            app.quit()


def apply_balance_flights(
    workbook_path: Path,
    cfg: BalanceAutoConfig,
    flights: list[FlightFillRow],
) -> tuple[int, int, int, int, list[str]]:
    """
    Очищает диапазоны (без формул), записывает строки по платформе.

    Возвращает (очищено_ячеек, записано_ecuador, записано_colombia, переполнение, заметки_для_лога).
    Для сохранения VBA-макросов в .xls нужен установленный Excel и режим xlwings (см. balance_auto.excel_engine).
    """
    path = workbook_path.resolve()
    if not path.exists():
        raise FileNotFoundError(path)

    notes: list[str] = []
    if cfg.backup_before_save:
        bak = path.with_name(path.name + cfg.backup_suffix)
        shutil.copy2(path, bak)

    engine = cfg.excel_engine

    def _prepend_save_explanation(*, used: str) -> None:
        notes.insert(
            0,
            f"Сохранение книги: {used} (в config.yaml balance_auto.excel_engine={engine!r}). "
            "Если при открытии Excel не спрашивает про макросы: (1) через xlwt VBA уже удалён из файла — спрашивать нечего; "
            "(2) через xlwings макросы в файле есть, но запрос «Включить содержимое» может не показываться из‑за доверенного расположения папки, "
            "настройки «Отключить все макросы без уведомления» или особенностей Excel на Mac (Параметры → Безопасность центра управления безопасностью → Параметры макросов).",
        )

    if engine == "xlwt":
        c, we, wc, ov = _apply_balance_flights_xlwt(path, cfg, flights, notes)
        _prepend_save_explanation(used="xlwt — файл пересобран без проекта VBA, макросы обычно теряются")
        return c, we, wc, ov, notes

    if engine == "xlwings":
        try:
            c, we, wc, ov = _apply_balance_flights_xlwings(path, cfg, flights, notes)
        except ImportError as e:
            raise RuntimeError(
                "Нужен пакет xlwings и установленный Microsoft Excel для сохранения макросов.",
            ) from e
        notes.append("Сохранено через Excel (xlwings) — проект VBA и макросы сохраняются.")
        _prepend_save_explanation(used="xlwings (через Excel.app)")
        return c, we, wc, ov, notes

    # auto
    try:
        c, we, wc, ov = _apply_balance_flights_xlwings(path, cfg, flights, notes)
        notes.append("Сохранено через Excel (xlwings) — проект VBA и макросы сохраняются.")
        _prepend_save_explanation(used="xlwings (через Excel.app, режим auto)")
        return c, we, wc, ov, notes
    except ImportError as e:
        notes.append(
            f"xlwings не установлен ({e!s}); используется xlwt — макросы VBA в файле будут потеряны. "
            "Установите зависимости и Excel, затем перезапустите.",
        )
        c, we, wc, ov = _apply_balance_flights_xlwt(path, cfg, flights, notes)
        _prepend_save_explanation(used="xlwt (fallback: xlwings не импортируется)")
        return c, we, wc, ov, notes
    except Exception as e:
        notes.append(
            f"xlwings/Excel недоступен ({e!s}); используется xlwt — макросы VBA в файле могут быть потеряны.",
        )
        c, we, wc, ov = _apply_balance_flights_xlwt(path, cfg, flights, notes)
        _prepend_save_explanation(used=f"xlwt (fallback после ошибки xlwings/Excel: {e!s})")
        return c, we, wc, ov, notes
