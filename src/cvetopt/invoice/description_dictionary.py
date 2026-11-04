from __future__ import annotations

import re
import shutil
import sys
from collections.abc import Callable, Iterable
from dataclasses import dataclass
from pathlib import Path

import xlrd

from cvetopt.invoice.xlsx_read import _COL_RE, read_xlsx_grid

LogFn = Callable[[str], None]

# Словарь.xls: B — английское название, C — перевод.
_DICT_COL_EN = 1  # B, 0-based
_DICT_COL_RU = 2  # C, 0-based
_DICT_COL_EN_LETTER = "B"
_DICT_COL_RU_LETTER = "C"

# Служебные строки инвойса Enigma — не сорта, не дописываем в словарь.
# Нижняя таблица Cargo (Carnations, Roses…) и блок fees/box в длинном файле папки 2.
_NON_PRODUCT_EXACT = frozenset(
    {
        "description",
        "subtotal",
        "subtotal goods",
        "total",
        "total in",
        "total out",
        "artnr",
        "art nr",
        "art nr.",
        "box nr.",
        "box nr",
        "package",
        "packages",
        "others",
        "other",
        "enigma flowers",
        "aqua gerbera",
        # Cargo summary (англ. категории внизу инвойса)
        "carnations",
        "chrysanthemums",
        "eustoma",
        "gerbera",
        "roses",
        "alstroemeria",
        "gypsophila",
        "lilies",
        "lilies (lilium spp.)",
        "lilium",
    }
)
_NON_PRODUCT_PREFIXES = (
    "subtotal",
    "administration",
    "declaration",
    "temp tale",
    "box ",
    "total ",
    "aqua ",
)
_NON_PRODUCT_CONTAINS = (
    " fees",
    " fee",
    " costs",
    " cost",
)
_ORDER_CODE_RE = re.compile(
    r"^(\d{5,}|\d+\s*/\s*[A-Z0-9]+\s*-\s*\d{6,}|[A-Z]{2,}\d{2,}\s*-\s*\d{6,})",
    re.IGNORECASE,
)
_ONLY_CODE_CHARS_RE = re.compile(r"^[\d\s./_\-]+$")
# Сводка Cargo внизу инвойса Enigma (Paeonia, Orchids, Tulipa…): первая ячейка строки —
# таможенный код 0603…/0604… из 8 цифр, у товаров там артикул из 5–6 цифр. (Тот же код
# есть и у товаров, но правее — поэтому смотрим только первую заполненную ячейку.)
_CUSTOMS_CODE_RE = re.compile(r"^'?06\d{6}'?$")


def is_cargo_summary_row(values: Iterable[object]) -> bool:
    """values — ячейки строки слева направо. Строка сводки Cargo — не сорт."""
    first = next((v for v in values if str(v if v is not None else "").strip()), None)
    if isinstance(first, (int, float)) and not isinstance(first, bool):
        # Excel мог съесть ведущий ноль: 06031970 → 6031970.
        return float(first).is_integer() and 6_010_000 <= first < 6_050_000
    return bool(_CUSTOMS_CODE_RE.match(str(first or "").strip()))


def row_values_in_order(row: dict[str, object]) -> list[object]:
    """Ячейки строки xlsx {«A»: …, «AB»: …} слева направо."""
    return [row[col] for col in sorted(row, key=lambda col: (len(col), col))]


@dataclass(frozen=True)
class SkipDescriptionRules:
    """Дополнительные фразы из Настроек — не переводить и не дописывать в словарь."""

    exact: frozenset[str] = frozenset()
    prefixes: tuple[str, ...] = ()
    contains: tuple[str, ...] = ()


SKIP_DESCRIPTION_RULES_EMPTY = SkipDescriptionRules()


def parse_skip_descriptions_text(raw: str) -> SkipDescriptionRules:
    """
    Одна фраза на строку (из UI Настроек):
    - точное совпадение (без учёта регистра);
    - «текст*» — строка начинается с «текст»;
    - «~текст» — строка содержит «текст»;
    - строки с # в начале — комментарий.
    """
    exact: set[str] = set()
    prefixes: list[str] = []
    contains: list[str] = []
    for line in (raw or "").splitlines():
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        if line.startswith("~"):
            token = line[1:].strip().casefold()
            if token:
                contains.append(token)
            continue
        if line.endswith("*"):
            token = line[:-1].strip().casefold()
            if token:
                prefixes.append(token)
            continue
        exact.add(line.casefold())
    return SkipDescriptionRules(
        frozenset(exact),
        tuple(prefixes),
        tuple(contains),
    )


def _matches_skip_rules(folded: str, rules: SkipDescriptionRules) -> bool:
    if folded in rules.exact:
        return True
    if any(folded.startswith(p) for p in rules.prefixes):
        return True
    if any(c in folded for c in rules.contains):
        return True
    return False


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


_CYRILLIC_LETTER_RE = re.compile(r"[а-яё]", re.IGNORECASE)
_LATIN_LETTER_RE = re.compile(r"[a-z]", re.IGNORECASE)
# Длина в названии Enigma: «60cm», «65 Cm», «60см». Не «*96» (штук в упаковке) и не «5+».
_LENGTH_IN_NAME_RE = re.compile(r"(?i)\b\d+\s*(?:cm|см)(?!\w)")


def _norm_key(text: str) -> str:
    return " ".join(str(text or "").split()).strip()


def looks_like_russian_description(text: str) -> bool:
    """Description уже на русском — переводить/дописывать в колонку B словаря не нужно."""
    key = _norm_key(text)
    if not key:
        return False
    cyr = len(_CYRILLIC_LETTER_RE.findall(key))
    if cyr == 0:
        return False
    lat = len(_LATIN_LETTER_RE.findall(key))
    if lat == 0:
        return True
    return cyr > lat


def is_holland_product_description(
    text: str,
    *,
    extra: SkipDescriptionRules | None = None,
) -> bool:
    """
    Строка Description из инвойса — товар/сорт (переводим и можно дописать в словарь).
    False для fees, box, кодов заказа, итогов и прочего служебного текста.
    """
    key = _norm_key(text)
    if not key:
        return False
    folded = key.casefold()
    if extra is not None and _matches_skip_rules(folded, extra):
        return False
    if folded in _NON_PRODUCT_EXACT:
        return False
    for prefix in _NON_PRODUCT_PREFIXES:
        if folded.startswith(prefix):
            return False
    padded = f" {folded}"
    if any(token in padded for token in _NON_PRODUCT_CONTAINS):
        return False
    if _ORDER_CODE_RE.match(key):
        return False
    if "fmyx" in folded:
        return False
    compact = folded.replace(" ", "")
    if len(compact) >= 10 and compact.isalnum() and not any(c.islower() for c in key if c.isalpha()):
        return False
    if _ONLY_CODE_CHARS_RE.match(key) and any(ch.isdigit() for ch in key):
        return False
    return True


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


def translate_description(
    dictionary: dict[str, str],
    description: str,
) -> tuple[str, bool, bool]:
    """
    Переводит Description по словарю — только точное совпадение всей строки
    (или той же строки без длины «60cm»).

    Без точного совпадения строка остаётся оригиналом **целиком**: частичный
    перевод по словам давал гибриды вроде «Hyp Mag зеленый Power», которые
    потом копировали в словарь.

    Возвращает (текст_для_колонки, был_точный_хит, хоть_один_токен_из_словаря);
    третий флаг теперь совпадает со вторым.
    """
    key = _norm_key(description)
    if not key:
        return "", False, False

    exact = lookup_translation(dictionary, key)
    if exact is not None:
        return exact, True, True
    # Тот же цветок без длины («Oxy Co Tan Pure Blu 60cm» → «Oxy Co Tan Pure Blu»):
    # переводим название, длина остаётся в колонке S1.
    base = _without_length(key)
    if base and base != key:
        exact = lookup_translation(dictionary, base)
        if exact is not None:
            return exact, True, True
    return key, False, False


def _without_length(text: str) -> str:
    return _norm_key(_LENGTH_IN_NAME_RE.sub(" ", text))


def lookup_translation(dictionary: dict[str, str], description: str) -> str | None:
    """Точное совпадение всей строки Description со словарём (B → C)."""
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
    *,
    extra: SkipDescriptionRules | None = None,
) -> tuple[list[str], int]:
    translated_folded = {k.casefold() for k in dictionary}
    out: list[str] = []
    seen: set[str] = set()
    skipped_non_product = 0
    for raw in descriptions:
        key = _norm_key(raw)
        if not key:
            continue
        if not is_holland_product_description(key, extra=extra):
            skipped_non_product += 1
            continue
        if looks_like_russian_description(key):
            skipped_non_product += 1
            continue
        folded = key.casefold()
        base = _without_length(key).casefold()
        if folded in seen or folded in translated_folded or folded in existing_keys:
            continue
        if base and base != folded and (base in translated_folded or base in existing_keys):
            continue
        seen.add(folded)
        out.append(key)
    return sorted(out, key=str.casefold), skipped_non_product


def _xls_dictionary_next_row(sheet: xlrd.sheet.Sheet) -> int:
    """Строка после содержимого B/C, без хвоста из пустого текста и пробелов."""
    for row in range(sheet.nrows - 1, -1, -1):
        if any(
            str(sheet.cell_value(row, col)).strip()
            for col in (_DICT_COL_EN, _DICT_COL_RU)
            if col < sheet.ncols
        ):
            return row + 2
    return 2


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
    return _xls_dictionary_next_row(book.sheet_by_index(0))


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
    """Без Excel: копия книги с оформлением (цвета, шрифты, ширина колонок) + новые строки."""
    from xlutils.copy import copy as xl_copy

    from cvetopt.mail.short_postprocess import _copy_workbook_palette, _open_xls_workbook

    rb = _open_xls_workbook(path.read_bytes())
    sheet_in = rb.sheet_by_index(0)
    wb_out = xl_copy(rb)
    _copy_workbook_palette(rb, wb_out)
    ws_out = wb_out.get_sheet(0)
    start = _xls_dictionary_next_row(sheet_in) - 1
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
    extra: SkipDescriptionRules | None = None,
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
    to_add, skipped = _unique_missing_descriptions(
        descriptions, dict_map, existing_keys, extra=extra
    )
    if log is not None and skipped:
        log(
            f"Словарь: пропущено {skipped} служебных строк "
            "(fees, box, коды заказа — не дописываем)"
        )
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
