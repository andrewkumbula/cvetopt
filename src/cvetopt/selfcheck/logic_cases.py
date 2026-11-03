"""
Проверки логики на синтетических файлах: перевод, словарь, «Переведено», почта, миксы.

Каждый кейс получает пустую папку tmp и создаёт в ней всё, что ему нужно.
Excel не запускается: словарь дописывается через xlwt, инвойсы — через xlutils.
"""

from __future__ import annotations

import hashlib
import shutil
from collections.abc import Callable
from pathlib import Path
from types import SimpleNamespace

import xlrd
import xlwt

from cvetopt.core.job_messages import humanize_error
from cvetopt.invoice.biflorica_mixes import run_mix_separation
from cvetopt.invoice.description_dictionary import (
    _append_rows_xlwt,
    _unique_missing_descriptions,
    load_description_dictionary,
    load_dictionary_english_keys,
    translate_description,
)
from cvetopt.invoice.holland_mail_continue import process_mail_workbook
from cvetopt.mail.attachments import _filter_ids_by_from, _imap_search_criteria
from cvetopt.selfcheck.mix_cases import make_biflorica, make_template, mix_row

TESTS: list[tuple[str, Callable[[Path], None]]] = []


def test(name: str) -> Callable[[Callable[[Path], None]], Callable[[Path], None]]:
    def deco(fn: Callable[[Path], None]) -> Callable[[Path], None]:
        TESTS.append((name, fn))
        return fn

    return deco


# --- Хелперы ---------------------------------------------------------------------

DICTIONARY = {
    "Oxy Co Tan Pure Blu": "Оксипеталум Перу",
    "Solidago groen 080cm": "Солидаго Зеленый 80 см",
    "Solidago groen": "Солидаго Зеленый",
    "Soli Carzan Glory Ex": "Солидаго Глори",
    "Hyp Mag Green Power": "Гиперикум Магик грин",
    "green": "зеленый",
    "R Tr Bombastic": "Роза Бомбастик Куст",
    "Ge Gr Aquabox Mix *96": "Гербера Аква бокс",
    "Li Ot Zambesi": "Лилия Замбези",
}


def make_dictionary_xls(path: Path, rows: dict[str, str]) -> Path:
    book = xlwt.Workbook()
    sheet = book.add_sheet("Список")
    sheet.write(0, 1, "Фраза")
    sheet.write(0, 2, "Перевод")
    for i, (en, ru) in enumerate(rows.items(), start=1):
        sheet.write(i, 1, en)
        sheet.write(i, 2, ru)
    book.save(str(path))
    return path


def make_enigma_invoice_xls(path: Path, lines: list[tuple[str, str]]) -> Path:
    """Как файл Enigma: шапка, затем Artnr … Description … S1; lines = (Description, S1)."""
    book = xlwt.Workbook()
    sheet = book.add_sheet("Invoice")
    sheet.write(0, 0, "Invoice")
    sheet.write(5, 0, "  Trading B.V.")
    header = ["Artnr", "Kolli", "Inh", "Total", "Description", "Price", "Amount", "S1"]
    for col, text in enumerate(header):
        sheet.write(8, col, text)
    for i, (desc, s1) in enumerate(lines, start=9):
        sheet.write(i, 0, 100000 + i)
        sheet.write(i, 4, desc)
        sheet.write(i, 7, s1)
    book.save(str(path))
    return path


def column_values(path: Path, col: int, first_row: int) -> list[str]:
    sheet = xlrd.open_workbook(str(path)).sheet_by_index(0)
    return [str(sheet.cell_value(r, col)) for r in range(first_row, sheet.nrows)]


def file_digest(path: Path) -> str:
    return hashlib.md5(path.read_bytes()).hexdigest()


# === Перевод ======================================================================


@test("Перевод: точное совпадение, регистр и пробелы не важны")
def _(tmp: Path) -> None:
    assert translate_description(DICTIONARY, "oxy co  tan pure blu")[:2] == ("Оксипеталум Перу", True)


@test("Перевод: длина 60cm / 65 Cm / 70см не мешает, в перевод не попадает")
def _(tmp: Path) -> None:
    for text in ("Oxy Co Tan Pure Blu 60cm", "Oxy Co Tan Pure Blu 65 Cm", "Oxy Co Tan Pure Blu 70см"):
        got = translate_description(DICTIONARY, text)
        assert got[:2] == ("Оксипеталум Перу", True), (text, got)


@test("Перевод: длина в середине названия («… 60cm Ex»)")
def _(tmp: Path) -> None:
    got = translate_description(DICTIONARY, "Soli Carzan Glory 60cm Ex")
    assert got[:2] == ("Солидаго Глори", True), got


@test("Перевод: запись словаря с длиной важнее записи без длины")
def _(tmp: Path) -> None:
    got = translate_description(DICTIONARY, "Solidago groen 080cm")
    assert got[0] == "Солидаго Зеленый 80 см", got


@test("Перевод: без совпадения — оригинал целиком, без частичного («зеленый»)")
def _(tmp: Path) -> None:
    got = translate_description(DICTIONARY, "Hyp Mag Green Spirit")
    assert got == ("Hyp Mag Green Spirit", False, False), got


@test("Перевод: «*96» и «5+» — не длина, не отрезаются")
def _(tmp: Path) -> None:
    assert translate_description(DICTIONARY, "Ge Gr Aquabox Mix *96")[1]
    assert not translate_description(DICTIONARY, "Ge Gr Aquabox Mix *184")[1]
    assert not translate_description(DICTIONARY, "Li Ot Zambesi 5+")[1]


# === Словарь ======================================================================


@test("Словарь: новое дописывается в конец с пустым переводом, дубли — нет")
def _(tmp: Path) -> None:
    path = make_dictionary_xls(tmp / "Словарь.xls", DICTIONARY)
    dictionary = load_description_dictionary(path)
    keys = load_dictionary_english_keys(path)
    to_add, _skipped = _unique_missing_descriptions(
        [
            "New Flower One",
            "new flower one",            # тот же, другой регистр
            "Oxy Co Tan Pure Blu 60cm",  # есть в словаре без длины
            "Hyp Mag Green Power",       # уже есть
            "Роза Бомбастик Куст 70",    # уже переведённый текст
            "Transport fees",            # служебная строка инвойса
        ],
        dictionary,
        keys,
    )
    assert to_add == ["New Flower One"], to_add
    _append_rows_xlwt(path, [(text, "") for text in to_add])
    sheet = xlrd.open_workbook(str(path)).sheet_by_index(0)
    last = sheet.row_values(sheet.nrows - 1)[1:3]
    assert last == ["New Flower One", ""], last
    assert sheet.nrows == len(DICTIONARY) + 2, sheet.nrows


# === «Переведено»: перевод инвойса Enigma на месте =================================


@test("«Переведено»: перевод в ячейках Description, ростовка у роз, неизвестное — как есть")
def _(tmp: Path) -> None:
    dict_path = make_dictionary_xls(tmp / "Словарь.xls", DICTIONARY)
    invoice = make_enigma_invoice_xls(
        tmp / "invoice.xls",
        [
            ("R Tr Bombastic", "70 Cm"),
            ("Oxy Co Tan Pure Blu 60cm", "60 Cm"),
            ("Brand New Sort", "55 Cm"),
        ],
    )
    result = process_mail_workbook(invoice, dict_path, append_missing_to_dictionary=False)
    got = column_values(invoice, 4, 9)
    assert got == ["Роза Бомбастик Куст 70", "Оксипеталум Перу", "Brand New Sort"], got
    assert result.missing == ["Brand New Sort"], result.missing


@test("«Переведено»: повторный прогон того же файла не находит новых слов для словаря")
def _(tmp: Path) -> None:
    dict_path = make_dictionary_xls(tmp / "Словарь.xls", DICTIONARY)
    invoice = make_enigma_invoice_xls(
        tmp / "invoice.xls",
        [("R Tr Bombastic", "70 Cm"), ("Hyp Mag Green Power", "60 Cm")],
    )
    process_mail_workbook(invoice, dict_path, append_missing_to_dictionary=False)
    before = column_values(invoice, 4, 9)
    again = process_mail_workbook(invoice, dict_path, append_missing_to_dictionary=False)
    assert again.missing == [], again.missing
    assert column_values(invoice, 4, 9) == before, column_values(invoice, 4, 9)


# === Почта ========================================================================


@test("Почта: обычный поиск — дата + отправитель, запасной — только дата")
def _(tmp: Path) -> None:
    cfg = SimpleNamespace(only_unread=False, from_contains=["noreply@enigmaflowers.nl"])
    assert _imap_search_criteria(cfg, "27-Oct-2026") == (
        '(SINCE 27-Oct-2026 FROM "noreply@enigmaflowers.nl")'
    )
    assert _imap_search_criteria(cfg, "27-Oct-2026", with_from=False) == "(SINCE 27-Oct-2026)"


@test("Почта: запасной поиск сам отбирает письма Enigma по заголовку From")
def _(tmp: Path) -> None:
    senders = {
        b"1": "Enigma <noreply@enigmaflowers.nl>",
        b"2": "shop@flowersplants.shop",
        b"3": "noreply@enigmaflowers.nl",
    }

    class FakeImap:
        def fetch(self, num: bytes, spec: str):
            assert "PEEK" in spec, "письма нельзя помечать прочитанными"
            return "OK", [(b"hdr", f"From: {senders[num]}\r\n\r\n".encode())]

    cfg = SimpleNamespace(from_contains=["noreply@enigmaflowers.nl"])
    assert _filter_ids_by_from(FakeImap(), list(senders), cfg) == [b"1", b"3"]


@test("Ошибки: пропал DNS — понятный текст вместо gaierror")
def _(tmp: Path) -> None:
    text = humanize_error("gaierror: [Errno 11001] getaddrinfo failed")
    assert "Нет связи" in text, text


# === Миксы ========================================================================


def _mix_case(tmp: Path) -> tuple[Path, Path]:
    bif = make_biflorica(tmp / "BiFlorica-test.xlsx", [
        {"plantation": "MIRO", "variety": "Mix", "prices": {"60": 0.3}, "stems": 300},
        {"plantation": "MIRO2", "variety": "Mix", "prices": {"60": 0.3}, "stems": 300},
    ])
    tpl = make_template(tmp / "Шаблон.xlsx", [
        {"code": "Mix W", "qtys": {"60": 200}},
        {"code": "Mix P", "qtys": {"60": 125}},
    ], lengths=("60",))
    return bif, tpl


@test("Миксы: 325 шт на 60 см = 300 из одной строки Mix + 25 из другой")
def _(tmp: Path) -> None:
    bif, tpl = _mix_case(tmp)
    _demand, plans, _out = run_mix_separation(template_path=tpl, biflorica_path=bif)
    takes = sorted(t.take_qty for t in plans[0].takes)
    assert takes == [25, 300], takes
    left = [mix_row(bif, p) for p in ("MIRO", "MIRO2")]
    remaining = [str(r["O"]) for r in left if r is not None]
    assert remaining == ["275"], remaining


@test("Миксы: повторный запуск на том же файле останавливается и ничего не меняет")
def _(tmp: Path) -> None:
    bif, tpl = _mix_case(tmp)
    run_mix_separation(template_path=tpl, biflorica_path=bif)
    digest = file_digest(bif)
    backups = sorted(tmp.glob("* до миксов *.xlsx"))
    try:
        run_mix_separation(template_path=tpl, biflorica_path=bif)
    except RuntimeError as e:
        assert "уже внесены" in str(e) and backups[-1].name in str(e), e
    else:
        raise AssertionError("повторный запуск должен был остановиться")
    assert file_digest(bif) == digest, "файл изменился"
    assert sorted(tmp.glob("* до миксов *.xlsx")) == backups, "появилась лишняя копия"


@test("Миксы: строка с плантацией «-» в исходном отчёте не мешает первому запуску")
def _(tmp: Path) -> None:
    # В отчётах Biflorica бывает «-» вместо плантации (Freedom, Explorer…), а копии «до миксов» нет.
    bif = make_biflorica(tmp / "BiFlorica-test.xlsx", [
        {"plantation": "MIRO", "variety": "Mix", "prices": {"60": 0.3}, "stems": 600},
        {"plantation": "-", "variety": "Freedom", "prices": {"60": 0.4}, "stems": 100},
    ])
    tpl = make_template(tmp / "Шаблон.xlsx", [
        {"code": "Mix W", "qtys": {"60": 200}},
        {"code": "Mix P", "qtys": {"60": 125}},
    ], lengths=("60",))
    _demand, plans, _out = run_mix_separation(template_path=tpl, biflorica_path=bif)
    assert sum(t.take_qty for t in plans[0].takes) == 325, plans


@test("Миксы: после возврата файла из копии «до миксов» снова запускаются")
def _(tmp: Path) -> None:
    bif, tpl = _mix_case(tmp)
    run_mix_separation(template_path=tpl, biflorica_path=bif)
    backup = sorted(tmp.glob("* до миксов *.xlsx"))[-1]
    shutil.copy2(backup, bif)
    _demand, plans, _out = run_mix_separation(template_path=tpl, biflorica_path=bif)
    assert sum(t.take_qty for t in plans[0].takes) == 325, plans
