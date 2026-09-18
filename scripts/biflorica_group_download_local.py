#!/usr/bin/env python3
"""
Тест-кейсы на объединение заказов Biflorica с одной датой вылета в один файл
(как при выборе нескольких галочек на портале и одном клике «Отчёт по сделкам»).

Проверяет только то, что можно проверить без живого портала и Playwright:
  1. Имя файла для одного/нескольких заказов, разбор имени обратно
  2. Реестр/архивация: частично зарегистрированная группа не архивируется раньше времени
  3. Группировка заказов по дате вылета (порядок, состав)
  4. archive_biflorica_download_dir целиком, на реальных файлах на диске

Взаимодействие с самим порталом (клик по галочкам, скачивание — _download_group_report,
_walk_pages_setting_checkboxes, _set_order_checkbox) здесь НЕ покрыто; см. отдельно
biflorica_group_download_e2e_local.py — там та же логика прогоняется через настоящий
Playwright против локального фейкового сервера с реальной вёрсткой портала.

Запуск: .venv/bin/python scripts/biflorica_group_download_local.py
"""

from __future__ import annotations

import shutil
import sys
import tempfile
import traceback
from collections.abc import Callable
from datetime import date
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from cvetopt.core.models import Order  # noqa: E402
from cvetopt.core.runtime_settings import (  # noqa: E402
    archive_biflorica_download_dir,
    biflorica_download_filename,
    flight_date_from_biflorica_report,
    order_id_from_biflorica_report,
    order_ids_from_biflorica_report,
    _biflorica_archive_candidate,
)

TESTS: list[tuple[str, Callable[[Path], None]]] = []


def test(name: str) -> Callable[[Callable[[Path], None]], Callable[[Path], None]]:
    def deco(fn: Callable[[Path], None]) -> Callable[[Path], None]:
        TESTS.append((name, fn))
        return fn

    return deco


def group_orders_by_flight_date(orders: list[Order]) -> dict[date, list[Order]]:
    """То же самое группирование, что в run_biflorica_job — воспроизведено здесь
    для теста, так как сама функция — часть большого async-джоба с Playwright."""
    groups: dict[date, list[Order]] = {}
    for o in orders:
        groups.setdefault(o.flight_date, []).append(o)
    return groups


# === Имя файла и разбор обратно ======================================================


@test("1.1 biflorica_download_filename: один заказ — как раньше")
def _(tmp: Path) -> None:
    name = biflorica_download_filename(["10800933"], date(2026, 9, 12))
    assert name == "BiFlorica-10800933__2026-09-12.xlsx", name


@test("1.2 biflorica_download_filename: несколько заказов — через «+»")
def _(tmp: Path) -> None:
    name = biflorica_download_filename(["10800933", "10801161"], date(2026, 9, 12))
    assert name == "BiFlorica-10800933+10801161__2026-09-12.xlsx", name


@test("1.3 order_ids_from_biflorica_report: групповое имя разбирается на все id")
def _(tmp: Path) -> None:
    p = Path("BiFlorica-10800933+10801161__2026-09-12.xlsx")
    assert order_ids_from_biflorica_report(p) == ["10800933", "10801161"]
    assert order_id_from_biflorica_report(p) == "10800933"  # первый — для обратной совместимости
    assert flight_date_from_biflorica_report(p) == date(2026, 9, 12)


@test("1.4 order_ids_from_biflorica_report: легаси-имя без префикса всё ещё работает")
def _(tmp: Path) -> None:
    p = Path("10800933__2026-09-12.xlsx")
    assert order_ids_from_biflorica_report(p) == ["10800933"]
    assert flight_date_from_biflorica_report(p) == date(2026, 9, 12)


@test("1.5 order_ids_from_biflorica_report: посторонний файл — пустой список, не ошибка")
def _(tmp: Path) -> None:
    assert order_ids_from_biflorica_report(Path("Auto_new.xlsx")) == []
    assert order_id_from_biflorica_report(Path("Auto_new.xlsx")) is None
    assert flight_date_from_biflorica_report(Path("Auto_new.xlsx")) is None


# === Реестр / политика архивации =====================================================


@test("2.1 archive_candidate: одиночный заказ, не зарегистрирован → архивировать")
def _(tmp: Path) -> None:
    p = Path("BiFlorica-10800933__2026-09-12.xlsx")
    assert _biflorica_archive_candidate(
        p, policy="unregistered_only", keep_order_ids=set(), keep_paths=set()
    ) is True


@test("2.2 archive_candidate: одиночный заказ, зарегистрирован → оставить")
def _(tmp: Path) -> None:
    p = Path("BiFlorica-10800933__2026-09-12.xlsx")
    assert _biflorica_archive_candidate(
        p, policy="unregistered_only", keep_order_ids={"10800933"}, keep_paths=set()
    ) is False


@test("2.3 archive_candidate: группа, ЧАСТИЧНО зарегистрирована → НЕ архивировать")
def _(tmp: Path) -> None:
    # Ключевой кейс: новый заказ добавился к уже скачанной дате, в реестре пока
    # только старый — файл всё равно нужно сохранить, а не унести в архив.
    p = Path("BiFlorica-10800933+10801161__2026-09-12.xlsx")
    assert _biflorica_archive_candidate(
        p, policy="unregistered_only", keep_order_ids={"10800933"}, keep_paths=set()
    ) is False


@test("2.4 archive_candidate: группа полностью НЕ зарегистрирована → архивировать")
def _(tmp: Path) -> None:
    p = Path("BiFlorica-10800933+10801161__2026-09-12.xlsx")
    assert _biflorica_archive_candidate(
        p, policy="unregistered_only", keep_order_ids=set(), keep_paths=set()
    ) is True


@test("2.5 archive_candidate: группа полностью зарегистрирована → оставить")
def _(tmp: Path) -> None:
    p = Path("BiFlorica-10800933+10801161__2026-09-12.xlsx")
    assert _biflorica_archive_candidate(
        p,
        policy="unregistered_only",
        keep_order_ids={"10800933", "10801161"},
        keep_paths=set(),
    ) is False


# === Группировка заказов по дате вылета ==============================================


@test("3.1 group_orders_by_flight_date: два заказа на одну дату — одна группа")
def _(tmp: Path) -> None:
    orders = [
        Order(portal_id="biflorica", order_id="10798006", flight_date=date(2026, 9, 5)),
        Order(portal_id="biflorica", order_id="10800933", flight_date=date(2026, 9, 12)),
        Order(portal_id="biflorica", order_id="10801161", flight_date=date(2026, 9, 12)),
    ]
    groups = group_orders_by_flight_date(orders)
    assert list(groups.keys()) == [date(2026, 9, 5), date(2026, 9, 12)], list(groups.keys())
    assert [o.order_id for o in groups[date(2026, 9, 12)]] == ["10800933", "10801161"]
    assert len(groups[date(2026, 9, 5)]) == 1


@test("3.2 group_orders_by_flight_date: пустой список — пустой результат")
def _(tmp: Path) -> None:
    assert group_orders_by_flight_date([]) == {}


# === Корнер-кейсы разбора имени файла ================================================


@test("5.1 order_ids_from_biflorica_report: висячий «+» в конце — не парсится")
def _(tmp: Path) -> None:
    p = Path("BiFlorica-123+__2026-09-12.xlsx")
    assert order_ids_from_biflorica_report(p) == [], order_ids_from_biflorica_report(p)
    assert flight_date_from_biflorica_report(p) is None


@test("5.2 order_ids_from_biflorica_report: висячий «+» в начале — не парсится")
def _(tmp: Path) -> None:
    p = Path("BiFlorica-+123__2026-09-12.xlsx")
    assert order_ids_from_biflorica_report(p) == []


@test("5.3 order_ids_from_biflorica_report: двойной «++» — не парсится")
def _(tmp: Path) -> None:
    p = Path("BiFlorica-123++456__2026-09-12.xlsx")
    assert order_ids_from_biflorica_report(p) == []


@test("5.4 order_ids_from_biflorica_report: регистр префикса не важен")
def _(tmp: Path) -> None:
    p = Path("biflorica-123__2026-09-12.xlsx")
    assert order_ids_from_biflorica_report(p) == ["123"]
    p2 = Path("BIFLORICA-123+456__2026-09-12.xlsx")
    assert order_ids_from_biflorica_report(p2) == ["123", "456"]


@test("5.5 biflorica_download_filename ↔ order_ids_from_biflorica_report: группа из 5, roundtrip")
def _(tmp: Path) -> None:
    ids = ["1", "22", "333", "4444", "55555"]
    name = biflorica_download_filename(ids, date(2026, 12, 31))
    assert order_ids_from_biflorica_report(Path(name)) == ids, (name, order_ids_from_biflorica_report(Path(name)))
    assert flight_date_from_biflorica_report(Path(name)) == date(2026, 12, 31)


@test("5.6 flight_date_from_biflorica_report: синтаксически похожая, но невалидная дата")
def _(tmp: Path) -> None:
    # Регулярка проверяет только форму ДДДД-ДД-ДД, не календарь — месяц 13 не бывает.
    p = Path("BiFlorica-123__2026-13-45.xlsx")
    assert order_ids_from_biflorica_report(p) == ["123"], "id должен разобраться, дата — нет"
    assert flight_date_from_biflorica_report(p) is None


@test("5.7 order_ids_from_biflorica_report: id с ведущими нулями — не теряются")
def _(tmp: Path) -> None:
    p = Path("BiFlorica-00123+00456__2026-09-12.xlsx")
    assert order_ids_from_biflorica_report(p) == ["00123", "00456"]


@test("5.8 order_ids_from_biflorica_report: посторонний Biflorica-файл без id (biflorica-deals…)")
def _(tmp: Path) -> None:
    # Такие имена ловит _is_biflorica_report_to_archive отдельной проверкой на префикс
    # "biflorica", но order_ids из них не достать — это ожидаемо, не баг.
    p = Path("Biflorica-deals-export.xlsx")
    assert order_ids_from_biflorica_report(p) == []


@test("5.9 archive_candidate: файл без разбираемых id, policy=unregistered_only → архивировать")
def _(tmp: Path) -> None:
    # order_ids=[] по определению не входит ни в один keep_order_ids — такой файл
    # всегда считается «неучтённым» и уходит в архив при уборке до скачивания.
    p = Path("Biflorica-deals-export.xlsx")
    assert _biflorica_archive_candidate(
        p, policy="unregistered_only", keep_order_ids={"1", "2", "3"}, keep_paths=set()
    ) is True


@test("5.10 archive_candidate: группа из трёх, зарегистрирован только средний → оставить")
def _(tmp: Path) -> None:
    p = Path("BiFlorica-111+222+333__2026-09-12.xlsx")
    assert _biflorica_archive_candidate(
        p, policy="unregistered_only", keep_order_ids={"222"}, keep_paths=set()
    ) is False


@test("5.11 group_orders_by_flight_date: одна дата, много заказов — порядок сохраняется")
def _(tmp: Path) -> None:
    orders = [
        Order(portal_id="biflorica", order_id=str(i), flight_date=date(2026, 9, 12))
        for i in range(1, 8)
    ]
    groups = group_orders_by_flight_date(orders)
    assert len(groups) == 1
    assert [o.order_id for o in groups[date(2026, 9, 12)]] == [str(i) for i in range(1, 8)]


@test("5.12 group_orders_by_flight_date: все заказы на разные даты — по одной группе каждый")
def _(tmp: Path) -> None:
    orders = [
        Order(portal_id="biflorica", order_id=str(i), flight_date=date(2026, 9, i))
        for i in range(1, 6)
    ]
    groups = group_orders_by_flight_date(orders)
    assert len(groups) == 5
    assert all(len(v) == 1 for v in groups.values())


# === archive_biflorica_download_dir целиком, на реальных файлах =====================


@test("6.1 archive_biflorica_download_dir: смешанный набор — сверено файл за файлом")
def _(tmp: Path) -> None:
    download_dir = tmp / "download"
    archive_dir = tmp / "архив"
    download_dir.mkdir()
    archive_dir.mkdir()

    files = {
        # старый одиночный отчёт, ещё не в реестре — до скачивания должен уйти в архив
        "BiFlorica-10798006__2026-09-05.xlsx": None,
        # группа с частичной регистрацией — должна остаться (не архивировать раньше времени)
        "BiFlorica-10800933+10801161__2026-09-12.xlsx": "10800933",
        # не-Biflorica файл — архив не должен его вообще замечать
        "Auto_new.xlsx": None,
    }
    for name in files:
        (download_dir / name).write_text("stub", encoding="utf-8")

    keep_order_ids = {"10800933"}  # только первый заказ группы уже в реестре
    archive_result_dir, archived, warnings, kept = archive_biflorica_download_dir(
        download_dir,
        archive_dir,
        keep_order_ids=keep_order_ids,
        keep_paths=set(),
        policy="unregistered_only",
    )

    assert not (download_dir / "BiFlorica-10798006__2026-09-05.xlsx").exists(), (
        "неучтённый одиночный отчёт должен был уйти в архив"
    )
    assert (download_dir / "BiFlorica-10800933+10801161__2026-09-12.xlsx").exists(), (
        "частично зарегистрированная группа не должна архивироваться"
    )
    assert (download_dir / "Auto_new.xlsx").exists(), "посторонний файл архив не должен трогать"
    assert "BiFlorica-10798006__2026-09-05.xlsx" in archived, archived
    assert "BiFlorica-10800933+10801161__2026-09-12.xlsx" in kept, kept


@test("6.2 archive_biflorica_download_dir: policy=stale_registered — сессия защищена")
def _(tmp: Path) -> None:
    download_dir = tmp / "download"
    archive_dir = tmp / "архив"
    download_dir.mkdir()
    archive_dir.mkdir()

    current = download_dir / "BiFlorica-10800933+10801161__2026-09-12.xlsx"
    stale = download_dir / "BiFlorica-10798006__2026-09-05.xlsx"
    current.write_text("stub", encoding="utf-8")
    stale.write_text("stub", encoding="utf-8")

    archive_biflorica_download_dir(
        download_dir,
        archive_dir,
        keep_order_ids={"10798006", "10800933", "10801161"},  # все формально в реестре
        keep_paths={current.resolve()},  # но скачан в ЭТОЙ сессии только current
        policy="stale_registered",
    )
    assert current.exists(), "файл текущей сессии не должен архивироваться"
    assert not stale.exists(), "старый отчёт прошлой сессии должен уйти в архив"


# === Раннер ===========================================================================


def main() -> int:
    base = Path(tempfile.mkdtemp(prefix="cvetopt-biflorica-group-tests-"))
    passed = 0
    failed: list[str] = []
    for i, (name, fn) in enumerate(TESTS):
        tmp = base / f"case_{i:02d}"
        tmp.mkdir(parents=True, exist_ok=True)
        try:
            fn(tmp)
        except Exception as e:  # noqa: BLE001
            failed.append(name)
            print(f"FAIL  {name}")
            print(f"      {type(e).__name__}: {e}")
            traceback.print_exc(limit=3)
            print()
        else:
            passed += 1
            print(f"PASS  {name}")
    shutil.rmtree(base, ignore_errors=True)

    print()
    print(f"Итого: {passed}/{len(TESTS)} прошли")
    print(
        "\nВзаимодействие с порталом (галочки, клик по «Отчёт по сделкам») — отдельно "
        "в scripts/biflorica_group_download_e2e_local.py (Playwright + фейковый сервер)."
    )
    if failed:
        print("Упавшие:")
        for name in failed:
            print(f"  - {name}")
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
