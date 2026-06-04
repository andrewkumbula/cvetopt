from __future__ import annotations

import os
import shutil
import sys
from collections.abc import Callable
from dataclasses import dataclass
from pathlib import Path

from cvetopt.core.settings import Auto1PipelineConfig

LogFn = Callable[[str], None]

# Порядок как в ручной работе на листе auto1 (см. docs/Auto_new_auto1_workflow.md).
PIPELINE_STEPS: tuple[tuple[str, str], ...] = (
    ("Scan", "btnScan_Click"),
    ("Import invoice", "btnImport_Click"),
    ("Calculate", "btnCalc_Click"),
    ("Sort", "btnSort_Click"),
    ("For sklad", "btnExport2_Click"),
)

# Пути зашиты в VBA; проверяем до запуска.
VBA_INVOICE_DIR = Path(r"C:\Invoice\1")
VBA_PRICES_DIR = Path(r"C:\Invoice\1\2")


@dataclass(frozen=True)
class Auto1StepResult:
    label: str
    macro: str


def _default_log(_msg: str) -> None:
    pass


def _preflight(cfg: Auto1PipelineConfig, workbook_path: Path, log: LogFn) -> None:
    if not workbook_path.is_file():
        raise FileNotFoundError(f"Книга не найдена: {workbook_path}")
    if workbook_path.stat().st_size < 1024:
        raise RuntimeError(f"Файл слишком маленький или пустой: {workbook_path}")

    log(f"Книга: {workbook_path}")
    log(
        "Макросы ожидают папки "
        f"{VBA_INVOICE_DIR} (инвойсы) и {VBA_PRICES_DIR} (цены, один файл)."
    )
    if not VBA_INVOICE_DIR.is_dir():
        raise FileNotFoundError(f"Нет папки инвойсов: {VBA_INVOICE_DIR}")
    invoices = [
        p
        for p in VBA_INVOICE_DIR.iterdir()
        if p.is_file() and p.suffix.lower() in (".xls", ".xlsx", ".xlsm")
    ]
    if not invoices:
        raise FileNotFoundError(
            f"В {VBA_INVOICE_DIR} нет файлов .xls/.xlsx — сначала положите инвойс."
        )
    log(f"Инвойсов в папке 1: {len(invoices)} (будет выбран первый после Scan).")

    if not VBA_PRICES_DIR.is_dir():
        raise FileNotFoundError(f"Нет папки цен: {VBA_PRICES_DIR}")
    prices = [
        p
        for p in VBA_PRICES_DIR.iterdir()
        if p.is_file() and p.suffix.lower() in (".xls", ".xlsx", ".xlsm")
    ]
    if len(prices) == 0:
        raise FileNotFoundError(
            f"В {VBA_PRICES_DIR} нет файла цен — нужен ровно один .xls/.xlsx."
        )
    if len(prices) > 1:
        log(
            f"Внимание: в папке 2 файлов цен: {len(prices)} — макрос Import может заблокироваться."
        )
    else:
        log(f"Файл цен: {prices[0].name}")


def _run_sheet_macro(app: object, sheet: object, macro: str) -> None:
    codename = sheet.api.CodeName
    app.api.Run(f"'{codename}'.{macro}")


def run_auto1_pipeline(
    workbook_path: Path,
    cfg: Auto1PipelineConfig,
    *,
    log: LogFn | None = None,
) -> list[Auto1StepResult]:
    """
    Выполняет цепочку макросов на листе auto1 через Excel (Windows).
    Сохраняет книгу после завершения.
    """
    if sys.platform != "win32":
        raise RuntimeError(
            "Прогон auto1 (Scan → Import → Calculate → Sort → for sklad) "
            "доступен только на Windows с установленным Microsoft Excel."
        )

    _lg = log or _default_log
    path = workbook_path.resolve()

    _preflight(cfg, path, _lg)

    if cfg.backup_before_run:
        bak = path.with_name(path.name + cfg.backup_suffix)
        shutil.copy2(path, bak)
        _lg(f"Резервная копия: {bak.name}")

    import xlwings as xw

    visible = os.environ.get("AUTO1_EXCEL_VISIBLE", "").strip().lower() in (
        "1",
        "true",
        "yes",
    )
    _lg("Запуск Excel" + (" (окно видно — AUTO1_EXCEL_VISIBLE=1)" if visible else "…"))

    app: object | None = None
    wb: object | None = None
    done: list[Auto1StepResult] = []

    try:
        app = xw.App(visible=visible, add_book=False)
        app.display_alerts = False
        app.screen_updating = False
        app.api.EnableEvents = True

        wb = app.books.open(str(path))
        try:
            sheet = wb.sheets[cfg.sheet_name]
        except Exception as e:
            names = [s.name for s in wb.sheets]
            raise RuntimeError(
                f"Лист {cfg.sheet_name!r} не найден. Листы: {names}"
            ) from e
        sheet.activate()

        codename = sheet.api.CodeName
        _lg(f"Лист «{cfg.sheet_name}» (код VBA: {codename})")

        for label, macro in PIPELINE_STEPS:
            _lg(f"Шаг «{label}»: {macro}…")
            _run_sheet_macro(app, sheet, macro)
            done.append(Auto1StepResult(label=label, macro=macro))
            _lg(f"Шаг «{label}» завершён.")

        wb.save()
        _lg(f"Книга сохранена: {path.name}")
        _lg(
            "Выгрузка для склада: C:\\Инвойсы склад\\Голландия_1_<дата>.xlsx "
            "(и копия в C:\\Invoice\\1\\copy\\ — см. макрос btnExport2)."
        )
        return done
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
