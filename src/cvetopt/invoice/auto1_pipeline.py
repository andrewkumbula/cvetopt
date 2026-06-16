from __future__ import annotations

import os
import shutil
import sys
import threading
import time
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
VBA_COPY_DIR = Path(r"C:\Invoice\1\copy")
DEFAULT_SKLAD_EXPORT_DIR = Path(r"C:\Инвойсы склад")

# auto1: данные с 8-й строки; Sort — A:BB по Description (F).
_AUTO1_DATA_FIRST_ROW = 8
_AUTO1_LAST_COL = 54  # BB
_COL_DESCRIPTION = 6  # F
_PASTE_VALUE_COL_RANGES = ((16, 17), (48, 54))  # P:Q, AV:BB
_XL_UP = -4162
_XL_PASTE_VALUES = -4163
_XL_ASCENDING = 1
_XL_SORT_NO_HEADER = 2


@dataclass(frozen=True)
class Auto1StepResult:
    label: str
    macro: str


def _default_log(_msg: str) -> None:
    pass


def _close_existing_excel(log: LogFn) -> None:
    """Снимает зависшие EXCEL.EXE от прошлых прогонов (блокируют SaveAs в btnExport2)."""
    import xlwings as xw

    apps = list(xw.apps)
    if not apps:
        return
    log(f"Закрываем {len(apps)} старый(х) Excel перед auto1…")
    for app in apps:
        try:
            app.quit()
        except Exception:
            pass
    time.sleep(1)


def _preflight(
    cfg: Auto1PipelineConfig,
    workbook_path: Path,
    sklad_export_dir: Path,
    log: LogFn,
) -> None:
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

    sklad_export_dir.mkdir(parents=True, exist_ok=True)
    VBA_COPY_DIR.mkdir(parents=True, exist_ok=True)
    log(f"Папка выгрузки: {sklad_export_dir}")
    log(f"Копия для склада: {VBA_COPY_DIR}")


CV_RUNNER_PREFIX = "cv_Run_"


def _runner_name(macro: str) -> str:
    return f"{CV_RUNNER_PREFIX}{macro}"


def _ensure_sheet_runner(wb: object, codename: str, macro: str) -> str:
    """
    Public-обёртка в модуле листа (тот же модуль, что Private btn*_Click).
    Требует «Доверять доступ к объектной модели VBA» в Excel.
    """
    runner = _runner_name(macro)
    mod = wb.api.VBProject.VBComponents(codename).CodeModule
    line_count = int(mod.CountOfLines)
    existing = mod.Lines(1, line_count) if line_count else ""
    if f"Sub {runner}" in existing:
        return runner
    proc = f"\r\nPublic Sub {runner}()\r\n    {macro}\r\nEnd Sub\r\n"
    mod.InsertLines(line_count + 1, proc)
    return runner


def _try_run(app: object, spec: str) -> None:
    app.api.Run(spec)


def _prepare_macro_step(app: object) -> None:
    """Private-макросы и сортировка часто «зависают» при ScreenUpdating=False."""
    app.screen_updating = True
    api = app.api
    api.DisplayAlerts = False
    try:
        api.EnableEvents = True
    except Exception:
        pass


def _run_with_heartbeat(label: str, action: Callable[[], None], log: LogFn) -> None:
    stop = threading.Event()

    def _beat() -> None:
        tick = 0
        while not stop.wait(20):
            tick += 20
            log(f"  … шаг «{label}» выполняется ({tick} с)")

    worker = threading.Thread(target=_beat, daemon=True)
    worker.start()
    try:
        action()
    finally:
        stop.set()


def _auto1_last_row(sheet: object) -> int:
    return int(
        sheet.api.Cells(sheet.api.Rows.Count, _COL_DESCRIPTION).End(_XL_UP).Row
    )


def _sort_auto1_sheet(sheet: object, log: LogFn) -> None:
    """
    Эквивалент btnSort_Click без VBA: формулы P:Q и AV:BB → значения, сортировка A:BB по F.
    btnSort_Click на сервере часто зависает (PasteSpecial/Sortirovka в .xls).
    """
    last_row = _auto1_last_row(sheet)
    if last_row < _AUTO1_DATA_FIRST_ROW:
        log("Sort: нет строк данных — пропуск")
        return
    log(f"Sort: P:Q и AV:BB → значения, сортировка F, строки {_AUTO1_DATA_FIRST_ROW}–{last_row}")
    api = sheet.api
    app_api = api.Application
    block = api.Range(
        api.Cells(_AUTO1_DATA_FIRST_ROW, 1),
        api.Cells(last_row, _AUTO1_LAST_COL),
    )
    for col_start, col_end in _PASTE_VALUE_COL_RANGES:
        part = api.Range(
            api.Cells(_AUTO1_DATA_FIRST_ROW, col_start),
            api.Cells(last_row, col_end),
        )
        part.Copy()
        part.PasteSpecial(Paste=_XL_PASTE_VALUES)
    app_api.CutCopyMode = False
    key = api.Range(
        api.Cells(_AUTO1_DATA_FIRST_ROW, _COL_DESCRIPTION),
        api.Cells(last_row, _COL_DESCRIPTION),
    )
    block.Sort(Key1=key, Order1=_XL_ASCENDING, Header=_XL_SORT_NO_HEADER)
    log("Sort: готово (Python)")


def _prepare_sklad_export(app: object, wb: object, log: LogFn) -> None:
    """btnExport2 создаёт новую книгу и SaveAs — без папок/экрана часто «висит» с диалогом."""
    api = app.api
    app.screen_updating = True
    api.DisplayAlerts = False
    try:
        api.AskToUpdateLinks = False
    except Exception:
        pass
    try:
        api.Calculation = -4135  # xlCalculationAutomatic
    except Exception:
        pass
    wb_name = str(wb.name)
    for book in list(app.books):
        try:
            name = str(book.name)
        except Exception:
            continue
        if name == wb_name:
            continue
        if name.startswith("Голландия_1") or name.startswith("copy_Голландия"):
            log(f"  закрываем лишнюю книгу: {name}")
            try:
                book.close(SaveChanges=False)
            except Exception:
                pass


def _on_action_matches_macro(on_action: str, macro: str) -> bool:
    text = on_action.strip()
    if not text:
        return False
    if text == macro:
        return True
    return text.endswith(f".{macro}") or text.endswith(f"!{macro}")


def _run_via_form_on_action(app: object, sheet: object, macro: str) -> str | None:
    """Кнопки Form Control с OnAction (запасной путь)."""
    for i in range(1, int(sheet.api.Shapes.Count) + 1):
        shp = sheet.api.Shapes(i)
        try:
            on_action = (shp.OnAction or "").strip()
        except Exception:
            on_action = ""
        if not _on_action_matches_macro(on_action, macro):
            continue
        _try_run(app, on_action)
        return on_action
    return None


def _invoke_sheet_click_macro(
    app: object,
    wb: object,
    sheet: object,
    macro: str,
    log: LogFn,
) -> None:
    """
    btn*_Click в Auto_new.xls — Private Sub в модуле листа (Лист1.cls).
    Application.Run("'Лист1'.btnScan_Click") их не видит; в списке макросов (Alt+F8)
    только Public вроде ActualCurs1.
    """
    codename = str(sheet.api.CodeName)
    sheet_name = str(sheet.name)
    wb_name = str(wb.name)
    errors: list[str] = []

    candidates = (
        f"{codename}.{macro}",
        f"'{codename}'.{macro}",
        f"'{sheet_name}'!{macro}",
        f"'{codename}'!{macro}",
        f"'{wb_name}'!{macro}",
        macro,
    )
    for spec in candidates:
        try:
            log(f"  → Application.Run({spec!r})…")
            _try_run(app, spec)
            log(f"  → Application.Run({spec!r}) готов")
            return
        except Exception as e:
            errors.append(f"{spec}: {e!s}")

    try:
        runner = _ensure_sheet_runner(wb, codename, macro)
        spec = f"'{codename}'.{runner}"
        log(f"  → Application.Run({spec!r})…")
        _try_run(app, spec)
        log(f"  → обёртка {runner} в модуле {codename}")
        return
    except Exception as e:
        errors.append(f"обёртка VBA: {e!s}")

    try:
        via = _run_via_form_on_action(app, sheet, macro)
        if via:
            log(f"  → OnAction: {via}")
            return
    except Exception as e:
        errors.append(f"OnAction: {e}")

    hint = (
        "Обработчики кнопок (btnScan_Click и др.) в книге объявлены как Private Sub. "
        "Включите в Excel: Файл → Параметры → Центр управления безопасностью → "
        "Параметры макросов → «Доверять доступ к объектной модели VBA-проекта» "
        "(Trust access to the VBA project object model), затем повторите прогон."
    )
    tail = "; ".join(errors[-4:]) if errors else "нет деталей"
    raise RuntimeError(f"Не удалось вызвать {macro}. {hint} ({tail})")


def run_auto1_pipeline(
    workbook_path: Path,
    cfg: Auto1PipelineConfig,
    *,
    sklad_export_dir: Path | None = None,
    holland_marker_assets_dir: Path | None = None,
    add_holland_row_markers: bool = False,
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
    export_dir = (sklad_export_dir or DEFAULT_SKLAD_EXPORT_DIR).resolve()

    _preflight(cfg, path, export_dir, _lg)
    _close_existing_excel(_lg)

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
        app.screen_updating = True
        app.api.EnableEvents = True
        # 1 = msoAutomationSecurityLow — иначе Run может блокироваться политикой.
        try:
            app.api.AutomationSecurity = 1
        except Exception:
            pass

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
            step_t0 = time.monotonic()
            _prepare_macro_step(app)
            if label == "Sort":
                _lg("Шаг «Sort»: сортировка по Description (Python)…")
                _run_with_heartbeat(
                    label,
                    lambda: _sort_auto1_sheet(sheet, _lg),
                    _lg,
                )
            else:
                _lg(f"Шаг «{label}»: {macro}…")
                if label == "For sklad":
                    _prepare_sklad_export(app, wb, _lg)
                    _lg(
                        "Экспорт для склада (btnExport2) — обычно 30–120 с; "
                        "если дольше 5 мин — откройте Excel (AUTO1_EXCEL_VISIBLE=1) "
                        "или остановите прогон."
                    )
                _run_with_heartbeat(
                    label,
                    lambda m=macro: _invoke_sheet_click_macro(app, wb, sheet, m, _lg),
                    _lg,
                )
            done.append(
                Auto1StepResult(
                    label=label,
                    macro="python_sort" if label == "Sort" else macro,
                )
            )
            _lg(f"Шаг «{label}» завершён ({time.monotonic() - step_t0:.1f} с).")
            if label == "For sklad":
                if add_holland_row_markers and holland_marker_assets_dir is not None:
                    from cvetopt.invoice.holland_markers import finalize_holland_after_auto1

                    marked = finalize_holland_after_auto1(
                        app,
                        export_dir,
                        holland_marker_assets_dir,
                        _lg,
                        auto1_sheet_name=cfg.sheet_name,
                    )
                    if marked is not None:
                        _lg(f"Голландия: готово → {marked.name}")

        wb.save()
        _lg(f"Книга сохранена: {path.name}")
        _lg(
            "Выгрузка для склада: C:\\Инвойсы склад\\Голландия_1_<дата>.xlsm "
            "(макрос btnExport2 + маркеры A–B при включённой опции)."
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
