from __future__ import annotations

import os
import shutil
import sys
import tempfile
from collections.abc import Callable
from datetime import datetime
from pathlib import Path

from cvetopt.core.runtime_settings import (
    resolve_ecuador_output_dir,
    resolve_ecuador_template,
)
from cvetopt.core.settings import EnvSettings
from cvetopt.invoice.ecuador_transform import EcuadorDealRow, transform_biflorica_deals

LogFn = Callable[[str], None]

_SHEET_DATA = 0
_SHEET_PATH = 1
_DATA_FIRST_ROW = 7
_CLEAR_LAST_ROW = 500
# msoAutomationSecurityForceDisable — без диалога макросов при открытии .xlsm
_MSO_AUTOMATION_SECURITY_FORCE_DISABLE = 3


def ecuador_output_basename(when: datetime | None = None) -> str:
    """
    Имя файла как в кнопке «Создать файл», но без «/» в дате —
    иначе Windows создаёт подпапки (Эквадор 03\\06\\…).
    """
    t = when or datetime.now()
    stamp = t.strftime("%d.%m.%y %H.%M")
    return f"Эквадор {stamp}.xlsm"


def _configure_excel_app(app: object) -> None:
    app.display_alerts = False
    app.screen_updating = False
    api = app.api
    api.DisplayAlerts = False
    api.EnableEvents = False
    try:
        api.AutomationSecurity = _MSO_AUTOMATION_SECURITY_FORCE_DISABLE
    except Exception:
        pass


def _write_deal_row(sheet: object, row: int, deal: EcuadorDealRow) -> None:
    sheet.range((row, 4)).value = deal.plantation
    sheet.range((row, 5)).value = deal.flower_type
    sheet.range((row, 6)).value = deal.variety
    sheet.range((row, 15)).value = deal.boxes
    sheet.range((row, 16)).value = deal.sm
    sheet.range((row, 17)).value = deal.box_type
    sheet.range((row, 18)).value = deal.total_stems
    for col_letter, value in deal.qty_by_length_col.items():
        if value:
            sheet.range(f"{col_letter}{row}").value = value


def _apply_create_file_ui(workbook: object, output_name: str) -> None:
    """Аналог cbCreateTimeFile: лист «Форматирование», флаги на Path."""
    sheet = workbook.sheets[_SHEET_DATA]
    try:
        sheet.name = "Форматирование"
    except Exception:
        pass
    for btn_name, visible in (("cbCreateTimeFile", False), ("cbTransaction", True)):
        try:
            sheet.api.OLEObjects(btn_name).Object.Visible = visible
        except Exception:
            pass
    try:
        workbook.sheets[_SHEET_PATH].range("A2").value = "False"
    except Exception:
        pass
    _ = output_name


def create_ecuador_file_from_biflorica(
    biflorica_path: Path,
    env: EnvSettings,
    *,
    template_path: Path | None = None,
    output_dir: Path | None = None,
    log: LogFn | None = None,
) -> Path:
    """
    Вариант B: преобразование в Python, запись в шаблон .xlsm, SaveAs как «Создать файл».
    Только Windows + установленный Excel (xlwings).
    """
    if sys.platform != "win32":
        raise RuntimeError("Создание файла Эквадор доступно только на Windows с Excel.")

    biflorica_path = biflorica_path.resolve()
    if not biflorica_path.is_file():
        raise FileNotFoundError(biflorica_path)

    from cvetopt.core.runtime_settings import load_runtime_settings

    runtime = load_runtime_settings(env)
    template = template_path or resolve_ecuador_template(env, runtime.ecuador_template_path)
    out_dir = output_dir or resolve_ecuador_output_dir(env, runtime.ecuador_output_dir)

    def _lg(msg: str) -> None:
        if log is not None:
            log(msg)

    _lg("Эквадор: проверка шаблона…")
    if not template.is_file() or template.stat().st_size < 1024:
        raise FileNotFoundError(f"Шаблон обработки не найден: {template}")
    if template.read_bytes()[:2] != b"PK":
        raise RuntimeError(f"Шаблон повреждён или не скачан полностью: {template}")

    _lg("Эквадор: разбор xlsx Biflorica…")
    deals = transform_biflorica_deals(biflorica_path)
    if not deals:
        raise ValueError(f"В отчёте нет строк сделок: {biflorica_path.name}")

    out_dir.mkdir(parents=True, exist_ok=True)
    out_name = ecuador_output_basename()
    out_path = (out_dir / out_name).resolve()

    _lg(f"Эквадор: сделок {len(deals)}, шаблон {template.name}")

    import xlwings as xw

    visible = os.environ.get("ECUADOR_EXCEL_VISIBLE", "").strip().lower() in (
        "1",
        "true",
        "yes",
    )
    _lg(
        "Эквадор: запуск Excel"
        + (" (окно видно — ECUADOR_EXCEL_VISIBLE=1)" if visible else "…")
    )

    app: object | None = None
    wb: object | None = None
    tmp_copy: Path | None = None
    try:
        app = xw.App(visible=visible, add_book=False)
        _configure_excel_app(app)

        with tempfile.NamedTemporaryFile(suffix=".xlsm", delete=False) as tmp:
            tmp_copy = Path(tmp.name)
        shutil.copy2(template, tmp_copy)

        _lg("Эквадор: открываю шаблон (без запроса макросов)…")
        wb = app.books.open(
            str(tmp_copy),
            update_links=0,
            read_only=False,
            ignore_read_only_recommended=True,
        )
        data_sheet = wb.sheets[_SHEET_DATA]
        path_sheet = wb.sheets[_SHEET_PATH]

        _lg("Эквадор: заполняю строки…")
        data_sheet.range(f"D{_DATA_FIRST_ROW}:AB{_CLEAR_LAST_ROW}").clear_contents()
        for idx, deal in enumerate(deals):
            _write_deal_row(data_sheet, _DATA_FIRST_ROW + idx, deal)

        path_sheet.range("A1").value = str(biflorica_path)
        path_sheet.range("B1").value = biflorica_path.name

        _lg(f"Эквадор: сохраняю → {out_path}")
        _apply_create_file_ui(wb, out_name)
        wb.api.SaveAs(str(out_path))
        wb.save()
        wb.close()
        wb = None
        _lg(f"Эквадор: файл создан → {out_path}")
        return out_path
    finally:
        if wb is not None:
            try:
                wb.close()
            except Exception:
                pass
        if tmp_copy is not None:
            tmp_copy.unlink(missing_ok=True)
        if app is not None:
            try:
                app.quit()
            except Exception:
                pass
