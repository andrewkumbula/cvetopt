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
# msoAutomationSecurityLow — иначе Application.Run блокируется («макросы отключены»).
_MSO_AUTOMATION_SECURITY_LOW = 1
_RESERVED_OLE_BUTTONS = frozenset({"cbTransaction", "cbRestart", "cbCreateTimeFile"})
_ZEBRA_EVEN = 12379351
_ZEBRA_ODD = 9944773
_CHECKBOX_BMPS = (
    "Red_Check_Off.bmp",
    "Red_Check_On.bmp",
    "Green_Check_Off.bmp",
    "Green_Check_On.bmp",
)
_CV_SYNC_MACRO = "cv_SyncRowCheckboxes"
_CV_SYNC_VBA = """
Public Sub cv_SyncRowCheckboxes(aFirst As Long, aLast As Long)
    Dim aI As Long
    Dim aCommandButton As MSForms.CommandButton
    Dim aSheet As Worksheet
    Set aSheet = ThisWorkbook.Sheets(1)
    Application.ScreenUpdating = False
    Call DelCommandButton(aSheet.Name)
    With aSheet
        .Columns("A:A").ColumnWidth = 2.2
        .Columns("B:B").ColumnWidth = 2.2
        For aI = aFirst To aLast
            Set aCommandButton = .OLEObjects.Add(ClassType:="Forms.CommandButton.1").Object
            With aCommandButton
                .Left = aSheet.Cells(aI, 1).Left
                .Top = aSheet.Cells(aI, 1).Top
                .Width = aSheet.Cells(aI, 1).Width
                .Height = aSheet.Cells(aI, 1).Height
                .Picture = LoadPicture(ThisWorkbook.Path & "\\" & "Red_Check_Off.bmp")
                .Caption = "1 " & Trim(Str(aI)) & " 0"
            End With
            Set aCommandButton = .OLEObjects.Add(ClassType:="Forms.CommandButton.1").Object
            With aCommandButton
                .Left = aSheet.Cells(aI, 2).Left
                .Top = aSheet.Cells(aI, 2).Top
                .Width = aSheet.Cells(aI, 2).Width
                .Height = aSheet.Cells(aI, 2).Height
                .Picture = LoadPicture(ThisWorkbook.Path & "\\" & "Green_Check_Off.bmp")
                .Caption = "2 " & Trim(Str(aI)) & " 0"
            End With
            If (aI Mod 2) = 0 Then
                .Range("D" & Trim(Str(aI)) & ":AB" & Trim(Str(aI))).Interior.Color = 12379351
            Else
                .Range("D" & Trim(Str(aI)) & ":AB" & Trim(Str(aI))).Interior.Color = 9944773
            End If
        Next aI
    End With
End Sub
"""


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
    api.EnableEvents = True
    try:
        api.AutomationSecurity = _MSO_AUTOMATION_SECURITY_LOW
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


def _copy_checkbox_assets(template_dir: Path, target_dir: Path) -> None:
    for name in _CHECKBOX_BMPS:
        src = template_dir / name
        if src.is_file():
            shutil.copy2(src, target_dir / name)


def _delete_row_command_buttons(sheet: object) -> None:
    oles = sheet.api.OLEObjects()
    for i in range(int(oles.Count), 0, -1):
        ole = oles.Item(i)
        if str(ole.Name) in _RESERVED_OLE_BUTTONS:
            continue
        try:
            ole.Delete()
        except Exception:
            pass


def _sync_row_checkboxes_com(
    wb: object,
    sheet: object,
    *,
    first_row: int,
    last_row: int,
    assets_dir: Path,
) -> None:
    """Запасной путь без Application.Run — как VBA SetCommandButton (один ClassType)."""
    red_img = str((assets_dir / "Red_Check_Off.bmp").resolve())
    green_img = str((assets_dir / "Green_Check_Off.bmp").resolve())
    if not Path(red_img).is_file() or not Path(green_img).is_file():
        raise FileNotFoundError(
            f"Нет bmp для чекбоксов в {assets_dir} (Red_Check_Off.bmp, Green_Check_Off.bmp)."
        )

    app_api = wb.app.api
    sheet.api.Columns("A:A").ColumnWidth = 2.2
    sheet.api.Columns("B:B").ColumnWidth = 2.2
    _delete_row_command_buttons(sheet)

    for row in range(first_row, last_row + 1):
        for col, img, prefix in ((1, red_img, "1"), (2, green_img, "2")):
            cell = sheet.api.Cells(row, col)
            ole = sheet.api.OLEObjects().Add("Forms.CommandButton.1")
            btn = ole.Object
            btn.Left = float(cell.Left)
            btn.Top = float(cell.Top)
            btn.Width = float(cell.Width)
            btn.Height = float(cell.Height)
            btn.Picture = app_api.LoadPicture(img)
            btn.Caption = f"{prefix} {row} 0"

        color = _ZEBRA_EVEN if row % 2 == 0 else _ZEBRA_ODD
        sheet.api.Range(f"D{row}:AB{row}").Interior.Color = color


def _ensure_cv_sync_macro(wb: object) -> None:
    """Public-макрос в Module1 (как SetCommandButton, без вставки колонки A)."""
    mod = wb.api.VBProject.VBComponents("Module1").CodeModule
    line_count = int(mod.CountOfLines)
    existing = mod.Lines(1, line_count) if line_count else ""
    if f"Sub {_CV_SYNC_MACRO}" in existing:
        return
    mod.InsertLines(line_count + 1, _CV_SYNC_VBA)


def _sync_row_checkboxes(
    wb: object,
    sheet: object,
    *,
    first_row: int,
    last_row: int,
    assets_dir: Path,
    log: LogFn | None = None,
) -> None:
    """Чекбоксы: сначала VBA (как в шаблоне), при блокировке макросов — COM."""
    if last_row < first_row:
        return

    def _lg(msg: str) -> None:
        if log is not None:
            log(msg)

    _lg(f"Эквадор: чекбоксы для строк {first_row}–{last_row}…")

    vba_errors: list[str] = []
    try:
        _ensure_cv_sync_macro(wb)
        app = wb.app.api
        wb_name = str(wb.name)
        for spec in (
            f"'{wb_name}'!{_CV_SYNC_MACRO}",
            f"Module1.{_CV_SYNC_MACRO}",
            _CV_SYNC_MACRO,
        ):
            try:
                app.Run(spec, first_row, last_row)
                return
            except Exception as e:
                vba_errors.append(f"{spec}: {e}")
    except Exception as e:
        vba_errors.append(f"inject: {e}")

    _lg("Эквадор: VBA недоступен, чекбоксы через COM…")
    try:
        _sync_row_checkboxes_com(
            wb,
            sheet,
            first_row=first_row,
            last_row=last_row,
            assets_dir=assets_dir,
        )
    except Exception as com_err:
        tail = "; ".join(vba_errors[-2:])
        raise RuntimeError(
            "Не удалось создать чекбоксы (VBA и COM). "
            "Проверьте: макросы разрешены для автоматизации, папка шаблона в доверенном "
            "расположении Excel, bmp Red_Check/Green_Check рядом с книгой. "
            f"VBA: {tail}. COM: {com_err}"
        ) from com_err


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
        _copy_checkbox_assets(template.parent, tmp_copy.parent)

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
        last_row = _DATA_FIRST_ROW + len(deals) - 1
        for idx, deal in enumerate(deals):
            _write_deal_row(data_sheet, _DATA_FIRST_ROW + idx, deal)

        path_sheet.range("A1").value = str(biflorica_path)
        path_sheet.range("B1").value = biflorica_path.name

        _sync_row_checkboxes(
            wb,
            data_sheet,
            first_row=_DATA_FIRST_ROW,
            last_row=last_row,
            assets_dir=tmp_copy.parent,
            log=log,
        )

        _lg(f"Эквадор: сохраняю → {out_path}")
        _apply_create_file_ui(wb, out_name)
        wb.api.SaveAs(str(out_path))
        _copy_checkbox_assets(template.parent, out_path.parent)
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
