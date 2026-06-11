from __future__ import annotations

import sys
from collections.abc import Callable
from pathlib import Path

from cvetopt.invoice.ecuador_create import (
    _CHECKBOX_BMPS,
    _copy_checkbox_assets,
    _delete_row_command_buttons,
    _load_picture,
)
from cvetopt.invoice.xlsx_read import grid_by_row, read_xlsx_grid

LogFn = Callable[[str], None]

_HOLLAND_DATA_FIRST_ROW = 2
_MSO_AUTOMATION_SECURITY_LOW = 1
_ZEBRA_EVEN = 12379351
_ZEBRA_ODD = 9944773

_CV_SYNC_MACRO = "cv_SyncHollandMarkers"
_CV_WIRE_MACRO = "cv_WireHollandMarkerButtons"
_MARKER_MODULE = "cvHollandMarkers"
_MARKER_CLASS = "cvHollandMarkerHandler"

_CV_SYNC_VBA = """
Public Sub cv_SyncHollandMarkers(aFirst As Long, aLast As Long)
    Dim aI As Long
    Dim aCommandButton As MSForms.CommandButton
    Dim aSheet As Worksheet
    Dim aLastCol As Long
    Set aSheet = ThisWorkbook.Worksheets(1)
    Application.ScreenUpdating = False
    Call cvDelHollandMarkerButtons(aSheet.Name)
    aLastCol = aSheet.Range("C1").End(xlToRight).Column
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
                .Range(.Cells(aI, 3), .Cells(aI, aLastCol)).Interior.Color = 12379351
            Else
                .Range(.Cells(aI, 3), .Cells(aI, aLastCol)).Interior.Color = 9944773
            End If
        Next aI
    End With
    Call cv_WireHollandMarkerButtons
End Sub

Public Sub cvDelHollandMarkerButtons(aSheet As String)
    Dim aButton As OLEObject
    For Each aButton In ThisWorkbook.Sheets(aSheet).OLEObjects
        If TypeOf aButton.Object Is MSForms.CommandButton Then
            aButton.Delete
        End If
    Next
End Sub

Public ColHollandButtons As Collection

Public Sub cv_WireHollandMarkerButtons()
    Dim aButton As OLEObject
    Dim h As cvHollandMarkerHandler
    On Error Resume Next
    Set ColHollandButtons = New Collection
    For Each aButton In ThisWorkbook.Worksheets(1).OLEObjects
        If TypeOf aButton.Object Is MSForms.CommandButton Then
            Set h = New cvHollandMarkerHandler
            Set h.EventButton = aButton.Object
            ColHollandButtons.Add h
        End If
    Next
End Sub
"""

_MARKER_CLASS_VBA = """
Public WithEvents EventButton As MSForms.CommandButton

Private Function _holland_marker_bmp(name As String) As String
    _holland_marker_bmp = ThisWorkbook.Path & "\\" & name
End Function

Private Sub EventButton_Click()
    Dim aStr As String
    Dim aAddress As String
    Dim aSheet As Worksheet
    Set aSheet = ThisWorkbook.Worksheets(1)
    With EventButton
        If Left(.Caption, 1) = "1" Then
            If Right(.Caption, 1) = "0" Then
                .Picture = LoadPicture(_holland_marker_bmp("Red_Check_On.bmp"))
                aStr = .Caption
                aStr = Right(aStr, Len(aStr) - 2)
                aStr = Left(aStr, Len(aStr) - 2)
                aAddress = aSheet.Cells(CInt(aStr), aSheet.Range("C1").End(xlToRight).Column). _
                    Address(RowAbsolute:=False, ColumnAbsolute:=False)
                aSheet.Range("C" & aStr & ":" & aAddress).Interior.Color = RGB(255, 209, 209)
                .Caption = Left(.Caption, Len(.Caption) - 2) & " 1"
            Else
                .Picture = LoadPicture(_holland_marker_bmp("Red_Check_Off.bmp"))
                aStr = .Caption
                aStr = Right(aStr, Len(aStr) - 2)
                aStr = Left(aStr, Len(aStr) - 2)
                aAddress = aSheet.Cells(CInt(aStr), aSheet.Range("C1").End(xlToRight).Column). _
                    Address(RowAbsolute:=False, ColumnAbsolute=False)
                aSheet.Range("C" & aStr & ":" & aAddress).Interior.Color = RGB(255, 255, 255)
                .Caption = Left(.Caption, Len(.Caption) - 2) & " 0"
            End If
        ElseIf Left(.Caption, 1) = "2" Then
            If Right(.Caption, 1) = "0" Then
                .Picture = LoadPicture(_holland_marker_bmp("Green_Check_On.bmp"))
                aStr = .Caption
                aStr = Right(aStr, Len(aStr) - 2)
                aStr = Left(aStr, Len(aStr) - 2)
                aAddress = aSheet.Cells(CInt(aStr), aSheet.Range("C1").End(xlToRight).Column). _
                    Address(RowAbsolute:=False, ColumnAbsolute:=False)
                aSheet.Range("C" & aStr & ":" & aAddress).Interior.Color = RGB(0, 255, 0)
                .Caption = Left(.Caption, Len(.Caption) - 2) & " 1"
            Else
                .Picture = LoadPicture(_holland_marker_bmp("Green_Check_Off.bmp"))
                aStr = .Caption
                aStr = Right(aStr, Len(aStr) - 2)
                aStr = Left(aStr, Len(aStr) - 2)
                aAddress = aSheet.Cells(CInt(aStr), aSheet.Range("C1").End(xlToRight).Column). _
                    Address(RowAbsolute:=False, ColumnAbsolute:=False)
                aSheet.Range("C" & aStr & ":" & aAddress).Interior.Color = RGB(255, 255, 255)
                .Caption = Left(.Caption, Len(.Caption) - 2) & " 0"
            End If
        End If
    End With
End Sub
"""

_SHEET_ACTIVATE_VBA = """
Private Sub Worksheet_Activate()
    Application.Run "cv_WireHollandMarkerButtons"
End Sub
"""


def _default_log(_msg: str) -> None:
    pass


def _last_data_row_xlsx(export_path: Path) -> int:
    grid = read_xlsx_grid(export_path)
    rows = grid_by_row(grid)
    last = _HOLLAND_DATA_FIRST_ROW - 1
    for row_no, cells in rows.items():
        if row_no < _HOLLAND_DATA_FIRST_ROW:
            continue
        if any(str(v).strip() for v in cells.values()):
            last = max(last, row_no)
    return last


def _ensure_std_module(vbproject: object, name: str, code: str) -> None:
    try:
        mod = vbproject.VBComponents(name)
    except Exception:
        mod = vbproject.VBComponents.Add(1)
        mod.Name = name
    code_module = mod.CodeModule
    existing = code_module.Lines(1, code_module.CountOfLines) if code_module.CountOfLines else ""
    if _CV_SYNC_MACRO in existing and "cvDelHollandMarkerButtons" in existing:
        return
    if code_module.CountOfLines:
        code_module.InsertLines(code_module.CountOfLines + 1, code)
    else:
        code_module.AddFromString(code)


def _ensure_class_module(vbproject: object, name: str, code: str) -> None:
    try:
        mod = vbproject.VBComponents(name)
    except Exception:
        mod = vbproject.VBComponents.Add(2)
        mod.Name = name
    code_module = mod.CodeModule
    if code_module.CountOfLines and "EventButton_Click" in code_module.Lines(1, code_module.CountOfLines):
        return
    if code_module.CountOfLines:
        code_module.DeleteLines(1, code_module.CountOfLines)
    code_module.AddFromString(code)


def _ensure_sheet_activate(vbproject: object, sheet_codename: str) -> None:
    mod = vbproject.VBComponents(sheet_codename)
    code_module = mod.CodeModule
    existing = code_module.Lines(1, code_module.CountOfLines) if code_module.CountOfLines else ""
    if "cv_WireHollandMarkerButtons" in existing:
        return
    if code_module.CountOfLines:
        code_module.InsertLines(code_module.CountOfLines + 1, _SHEET_ACTIVATE_VBA)
    else:
        code_module.AddFromString(_SHEET_ACTIVATE_VBA)


def _missing_marker_assets(assets_dir: Path) -> list[str]:
    return [name for name in _CHECKBOX_BMPS if not (assets_dir / name).is_file()]


def _holland_last_col(sheet: object) -> int:
    return int(sheet.api.Range("C1").End(-4161).Column)  # xlToRight


def _count_marker_buttons(sheet: object) -> int:
    count = 0
    oles = sheet.api.OLEObjects()
    for i in range(1, int(oles.Count) + 1):
        try:
            if "CommandButton" in str(oles.Item(i).ClassType):
                count += 1
        except Exception:
            pass
    return count


def _marker_columns_already(ws: object) -> bool:
    """Колонки A–B уже вставлены (повторный прогон)."""
    try:
        header = str(ws.range("C1").value or "").strip().casefold()
        if header.startswith("box"):
            return True
    except Exception:
        pass
    return False


def _sync_holland_markers_com(
    wb: object,
    sheet: object,
    *,
    first_row: int,
    last_row: int,
    assets_dir: Path,
) -> None:
    red_img = str((assets_dir / "Red_Check_Off.bmp").resolve())
    green_img = str((assets_dir / "Green_Check_Off.bmp").resolve())
    if not Path(red_img).is_file() or not Path(green_img).is_file():
        raise FileNotFoundError(f"Нет bmp в {assets_dir}")

    app_api = wb.app.api
    sheet.api.Columns("A:A").ColumnWidth = 2.2
    sheet.api.Columns("B:B").ColumnWidth = 2.2
    _delete_row_command_buttons(sheet)
    last_col = _holland_last_col(sheet)

    for row in range(first_row, last_row + 1):
        for col, img, prefix in ((1, red_img, "1"), (2, green_img, "2")):
            cell = sheet.api.Cells(row, col)
            ole = sheet.api.OLEObjects().Add(
                "Forms.CommandButton.1",
                "",
                False,
                False,
                float(cell.Left),
                float(cell.Top),
                float(cell.Width),
                float(cell.Height),
            )
            btn = ole.Object
            btn.Picture = _load_picture(app_api, img)
            btn.Caption = f"{prefix} {row} 0"

        color = _ZEBRA_EVEN if row % 2 == 0 else _ZEBRA_ODD
        sheet.api.Range(
            sheet.api.Cells(row, 3),
            sheet.api.Cells(row, last_col),
        ).Interior.Color = color


def _inject_marker_vba(wb: object, sheet_codename: str) -> None:
    vb = wb.api.VBProject
    _ensure_std_module(vb, _MARKER_MODULE, _CV_SYNC_VBA)
    _ensure_class_module(vb, _MARKER_CLASS, _MARKER_CLASS_VBA)
    _ensure_sheet_activate(vb, sheet_codename)


def _wire_marker_clicks(app: object, wb: object, *, log: LogFn) -> None:
    specs = (
        f"{_MARKER_MODULE}.{_CV_WIRE_MACRO}",
        _CV_WIRE_MACRO,
        f"'{wb.name}'!{_CV_WIRE_MACRO}",
    )
    last_err: Exception | None = None
    for spec in specs:
        try:
            app.api.Run(spec)
            log("Голландия: обработчики кликов подключены (VBA).")
            return
        except Exception as e:
            last_err = e
    log(f"Голландия: клики по маркерам — без VBA ({last_err})")


def _sync_holland_markers_vba(
    app: object,
    wb: object,
    *,
    first_row: int,
    last_row: int,
    sheet_codename: str,
    log: LogFn,
) -> bool:
    _inject_marker_vba(wb, sheet_codename)
    specs = (
        f"{_MARKER_MODULE}.{_CV_SYNC_MACRO}",
        _CV_SYNC_MACRO,
        f"'{wb.name}'!{_CV_SYNC_MACRO}",
    )
    for spec in specs:
        try:
            app.api.Run(spec, first_row, last_row)
            log("Голландия: маркеры созданы (VBA).")
            return True
        except Exception as e:
            last_err = e
    log(f"Голландия: VBA маркеры не запустились — {last_err}")
    return False


def add_holland_row_markers(
    export_path: Path,
    assets_dir: Path,
    *,
    log: LogFn | None = None,
) -> Path:
    """
    Добавляет слева два столбца с красной/зелёной кнопкой (как в Эквадор).
    Сохраняет книгу как .xlsm с VBA; исходный .xlsx удаляется.
    """
    _lg = log or _default_log
    if sys.platform != "win32":
        raise RuntimeError("Маркеры Голландия: нужен Windows + Excel.")

    export_path = export_path.resolve()
    if not export_path.is_file():
        raise FileNotFoundError(export_path)
    if export_path.suffix.lower() not in (".xlsx", ".xlsm"):
        raise ValueError(f"Ожидался .xlsx/.xlsm: {export_path.name}")

    missing = _missing_marker_assets(assets_dir)
    if missing:
        raise FileNotFoundError(
            f"Нет bmp для маркеров в {assets_dir}: {', '.join(missing)} "
            "(скопируйте из папки шаблона Эквадор)."
        )

    last_row = _last_data_row_xlsx(export_path)
    if last_row < _HOLLAND_DATA_FIRST_ROW:
        _lg("Голландия: маркеры пропущены — нет строк данных.")
        return export_path

    assets_target = export_path.parent
    _copy_checkbox_assets(assets_dir, assets_target)
    xlsm_path = export_path.with_suffix(".xlsm")

    import xlwings as xw

    app: object | None = None
    wb: object | None = None
    try:
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        app.api.AutomationSecurity = _MSO_AUTOMATION_SECURITY_LOW
        wb = app.books.open(str(export_path), update_links=False)
        ws = wb.sheets[0]
        sheet_codename = str(ws.api.CodeName)

        need_insert = export_path.suffix.lower() == ".xlsx" and not _marker_columns_already(ws)
        if need_insert:
            ws.api.Columns("A:B").Insert()
        elif _marker_columns_already(ws):
            _lg("Голландия: колонки A–B уже есть — только маркеры.")

        if export_path.suffix.lower() == ".xlsx":
            wb.api.SaveAs(str(xlsm_path), FileFormat=52)
        else:
            xlsm_path = export_path

        assets_target = xlsm_path.parent
        _copy_checkbox_assets(assets_dir, assets_target)

        vba_ok = False
        try:
            vba_ok = _sync_holland_markers_vba(
                app,
                wb,
                first_row=_HOLLAND_DATA_FIRST_ROW,
                last_row=last_row,
                sheet_codename=sheet_codename,
                log=_lg,
            )
        except Exception as e:
            _lg(f"Голландия: VBA недоступен ({e})")

        if not vba_ok or _count_marker_buttons(ws) < (last_row - _HOLLAND_DATA_FIRST_ROW + 1):
            _lg("Голландия: маркеры через COM…")
            _sync_holland_markers_com(
                wb,
                ws,
                first_row=_HOLLAND_DATA_FIRST_ROW,
                last_row=last_row,
                assets_dir=assets_target,
            )

        btn_count = _count_marker_buttons(ws)
        expected = (last_row - _HOLLAND_DATA_FIRST_ROW + 1) * 2
        if btn_count < expected:
            raise RuntimeError(
                f"Создано кнопок {btn_count} из {expected}. "
                "Проверьте bmp рядом с файлом и доверие VBA в Excel."
            )

        try:
            _inject_marker_vba(wb, sheet_codename)
            _wire_marker_clicks(app, wb, log=_lg)
        except Exception as e:
            _lg(f"Голландия: клики по маркерам могут не работать — {e}")

        wb.save()
        _lg(
            f"Голландия: маркеры {btn_count} кн., строки "
            f"{_HOLLAND_DATA_FIRST_ROW}–{last_row} → {xlsm_path.name}"
        )
        if export_path.suffix.lower() == ".xlsx" and export_path.exists():
            try:
                export_path.unlink()
            except OSError as e:
                _lg(f"Голландия: не удалось удалить {export_path.name} — {e}")
        return xlsm_path.resolve()
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
