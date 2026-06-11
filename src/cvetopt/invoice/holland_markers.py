from __future__ import annotations

import sys
from collections.abc import Callable
from pathlib import Path

from cvetopt.invoice.ecuador_create import (
    _CHECKBOX_BMPS,
    _copy_checkbox_assets,
    _delete_row_command_buttons,
    _ensure_picture_helper_vba,
    _ole_rgb,
)
from cvetopt.invoice.xlsx_read import grid_by_row, read_xlsx_grid

LogFn = Callable[[str], None]

_HOLLAND_DATA_FIRST_ROW = 2
_MSO_AUTOMATION_SECURITY_LOW = 1
_ZEBRA_EVEN = 12379351
_ZEBRA_ODD = 9944773

_CV_SYNC_MACRO = "cv_SyncHollandMarkers"
_CV_CLICK_MACRO = "cvHollandMarkerClick"
_MARKER_MODULE = "cvHollandMarkers"
_OBSOLETE_MARKER_CLASS = "cvHollandMarkerHandler"
_XL_BUTTON_CONTROL = 0
_MSO_FORM_CONTROL = 12

_CV_SYNC_VBA = """
Public Sub cv_SyncHollandMarkers(aFirst As Long, aLast As Long)
    Dim aI As Long
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
            Call cvAddHollandMarkerBtn(aSheet, aI, 1, "1", RGB(255, 0, 0))
            Call cvAddHollandMarkerBtn(aSheet, aI, 2, "2", RGB(0, 128, 0))
            If (aI Mod 2) = 0 Then
                .Range(.Cells(aI, 3), .Cells(aI, aLastCol)).Interior.Color = 12379351
            Else
                .Range(.Cells(aI, 3), .Cells(aI, aLastCol)).Interior.Color = 9944773
            End If
        Next aI
    End With
End Sub

Private Sub cvAddHollandMarkerBtn(aSheet As Worksheet, aRow As Long, aCol As Long, aPrefix As String, aColor As Long)
    Dim btn As Shape
    Dim cell As Range
    Set cell = aSheet.Cells(aRow, aCol)
    Set btn = aSheet.Shapes.AddFormControl(0, cell.Left, cell.Top, cell.Width, cell.Height)
    btn.Name = "cvM" & aPrefix & aRow
    btn.OnAction = "cvHollandMarkerClick"
    btn.TextFrame.Characters.Text = aPrefix & " " & aRow & " 0"
    btn.Fill.Visible = True
    btn.Fill.ForeColor.RGB = aColor
    btn.Line.Visible = False
End Sub

Public Sub cvDelHollandMarkerButtons(aSheet As String)
    Dim ws As Worksheet
    Dim shp As Shape
    Dim aButton As OLEObject
    Dim i As Long
    Set ws = ThisWorkbook.Sheets(aSheet)
    For i = ws.Shapes.Count To 1 Step -1
        Set shp = ws.Shapes(i)
        If shp.Type = 12 And shp.FormControlType = 0 Then
            If Left(shp.Name, 3) = "cvM" Then shp.Delete
        End If
    Next i
    For Each aButton In ws.OLEObjects
        If InStr(1, aButton.ClassType, "CommandButton", vbTextCompare) > 0 Then
            aButton.Delete
        End If
    Next
End Sub

Public Sub cvHollandMarkerClick()
    Dim btn As Shape
    Dim cap As String
    Dim aStr As String
    Dim aAddress As String
    Dim aSheet As Worksheet
    Set aSheet = ThisWorkbook.Worksheets(1)
    Set btn = aSheet.Shapes(CStr(Application.Caller))
    cap = btn.TextFrame.Characters.Text
    If Left(cap, 1) = "1" Then
        If Right(cap, 1) = "0" Then
            btn.Fill.ForeColor.RGB = RGB(200, 0, 0)
            aStr = cap
            aStr = Right(aStr, Len(aStr) - 2)
            aStr = Left(aStr, Len(aStr) - 2)
            aAddress = aSheet.Cells(CInt(aStr), aSheet.Range("C1").End(xlToRight).Column). _
                Address(RowAbsolute:=False, ColumnAbsolute:=False)
            aSheet.Range("C" & aStr & ":" & aAddress).Interior.Color = RGB(255, 209, 209)
            btn.TextFrame.Characters.Text = Left(cap, Len(cap) - 1) & "1"
        Else
            btn.Fill.ForeColor.RGB = RGB(255, 0, 0)
            aStr = cap
            aStr = Right(aStr, Len(aStr) - 2)
            aStr = Left(aStr, Len(aStr) - 2)
            aAddress = aSheet.Cells(CInt(aStr), aSheet.Range("C1").End(xlToRight).Column). _
                Address(RowAbsolute:=False, ColumnAbsolute:=False)
            aSheet.Range("C" & aStr & ":" & aAddress).Interior.Color = RGB(255, 255, 255)
            btn.TextFrame.Characters.Text = Left(cap, Len(cap) - 1) & "0"
        End If
    ElseIf Left(cap, 1) = "2" Then
        If Right(cap, 1) = "0" Then
            btn.Fill.ForeColor.RGB = RGB(0, 200, 0)
            aStr = cap
            aStr = Right(aStr, Len(aStr) - 2)
            aStr = Left(aStr, Len(aStr) - 2)
            aAddress = aSheet.Cells(CInt(aStr), aSheet.Range("C1").End(xlToRight).Column). _
                Address(RowAbsolute:=False, ColumnAbsolute:=False)
            aSheet.Range("C" & aStr & ":" & aAddress).Interior.Color = RGB(0, 255, 0)
            btn.TextFrame.Characters.Text = Left(cap, Len(cap) - 1) & "1"
        Else
            btn.Fill.ForeColor.RGB = RGB(0, 128, 0)
            aStr = cap
            aStr = Right(aStr, Len(aStr) - 2)
            aStr = Left(aStr, Len(aStr) - 2)
            aAddress = aSheet.Cells(CInt(aStr), aSheet.Range("C1").End(xlToRight).Column). _
                Address(RowAbsolute:=False, ColumnAbsolute:=False)
            aSheet.Range("C" & aStr & ":" & aAddress).Interior.Color = RGB(255, 255, 255)
            btn.TextFrame.Characters.Text = Left(cap, Len(cap) - 1) & "0"
        End If
    End If
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


def _remove_obsolete_marker_class(vbproject: object) -> None:
    try:
        vbproject.VBComponents.Remove(vbproject.VBComponents(_OBSOLETE_MARKER_CLASS))
    except Exception:
        pass


def _ensure_std_module(vbproject: object, name: str, code: str) -> None:
    try:
        mod = vbproject.VBComponents(name)
    except Exception:
        mod = vbproject.VBComponents.Add(1)
        mod.Name = name
    code_module = mod.CodeModule
    existing = code_module.Lines(1, code_module.CountOfLines) if code_module.CountOfLines else ""
    if (
        _CV_SYNC_MACRO in existing
        and _CV_CLICK_MACRO in existing
        and "MSForms" not in existing
    ):
        return
    if code_module.CountOfLines:
        code_module.DeleteLines(1, code_module.CountOfLines)
    code_module.AddFromString(code)


def _missing_marker_assets(assets_dir: Path) -> list[str]:
    return [name for name in _CHECKBOX_BMPS if not (assets_dir / name).is_file()]


def _holland_last_col(sheet: object) -> int:
    return int(sheet.api.Range("C1").End(-4161).Column)  # xlToRight


def _delete_holland_marker_buttons(sheet: object) -> None:
    shapes = sheet.api.Shapes
    for i in range(int(shapes.Count), 0, -1):
        shp = shapes.Item(i)
        try:
            if int(shp.Type) == _MSO_FORM_CONTROL and int(shp.FormControlType) == _XL_BUTTON_CONTROL:
                if str(shp.Name).startswith("cvM"):
                    shp.Delete()
        except Exception:
            pass
    _delete_row_command_buttons(sheet)


def _add_form_marker_button(
    sheet: object,
    *,
    row: int,
    col: int,
    prefix: str,
    color_rgb: int,
    macro: str,
) -> None:
    cell = sheet.api.Cells(row, col)
    shp = sheet.api.Shapes.AddFormControl(
        _XL_BUTTON_CONTROL,
        float(cell.Left),
        float(cell.Top),
        float(cell.Width),
        float(cell.Height),
    )
    shp.Name = f"cvM{prefix}{row}"
    shp.OnAction = macro
    shp.TextFrame.Characters().Text = f"{prefix} {row} 0"
    shp.Fill.Visible = True
    shp.Fill.ForeColor.RGB = color_rgb
    shp.Line.Visible = False


def _count_marker_buttons(sheet: object) -> int:
    count = 0
    shapes = sheet.api.Shapes
    for i in range(1, int(shapes.Count) + 1):
        try:
            shp = shapes.Item(i)
            if int(shp.Type) == _MSO_FORM_CONTROL and int(shp.FormControlType) == _XL_BUTTON_CONTROL:
                if str(shp.Name).startswith("cvM"):
                    count += 1
        except Exception:
            pass
    if count:
        return count
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
    del assets_dir
    click_macro = f"{_MARKER_MODULE}.{_CV_CLICK_MACRO}"
    sheet.api.Columns("A:A").ColumnWidth = 2.2
    sheet.api.Columns("B:B").ColumnWidth = 2.2
    _delete_holland_marker_buttons(sheet)
    last_col = _holland_last_col(sheet)

    for row in range(first_row, last_row + 1):
        _add_form_marker_button(
            sheet,
            row=row,
            col=1,
            prefix="1",
            color_rgb=_ole_rgb(255, 0, 0),
            macro=click_macro,
        )
        _add_form_marker_button(
            sheet,
            row=row,
            col=2,
            prefix="2",
            color_rgb=_ole_rgb(0, 128, 0),
            macro=click_macro,
        )

        color = _ZEBRA_EVEN if row % 2 == 0 else _ZEBRA_ODD
        sheet.api.Range(
            sheet.api.Cells(row, 3),
            sheet.api.Cells(row, last_col),
        ).Interior.Color = color


def _inject_marker_vba(wb: object) -> list[str]:
    vb = wb.api.VBProject
    warnings: list[str] = []
    steps: list[tuple[str, Callable[[], None]]] = [
        ("picture helper", lambda: _ensure_picture_helper_vba(wb)),
        ("remove old class", lambda: _remove_obsolete_marker_class(vb)),
        ("markers module", lambda: _ensure_std_module(vb, _MARKER_MODULE, _CV_SYNC_VBA)),
    ]
    for label, action in steps:
        try:
            action()
        except Exception as e:
            warnings.append(f"{label}: {e}")
    return warnings


def _sync_holland_markers_vba(
    app: object,
    wb: object,
    *,
    first_row: int,
    last_row: int,
    log: LogFn,
) -> bool:
    last_err: Exception | None = None
    for warn in _inject_marker_vba(wb):
        log(f"Голландия: VBA inject — {warn}")
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

        try:
            _ensure_picture_helper_vba(wb)
        except Exception as e:
            _lg(f"Голландия: VBA helper — {e}")

        vba_ok = False
        try:
            vba_ok = _sync_holland_markers_vba(
                app,
                wb,
                first_row=_HOLLAND_DATA_FIRST_ROW,
                last_row=last_row,
                log=_lg,
            )
        except Exception as e:
            _lg(f"Голландия: VBA недоступен ({e})")

        expected = (last_row - _HOLLAND_DATA_FIRST_ROW + 1) * 2
        if not vba_ok or _count_marker_buttons(ws) < expected:
            _lg("Голландия: маркеры через COM…")
            for warn in _inject_marker_vba(wb):
                _lg(f"Голландия: VBA inject — {warn}")
            _sync_holland_markers_com(
                wb,
                ws,
                first_row=_HOLLAND_DATA_FIRST_ROW,
                last_row=last_row,
                assets_dir=assets_target,
            )

        btn_count = _count_marker_buttons(ws)
        if btn_count < expected:
            raise RuntimeError(
                f"Создано кнопок {btn_count} из {expected}. "
                "Проверьте «Доверять доступ к VBA» в Excel."
            )

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
