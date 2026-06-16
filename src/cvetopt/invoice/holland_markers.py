from __future__ import annotations

import sys
from collections.abc import Callable
from pathlib import Path

from cvetopt.invoice.ecuador_create import (
    _CHECKBOX_BMPS,
    _apply_command_button_picture,
    _copy_checkbox_assets,
    _ensure_picture_helper_vba,
    ensure_vba_references,
)
from cvetopt.invoice.xlsx_read import grid_by_row, read_xlsx_grid

LogFn = Callable[[str], None]

_HOLLAND_DATA_FIRST_ROW = 2  # в выгрузке Голландия_1_*: строка 1 — заголовки
_MSO_AUTOMATION_SECURITY_LOW = 1
_ZEBRA_EVEN = 12379351
_ZEBRA_ODD = 9944773

_CV_SYNC_MACRO = "cv_SyncHollandMarkers"
_CV_WIRE_MACRO = "cv_WireHollandMarkerButtons"
_MARKER_MODULE = "Module1"  # как cv_SyncRowCheckboxes в Эквадоре
_LEGACY_MARKER_MODULE = "cvHollandMarkers"
_MARKER_CLASS = "cvHollandButtonHandler"
_EDIT_BUTTON_NAME = "cbHollandEdit"
_XL_CHECKBOX = 1
_XL_EXCEL_LINKS = 1
_XL_UP = -4162
_HOLLAND_RESERVED_OLE = frozenset({_EDIT_BUTTON_NAME})

_CV_SYNC_VBA = """
Public ColHollandButtons As Collection

Public Sub Auto_Open()
    Call cv_WireHollandMarkerButtons
End Sub

Public Sub cv_WireHollandMarkerButtons()
    Dim aButton As OLEObject
    Dim h As cvHollandButtonHandler
    Application.EnableEvents = True
    Set ColHollandButtons = New Collection
    For Each aButton In ThisWorkbook.Worksheets(1).OLEObjects
        If aButton.Name = "cbHollandEdit" Then GoTo NextBtn
        If TypeOf aButton.Object Is MSForms.CommandButton Then
            Set h = New cvHollandButtonHandler
            Set h.EventButton = aButton.Object
            ColHollandButtons.Add h
        End If
NextBtn:
    Next
End Sub

Public Sub cv_SyncHollandMarkers(aFirst As Long, aLast As Long)
    Dim aI As Long
    Dim aCommandButton As MSForms.CommandButton
    Dim aSheet As Worksheet
    Dim aLastCol As Long
    Dim aPath As String
    Set aSheet = ThisWorkbook.Worksheets(1)
    aPath = ThisWorkbook.Path & "\\"
    Application.ScreenUpdating = False
    Call cvDelExportCheckboxes(aSheet.Name)
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
                .Picture = LoadPicture(aPath & "Red_Check_Off.bmp")
                .Caption = "1 " & Trim(Str(aI)) & " 0"
            End With
            Set aCommandButton = .OLEObjects.Add(ClassType:="Forms.CommandButton.1").Object
            With aCommandButton
                .Left = aSheet.Cells(aI, 2).Left
                .Top = aSheet.Cells(aI, 2).Top
                .Width = aSheet.Cells(aI, 2).Width
                .Height = aSheet.Cells(aI, 2).Height
                .Picture = LoadPicture(aPath & "Green_Check_Off.bmp")
                .Caption = "2 " & Trim(Str(aI)) & " 0"
            End With
            If (aI Mod 2) = 0 Then
                .Range(.Cells(aI, 3), .Cells(aI, aLastCol)).Interior.Color = 12379351
            Else
                .Range(.Cells(aI, 3), .Cells(aI, aLastCol)).Interior.Color = 9944773
            End If
        Next aI
    End With
End Sub

Public Sub cvDelExportCheckboxes(aSheet As String)
    Dim ws As Worksheet
    Dim shp As Shape
    Dim i As Long
    Set ws = ThisWorkbook.Sheets(aSheet)
    For i = ws.Shapes.Count To 1 Step -1
        Set shp = ws.Shapes(i)
        On Error Resume Next
        If shp.FormControlType = 1 Then shp.Delete
        On Error GoTo 0
    Next i
End Sub

Public Sub cvDelHollandMarkerButtons(aSheet As String)
    Dim ws As Worksheet
    Dim shp As Shape
    Dim aButton As OLEObject
    Dim i As Long
    Set ws = ThisWorkbook.Sheets(aSheet)
    For i = ws.Shapes.Count To 1 Step -1
        Set shp = ws.Shapes(i)
        If Left(shp.Name, 4) = "cvM_" Then shp.Delete
    Next i
    For Each aButton In ws.OLEObjects
        If aButton.Name = "cbHollandEdit" Then GoTo NextOle
        If InStr(1, aButton.ClassType, "CommandButton", vbTextCompare) > 0 Then
            aButton.Delete
        End If
NextOle:
    Next
End Sub
"""

_MARKER_CLASS_VBA = """
Public WithEvents EventButton As MSForms.CommandButton

Private Function cvHollandBmp(name As String) As String
    cvHollandBmp = ThisWorkbook.Path & "\\" & name
End Function

Private Sub EventButton_Click()
    Dim aStr As String
    Dim aAddress As String
    Dim aSheet As Worksheet
    Set aSheet = ThisWorkbook.Worksheets(1)
    With EventButton
        If .Name = "cbHollandEdit" Then Exit Sub
        If Left(.Caption, 1) = "1" Then
            If Right(.Caption, 1) = "0" Then
                .Picture = LoadPicture(cvHollandBmp("Red_Check_On.bmp"))
                aStr = .Caption
                aStr = Right(aStr, Len(aStr) - 2)
                aStr = Left(aStr, Len(aStr) - 2)
                aAddress = aSheet.Cells(CInt(aStr), aSheet.Range("C1").End(xlToRight).Column). _
                    Address(RowAbsolute:=False, ColumnAbsolute:=False)
                aSheet.Range("C" & aStr & ":" & aAddress).Interior.Color = RGB(255, 209, 209)
                .Caption = Left(.Caption, Len(.Caption) - 2) & " 1"
            Else
                .Picture = LoadPicture(cvHollandBmp("Red_Check_Off.bmp"))
                aStr = .Caption
                aStr = Right(aStr, Len(aStr) - 2)
                aStr = Left(aStr, Len(aStr) - 2)
                aAddress = aSheet.Cells(CInt(aStr), aSheet.Range("C1").End(xlToRight).Column). _
                    Address(RowAbsolute:=False, ColumnAbsolute:=False)
                aSheet.Range("C" & aStr & ":" & aAddress).Interior.Color = RGB(255, 255, 255)
                .Caption = Left(.Caption, Len(.Caption) - 2) & " 0"
            End If
        ElseIf Left(.Caption, 1) = "2" Then
            If Right(.Caption, 1) = "0" Then
                .Picture = LoadPicture(cvHollandBmp("Green_Check_On.bmp"))
                aStr = .Caption
                aStr = Right(aStr, Len(aStr) - 2)
                aStr = Left(aStr, Len(aStr) - 2)
                aAddress = aSheet.Cells(CInt(aStr), aSheet.Range("C1").End(xlToRight).Column). _
                    Address(RowAbsolute:=False, ColumnAbsolute:=False)
                aSheet.Range("C" & aStr & ":" & aAddress).Interior.Color = RGB(0, 255, 0)
                .Caption = Left(.Caption, Len(.Caption) - 2) & " 1"
            Else
                .Picture = LoadPicture(cvHollandBmp("Green_Check_Off.bmp"))
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

_WORKBOOK_OPEN_VBA = """
Private Sub Workbook_Open()
    Application.EnableEvents = True
    Call cv_WireHollandMarkerButtons
End Sub
"""

_WORKSHEET_ACTIVATE_VBA = """
Private Sub Worksheet_Activate()
    Application.EnableEvents = True
    Call cv_WireHollandMarkerButtons
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


def _remove_obsolete_marker_classes(vbproject: object) -> None:
    for name in (_MARKER_CLASS, "cvHollandMarkerHandler"):
        try:
            vbproject.VBComponents.Remove(vbproject.VBComponents(name))
        except Exception:
            pass


def _remove_legacy_holland_module(vbproject: object) -> None:
    try:
        vbproject.VBComponents.Remove(vbproject.VBComponents(_LEGACY_MARKER_MODULE))
    except Exception:
        pass


def _vb_has_macro(vbproject: object, module_name: str, macro_name: str) -> bool:
    try:
        mod = vbproject.VBComponents(module_name)
        cm = mod.CodeModule
        line_count = int(cm.CountOfLines)
        if not line_count:
            return False
        existing = cm.Lines(1, line_count)
        return f"Sub {macro_name}" in existing
    except Exception:
        return False


def _ensure_holland_marker_macros(wb: object) -> None:
    """Public-макросы в Module1 — тот же приём, что cv_SyncRowCheckboxes в Эквадоре."""
    vb = wb.api.VBProject
    ensure_vba_references(vb)
    try:
        mod = vb.VBComponents(_MARKER_MODULE)
    except Exception:
        mod = vb.VBComponents.Add(1)
        mod.Name = _MARKER_MODULE
    code_module = mod.CodeModule
    line_count = int(code_module.CountOfLines)
    existing = code_module.Lines(1, line_count) if line_count else ""
    if (
        f"Sub {_CV_SYNC_MACRO}" in existing
        and f"Sub {_CV_WIRE_MACRO}" in existing
        and "Sub Auto_Open" in existing
        and "cvDelExportCheckboxes" in existing
    ):
        return
    if code_module.CountOfLines:
        code_module.DeleteLines(1, code_module.CountOfLines)
    code_module.AddFromString(_CV_SYNC_VBA)


def _ensure_workbook_open_hook(wb: object) -> None:
    """Workbook_Open в ThisWorkbook → клики маркеров подключаются сами при открытии."""
    vb = wb.api.VBProject
    doc = vb.VBComponents("ThisWorkbook").CodeModule
    line_count = int(doc.CountOfLines)
    existing = doc.Lines(1, line_count) if line_count else ""
    if "Workbook_Open" in existing and _CV_WIRE_MACRO in existing:
        return
    if "Workbook_Open" in existing:
        # Чужой Workbook_Open — не трогаем; сработают Auto_Open и Worksheet_Activate.
        return
    doc.AddFromString(_WORKBOOK_OPEN_VBA)


def _ensure_sheet_activate_hook(wb: object) -> None:
    """Worksheet_Activate — запасной путь, если Workbook_Open не сработал."""
    vb = wb.api.VBProject
    codename = str(wb.sheets[0].api.CodeName)
    doc = vb.VBComponents(codename).CodeModule
    line_count = int(doc.CountOfLines)
    existing = doc.Lines(1, line_count) if line_count else ""
    if "Worksheet_Activate" in existing and _CV_WIRE_MACRO in existing:
        return
    doc.AddFromString(_WORKSHEET_ACTIVATE_VBA)


def _prepare_workbook_for_macro_run(wb: object) -> None:
    try:
        wb.activate()
    except Exception:
        pass
    try:
        wb.save()
    except Exception:
        pass


def _ensure_class_module(vbproject: object, name: str, code: str) -> None:
    try:
        mod = vbproject.VBComponents(name)
    except Exception:
        mod = vbproject.VBComponents.Add(2)
        mod.Name = name
    code_module = mod.CodeModule
    existing = code_module.Lines(1, code_module.CountOfLines) if code_module.CountOfLines else ""
    if (
        "EventButton_Click" in existing
        and "Red_Check_On.bmp" in existing
        and "cvHollandBmp" in existing
        and "_holland_bmp" not in existing
    ):
        return
    if code_module.CountOfLines:
        code_module.DeleteLines(1, code_module.CountOfLines)
    code_module.AddFromString(code)


def _calculate_workbook(app: object | None) -> None:
    if app is None:
        return
    api = app.api
    for fn in ("CalculateFullRebuild", "CalculateFull", "Calculate"):
        try:
            getattr(api, fn)()
            return
        except Exception:
            continue


def _break_external_links(wb_api: object) -> None:
    try:
        sources = wb_api.LinkSources(_XL_EXCEL_LINKS)
    except Exception:
        return
    if not sources:
        return
    if isinstance(sources, str):
        sources = (sources,)
    for name in sources:
        try:
            wb_api.BreakLink(Name=name, Type=_XL_EXCEL_LINKS)
        except Exception:
            pass


def _freeze_sheet_values(ws: object, *, app: object | None = None, recalc: bool = True) -> None:
    """Формулы btnExport2 (ссылки на Auto_new.xls) → значения до вставки колонок A–B.

    recalc=False: НЕ пересчитывать и не обновлять ссылки — берём кэшированные
    значения (Auto_new уже закрыт, пересчёт превратил бы Quant в #ССЫЛКА!/код ошибки).
    """
    wb_api = ws.api.Parent
    if recalc:
        try:
            wb_api.UpdateLink(Name=None, Type=_XL_EXCEL_LINKS)
        except Exception:
            pass
        _calculate_workbook(app)
    # Сначала фиксируем вычисленные значения (пока Auto_new открыт и ссылки живы),
    # и только потом рвём ссылки — иначе BreakLink превратит Quant в #ССЫЛКА!.
    used = ws.api.UsedRange
    if used is not None:
        used.Value = used.Value2
        ws.api.Application.CutCopyMode = False
    _break_external_links(wb_api)


def _delete_export_checkboxes(sheet: object) -> int:
    """Чекбоксы из btnExport2 (колонка L и др.) — не нужны с маркерами A–B."""
    removed = 0
    shapes = sheet.api.Shapes
    for i in range(int(shapes.Count), 0, -1):
        shp = shapes.Item(i)
        try:
            if int(shp.FormControlType) == _XL_CHECKBOX:
                shp.Delete()
                removed += 1
        except Exception:
            pass
    return removed


def _header_columns(ws: object) -> dict[str, int]:
    api = ws.api
    cols: dict[str, int] = {}
    for col in range(1, 30):
        try:
            header = str(api.Cells(1, col).Value2 or "").strip().casefold()
        except Exception:
            continue
        if header and header not in cols:
            cols[header] = col
    return cols


def _ws_last_data_row(ws: object, key_col: int = 2) -> int:
    api = ws.api
    try:
        return int(api.Cells(api.Rows.Count, key_col).End(_XL_UP).Row)
    except Exception:
        return 1


def _is_excel_error_value(value: object) -> bool:
    try:
        return isinstance(value, (int, float)) and -2146826300 < float(value) < -2146826200
    except Exception:
        return False


_BROKEN_REF_TOKENS = ("#REF!", "#ССЫЛКА!")
_LOOKUP_KEY_ORDER = ("s2", "description", "packing", "kolli", "box nr.")


def _formula_has_broken_ref(formula: str) -> bool:
    upper = formula.upper()
    return any(tok in upper for tok in _BROKEN_REF_TOKENS)


def _replace_broken_ref(formula: str, cell_ref: str) -> str:
    for tok in _BROKEN_REF_TOKENS:
        if tok in formula:
            return formula.replace(tok, cell_ref, 1)
    folded = formula.casefold()
    for tok in _BROKEN_REF_TOKENS:
        idx = folded.find(tok.casefold())
        if idx >= 0:
            return formula[:idx] + cell_ref + formula[idx + len(tok) :]
    return formula


def _repair_export_ref_formulas(ws: object, app: object, log: LogFn) -> None:
    """btnExport2 копирует формулы со сломанным ключом (#REF!/ #ССЫЛКА!).

    Чиним все такие колонки (Quant, S1, …), подставляя ключ из выгрузки.
  Порядок: S2 → Description → Packing (пока Auto_new открыт).
    """
    api = ws.api
    headers = _header_columns(ws)
    if not headers:
        return
    key_candidates = [name for name in _LOOKUP_KEY_ORDER if name in headers]
    if not key_candidates:
        log("Голландия: нет колонок-ключей для восстановления формул.")
        return
    anchor_col = headers.get("packing") or headers.get("s2") or next(iter(headers.values()))
    last_row = _ws_last_data_row(ws, key_col=anchor_col)
    if last_row < 2:
        return

    repaired_any = False
    for header, col in sorted(headers.items(), key=lambda x: x[1]):
        try:
            probe = str(api.Cells(2, col).Formula)
        except Exception:
            continue
        if not probe.startswith("=") or not _formula_has_broken_ref(probe):
            continue

        col_fixed = False
        for key_name in key_candidates:
            if headers.get(key_name) == col:
                continue
            key_letter = _col_letter(headers[key_name])
            for row in range(2, last_row + 1):
                try:
                    cell = api.Cells(row, col)
                    formula = str(cell.Formula)
                    if not formula.startswith("="):
                        continue
                    if not _formula_has_broken_ref(formula):
                        formula = probe
                    cell.Formula = _replace_broken_ref(formula, f"{key_letter}{row}")
                except Exception:
                    pass
            _calculate_workbook(app)
            try:
                sample = api.Cells(2, col).Value2
            except Exception:
                sample = None
            if not _is_excel_error_value(sample):
                log(
                    f"Голландия: «{header}» восстановлена по ключу «{key_name}» "
                    f"({key_letter}), пример={sample!r}."
                )
                col_fixed = True
                repaired_any = True
                break
            log(f"Голландия: «{header}» — ключ «{key_name}» не подошёл, пробую следующий…")
        if not col_fixed:
            log(f"Голландия: «{header}» — не удалось восстановить (#ССЫЛКА!/ #ЗНАЧ!).")

    if not repaired_any:
        log("Голландия: колонок с #ССЫЛКА! в формулах не найдено.")


def _log_frozen_column_errors(ws: object, log: LogFn) -> None:
    headers = _header_columns(ws)
    anchor = headers.get("packing") or headers.get("s2") or 2
    last_row = _ws_last_data_row(ws, key_col=anchor)
    api = ws.api
    for name in ("quant", "s1", "s2", "description", "kolli"):
        col = headers.get(name)
        if not col:
            continue
        err_rows = 0
        for row in range(2, last_row + 1):
            try:
                if _is_excel_error_value(api.Cells(row, col).Value2):
                    err_rows += 1
            except Exception:
                pass
        if err_rows:
            log(f"Голландия: в «{name}» после заморозки кодов ошибки: {err_rows} строк.")


def _col_letter(col: int) -> str:
    letters = ""
    while col > 0:
        col, rem = divmod(col - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def _log_export_formulas(ws: object, log: LogFn) -> None:
    """Диагностика: что btnExport2 положил в строку заголовка и первую строку данных."""
    api = ws.api
    for row in (1, 2):
        parts: list[str] = []
        for col in range(1, 16):
            try:
                cell = api.Cells(row, col)
                formula = str(cell.Formula)
                value = cell.Value2
            except Exception:
                continue
            if formula in ("", "None"):
                continue
            parts.append(f"{_col_letter(col)}{row}={formula!r}(={value!r})")
        if parts:
            log("Голландия диагностика: " + "; ".join(parts))


def fix_holland_export_after_auto1(app: object, export_dir: Path, log: LogFn) -> None:
    """Сразу после btnExport2: пересчёт пока Auto_new открыт → значения, без чекбоксов."""
    candidates = [
        p for p in export_dir.glob("Голландия_1_*.xlsx") if p.is_file()
    ]
    if not candidates:
        log("Голландия: файл экспорта не найден — постобработка пропущена")
        return
    export_path = max(candidates, key=lambda p: p.stat().st_mtime)

    wb_holland: object | None = None
    for book in app.books:
        if str(book.name).casefold() == export_path.name.casefold():
            wb_holland = book
            break
    opened_here = False
    if wb_holland is None:
        wb_holland = app.books.open(str(export_path), update_links=3)
        opened_here = True
    ws = wb_holland.sheets[0]
    try:
        _log_export_formulas(ws, log)
    except Exception:
        pass
    try:
        _repair_export_ref_formulas(ws, app, log)
    except Exception as e:
        log(f"Голландия: восстановление формул пропущено — {e}")
    _freeze_sheet_values(ws, app=app)
    try:
        _log_frozen_column_errors(ws, log)
    except Exception:
        pass
    removed = _delete_export_checkboxes(ws)
    wb_holland.save()
    if opened_here:
        wb_holland.close()
    log(
        f"Голландия: {export_path.name} — формулы → значения, "
        f"удалено чекбоксов: {removed}"
    )

def _missing_marker_assets(assets_dir: Path) -> list[str]:
    return [name for name in _CHECKBOX_BMPS if not (assets_dir / name).is_file()]


def _holland_last_col(sheet: object) -> int:
    return int(sheet.api.Range("C1").End(-4161).Column)  # xlToRight


def _delete_holland_marker_buttons(sheet: object) -> None:
    shapes = sheet.api.Shapes
    for i in range(int(shapes.Count), 0, -1):
        shp = shapes.Item(i)
        try:
            if str(shp.Name).startswith("cvM_"):
                shp.Delete()
        except Exception:
            pass
    oles = sheet.api.OLEObjects()
    for i in range(int(oles.Count), 0, -1):
        ole = oles.Item(i)
        try:
            if str(ole.Name) in _HOLLAND_RESERVED_OLE:
                continue
            if "CommandButton" in str(ole.ClassType):
                ole.Delete()
        except Exception:
            pass


def _is_command_button_ole(ole: object) -> bool:
    """У OLEObject читается progID (Forms.CommandButton.1), а не ClassType."""
    seen_any = False
    for attr in ("progID", "ClassType", "OLEClass"):
        try:
            val = str(getattr(ole, attr))
        except Exception:
            continue
        seen_any = True
        if "CommandButton" in val:
            return True
    # Свойства прочитались, но это не CommandButton.
    if seen_any:
        return False
    # Тип определить не удалось — на листе маркеров остаются только кнопки.
    return True


def _count_marker_buttons(sheet: object) -> int:
    count = 0
    oles = sheet.api.OLEObjects()
    for i in range(1, int(oles.Count) + 1):
        try:
            ole = oles.Item(i)
        except Exception:
            continue
        try:
            if str(ole.Name) in _HOLLAND_RESERVED_OLE:
                continue
        except Exception:
            pass
        if _is_command_button_ole(ole):
            count += 1
    return count


def _remove_holland_edit_button(sheet: object) -> None:
    """Убрать устаревшую кнопку «Редактировать» (клики подключаются при открытии)."""
    api = sheet.api
    for i in range(int(api.OLEObjects().Count), 0, -1):
        try:
            if str(api.OLEObjects().Item(i).Name) == _EDIT_BUTTON_NAME:
                api.OLEObjects().Item(i).Delete()
        except Exception:
            pass
    shapes = api.Shapes
    for i in range(int(shapes.Count), 0, -1):
        try:
            if str(shapes.Item(i).Name) == _EDIT_BUTTON_NAME:
                shapes.Item(i).Delete()
        except Exception:
            pass


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
    _delete_export_checkboxes(sheet)
    _delete_holland_marker_buttons(sheet)
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
            _apply_command_button_picture(btn, img, app_api, wb=wb)
            btn.Caption = f"{prefix} {row} 0"

        color = _ZEBRA_EVEN if row % 2 == 0 else _ZEBRA_ODD
        sheet.api.Range(
            sheet.api.Cells(row, 3),
            sheet.api.Cells(row, last_col),
        ).Interior.Color = color


def _inject_marker_vba(wb: object) -> list[str]:
    warnings: list[str] = []
    try:
        vb = wb.api.VBProject
    except Exception as e:
        return [f"VBProject: {e} (включите «Доверять доступ к объектной модели VBA»)"]
    steps: list[tuple[str, Callable[[], None]]] = [
        ("picture helper", lambda: _ensure_picture_helper_vba(wb)),
        ("legacy module", lambda: _remove_legacy_holland_module(vb)),
        ("remove old class", lambda: _remove_obsolete_marker_classes(vb)),
        ("markers module", lambda: _ensure_holland_marker_macros(wb)),
        ("click class", lambda: _ensure_class_module(vb, _MARKER_CLASS, _MARKER_CLASS_VBA)),
        ("open hook", lambda: _ensure_workbook_open_hook(wb)),
        ("activate hook", lambda: _ensure_sheet_activate_hook(wb)),
    ]
    for label, action in steps:
        try:
            action()
        except Exception as e:
            warnings.append(f"{label}: {e}")
    if not _vb_has_macro(vb, _MARKER_MODULE, _CV_SYNC_MACRO):
        warnings.append(f"макрос {_CV_SYNC_MACRO} не найден в {_MARKER_MODULE}")
    return warnings


def _wire_holland_marker_clicks(app: object, wb: object, *, log: LogFn) -> None:
    _prepare_workbook_for_macro_run(wb)
    wb_name = str(wb.name)
    specs = (
        _CV_WIRE_MACRO,
        f"{_MARKER_MODULE}.{_CV_WIRE_MACRO}",
        f"'{wb_name}'!{_CV_WIRE_MACRO}",
    )
    for spec in specs:
        try:
            app.api.Run(spec)
            log("Голландия: клики подключены (WithEvents).")
            return
        except Exception:
            continue
    log("Голландия: клики не подключились — откройте файл с включёнными макросами.")


def _sync_holland_markers_vba(
    app: object,
    wb: object,
    *,
    first_row: int,
    last_row: int,
    log: LogFn,
) -> bool:
    last_err: Exception | None = None
    inject_warnings = _inject_marker_vba(wb)
    for warn in inject_warnings:
        log(f"Голландия: VBA inject — {warn}")
    if not _vb_has_macro(wb.api.VBProject, _MARKER_MODULE, _CV_SYNC_MACRO):
        log("Голландия: VBA inject — макрос не в Module1, пропускаем Run")
        return False

    _prepare_workbook_for_macro_run(wb)
    wb_name = str(wb.name)
    specs = (
        _CV_SYNC_MACRO,
        f"{_MARKER_MODULE}.{_CV_SYNC_MACRO}",
        f"'{wb_name}'!{_CV_SYNC_MACRO}",
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
            "(скопируйте Red/Green_Check_*.bmp рядом с файлом)."
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
        # Ручной расчёт + без обновления ссылок: Auto_new уже закрыт,
        # пересчёт превратил бы Quant в код ошибки (-2146826265).
        try:
            app.api.Calculation = -4135  # xlCalculationManual
        except Exception:
            pass
        wb = app.books.open(str(export_path), update_links=0)
        ws = wb.sheets[0]
        _freeze_sheet_values(ws, app=app, recalc=False)
        removed = _delete_export_checkboxes(ws)
        _lg(
            f"Голландия: формулы → значения; удалено старых чекбоксов: {removed}"
        )
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
        btn_count = _count_marker_buttons(ws)
        if not vba_ok and btn_count < expected:
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
            _lg(f"Голландия: маркеры готовы (COM), кнопок: {btn_count}.")
        elif vba_ok and btn_count < expected:
            _lg(
                f"Голландия: VBA отработал, в подсчёте {btn_count}/{expected} кн. "
                "— доверяем VBA, сохраняем как есть."
            )
        if not vba_ok and btn_count < expected:
            raise RuntimeError(
                f"Создано кнопок {btn_count} из {expected}. "
                "Проверьте «Доверять доступ к VBA» в Excel."
            )

        try:
            _remove_holland_edit_button(ws)
            _wire_holland_marker_clicks(app, wb, log=_lg)
        except Exception as e:
            _lg(f"Голландия: подключение кликов — {e}")

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
