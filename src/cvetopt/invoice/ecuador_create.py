from __future__ import annotations

import os
import re
import shutil
import sys
import time
import zipfile
from collections.abc import Callable
from datetime import datetime
from pathlib import Path

from cvetopt.core.runtime_settings import (
    _archive_one_entry,
    _archive_target_path,
    order_id_from_biflorica_report,
    resolve_ecuador_archive_dir,
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
_COL_DATA_FIRST = 4  # D
_COL_DATA_LAST = 28  # AB
# Макросы должны быть включены при открытии книги — иначе Run требует переоткрытия.
_MSO_AUTOMATION_SECURITY_LOW = 1
_WORK_COPY_NAME = "_cvetopt_ecuador_work.xlsm"
_PATH_SHEET_AUTO_OPEN_CELL = '<c r="A2" t="inlineStr"><is><t>False</t></is></c>'
_RESERVED_OLE_BUTTONS = frozenset({"cbTransaction", "cbRestart", "cbCreateTimeFile"})
_ZEBRA_EVEN = 12379351
_ZEBRA_ODD = 9944773
_CHECKBOX_BMPS = (
    "Red_Check_Off.bmp",
    "Red_Check_On.bmp",
    "Green_Check_Off.bmp",
    "Green_Check_On.bmp",
)
_CV_CREATE_MACRO = "cv_Run_CreateFile"
_CV_CREATE_VBA = """
Public Sub cv_Run_CreateFile(aSaveDir As String)
    Dim aName As String
    Dim aPath As String
    Dim saveDir As String
    saveDir = aSaveDir
    If saveDir = "" Then Exit Sub
    aName = "Эквадор " & Format(Now(), "d/m/yy hh.nn") & ".xlsm"
    aPath = saveDir & "\\" & aName
    Call cv_ClearDataNotes(7, 500)
    ThisWorkbook.Save
    ThisWorkbook.SaveAs aPath
    With Workbooks(aName).Sheets(1)
        .Name = "Форматирование"
        .cbCreateTimeFile.Visible = False
        .cbTransaction.Visible = True
    End With
    With Workbooks(aName).Sheets(2)
        .Cells(2, 1).Value = "False"
        .Range("B3").Value = aPath
    End With
    Workbooks(aName).Save
End Sub
"""
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
_CV_CLEAR_NOTES_MACRO = "cv_ClearDataNotes"
_CV_CLEAR_NOTES_VBA = """
Public Sub cv_ClearDataNotes(aFirst As Long, aLast As Long)
    Dim aSheet As Worksheet
    Dim rng As Range
    Dim c As Range
    If aLast < aFirst Then Exit Sub
    Set aSheet = ThisWorkbook.Sheets(1)
    Set rng = aSheet.Range("D" & aFirst & ":AB" & aLast)
    On Error Resume Next
    rng.ClearComments
    rng.ClearNotes
    On Error GoTo 0
    For Each c In rng.Cells
        On Error Resume Next
        c.ClearNote
        If Not c.Comment Is Nothing Then c.Comment.Delete
        If Not c.CommentThread Is Nothing Then c.CommentThread.Delete
        On Error GoTo 0
    Next c
End Sub
"""


def find_latest_biflorica_report(download_dir: Path) -> Path | None:
    """Самый свежий BiFlorica-<id>__<дата>.xlsx в папке скачивания."""
    if not download_dir.is_dir():
        return None
    candidates = [
        entry
        for entry in download_dir.iterdir()
        if entry.is_file()
        and entry.suffix.lower() == ".xlsx"
        and order_id_from_biflorica_report(entry) is not None
    ]
    if not candidates:
        return None
    return max(candidates, key=lambda path: path.stat().st_mtime)


def ecuador_output_basename(when: datetime | None = None) -> str:
    """
    Имя файла как в кнопке «Создать файл», но без «/» в дате —
    иначе Windows создаёт подпапки (Эквадор 03\\06\\…).
    """
    t = when or datetime.now()
    stamp = t.strftime("%d.%m.%y %H.%M")
    return f"Эквадор {stamp}.xlsm"


def _configure_excel_app_base(app: object) -> None:
    app.display_alerts = False
    app.screen_updating = False
    api = app.api
    api.DisplayAlerts = False


def _configure_excel_app(app: object) -> None:
    """Path!A2 патчится до открытия — макросы можно держать включёнными."""
    _configure_excel_app_base(app)
    api = app.api
    api.EnableEvents = True
    try:
        api.AutomationSecurity = _MSO_AUTOMATION_SECURITY_LOW
    except Exception:
        pass


def _prepare_work_copy(template: Path) -> Path:
    """Копия в папке шаблона (доверенное расположение + bmp рядом)."""
    work = template.parent / _WORK_COPY_NAME
    if work.exists():
        work.unlink()
    shutil.copy2(template, work)
    _patch_path_auto_open_off(work)
    return work


_PICTURE_HELPER_MODULE = "cvPictureHelper"
_PICTURE_HELPER_VBA = """
Public Function cv_LoadPicture(path As String)
    Set cv_LoadPicture = LoadPicture(path)
End Function
"""

_MSFORMS_GUID = "{0D452EE1-E08F-101A-8523-02608C4D0BB4}"


def _vb_has_reference(vbproject: object, hint: str) -> bool:
    hint_cf = hint.casefold()
    refs = vbproject.References
    for i in range(1, int(refs.Count) + 1):
        ref = refs.Item(i)
        for attr in ("Name", "Description", "FullPath"):
            try:
                if hint_cf in str(getattr(ref, attr, "")).casefold():
                    return True
            except Exception:
                pass
    return False


def _ensure_msforms_reference(vbproject: object) -> None:
    """Нужен для MSForms.CommandButton и WithEvents в инжектируемом VBA."""
    if _vb_has_reference(vbproject, "Forms"):
        return
    try:
        vbproject.References.AddFromGuid(_MSFORMS_GUID, 2, 0)
        return
    except Exception:
        pass
    system_root = Path(os.environ.get("SystemRoot", r"C:\Windows"))
    for fm20 in (system_root / "SysWOW64" / "FM20.DLL", system_root / "System32" / "FM20.DLL"):
        if not fm20.is_file():
            continue
        try:
            vbproject.References.AddFromFile(str(fm20))
            return
        except Exception:
            continue
    raise RuntimeError(
        "Не удалось подключить Microsoft Forms 2.0 (FM20.DLL). "
        "Нужно для CommandButton в VBA-маркерах."
    )


def ensure_vba_references(vbproject: object) -> None:
    _ensure_msforms_reference(vbproject)


def _ensure_picture_helper_vba(wb: object) -> None:
    vb = wb.api.VBProject
    try:
        mod = vb.VBComponents(_PICTURE_HELPER_MODULE)
    except Exception:
        mod = vb.VBComponents.Add(1)
        mod.Name = _PICTURE_HELPER_MODULE
    code_module = mod.CodeModule
    existing = code_module.Lines(1, code_module.CountOfLines) if code_module.CountOfLines else ""
    if "cv_LoadPicture" in existing:
        return
    if code_module.CountOfLines:
        code_module.InsertLines(code_module.CountOfLines + 1, _PICTURE_HELPER_VBA)
    else:
        code_module.AddFromString(_PICTURE_HELPER_VBA)


def _ole_rgb(r: int, g: int, b: int) -> int:
    return int(r) + int(g) * 256 + int(b) * 65536


def _load_picture(app_api: object, image_path: str, *, wb: object | None = None) -> object | None:
    path = str(Path(image_path).resolve())
    if wb is not None:
        try:
            _ensure_picture_helper_vba(wb)
            return wb.app.api.Run(f"{_PICTURE_HELPER_MODULE}.cv_LoadPicture", path)
        except Exception:
            pass
    try:
        escaped = path.replace('"', '""')
        return app_api.Evaluate(f'LoadPicture("{escaped}")')
    except Exception:
        pass
    try:
        return app_api.LoadPicture(path)
    except Exception:
        pass
    try:
        import win32com.client

        return win32com.client.Dispatch(app_api).LoadPicture(path)
    except Exception:
        return None


def _apply_command_button_picture(btn: object, image_path: str, app_api: object, *, wb: object | None) -> None:
    pic = _load_picture(app_api, image_path, wb=wb)
    if pic is not None:
        btn.Picture = pic
        return
    # Запасной вид без bmp (если VBA/LoadPicture недоступны).
    name = Path(image_path).name.casefold()
    if "red" in name:
        btn.BackColor = _ole_rgb(255, 0, 0)
    elif "green" in name:
        btn.BackColor = _ole_rgb(0, 128, 0)
    btn.Caption = ""


def _patch_path_auto_open_off(xlsm_path: Path) -> None:
    """
    Path!A2 в шаблоне = True → Workbook_Open сразу открывает диалог «Выбрать файл».
    Меняем на False в XML до открытия Excel.
    """
    sheet2 = "xl/worksheets/sheet2.xml"
    with zipfile.ZipFile(xlsm_path, "r") as zin:
        raw = zin.read(sheet2).decode("utf-8")
        patched = re.sub(
            r"<c r=\"A2\"(?:[^>]*)>(?:.*?</c>|/>)",
            _PATH_SHEET_AUTO_OPEN_CELL,
            raw,
            count=1,
            flags=re.DOTALL,
        )
        if patched == raw:
            raise RuntimeError("Эквадор: не удалось выставить Path!A2=False в копии шаблона")
        tmp_out = xlsm_path.with_name(xlsm_path.name + ".a2patch")
        with zipfile.ZipFile(tmp_out, "w") as zout:
            for item in zin.infolist():
                payload = patched.encode("utf-8") if item.filename == sheet2 else zin.read(item.filename)
                zout.writestr(item, payload, compress_type=item.compress_type)
    tmp_out.replace(xlsm_path)


def _open_workbook_quiet(app: object, path: Path) -> object:
    """Открытие без Notify и без срабатывания Workbook_Open."""
    api = app.api
    wb_api = api.Workbooks.Open(
        str(path),
        0,
        False,
        None,
        None,
        None,
        True,
        None,
        None,
        None,
        None,
        False,
        None,
        False,
    )
    for book in app.books:
        if str(book.fullname).lower() == str(path.resolve()).lower():
            return book
    return app.books[-1]


def _cell_has_annotation(cell_api: object) -> bool:
    try:
        if cell_api.Comment is not None:
            return True
    except Exception:
        pass
    try:
        if getattr(cell_api, "Note", None) is not None:
            return True
    except Exception:
        pass
    try:
        if cell_api.CommentThread is not None:
            return True
    except Exception:
        pass
    return False


def _strip_cell_annotation(cell_api: object) -> None:
    for method_name in ("ClearNote",):
        try:
            getattr(cell_api, method_name)()
        except Exception:
            pass
    try:
        comment = cell_api.Comment
        if comment is not None:
            comment.Delete()
    except Exception:
        pass
    try:
        note = getattr(cell_api, "Note", None)
        if note is not None:
            note.Delete()
    except Exception:
        pass
    try:
        thread = cell_api.CommentThread
        if thread is not None:
            thread.Delete()
    except Exception:
        pass


def _strip_range_annotations_python(
    sheet: object,
    *,
    first_row: int,
    last_row: int,
    first_col: int = _COL_DATA_FIRST,
    last_col: int = _COL_DATA_LAST,
) -> int:
    """Построчно ClearNote/Comment — запасной путь, если VBA недоступен."""
    api = sheet.api
    removed = 0
    for row in range(first_row, last_row + 1):
        for col in range(first_col, last_col + 1):
            cell = api.Cells(row, col)
            if _cell_has_annotation(cell):
                removed += 1
            _strip_cell_annotation(cell)
    return removed


def _ensure_cv_clear_notes_macro(wb: object) -> None:
    vb = wb.api.VBProject
    ensure_vba_references(vb)
    mod = vb.VBComponents("Module1").CodeModule
    line_count = int(mod.CountOfLines)
    existing = mod.Lines(1, line_count) if line_count else ""
    if f"Sub {_CV_CLEAR_NOTES_MACRO}" in existing:
        return
    mod.InsertLines(line_count + 1, _CV_CLEAR_NOTES_VBA)


def _run_clear_data_notes_macro(
    wb: object,
    *,
    first_row: int,
    last_row: int,
    log: LogFn | None = None,
) -> bool:
    if last_row < first_row:
        return True

    def _lg(msg: str) -> None:
        if log is not None:
            log(msg)

    try:
        _ensure_cv_clear_notes_macro(wb)
    except Exception as e:
        _lg(f"Эквадор: макрос очистки примечаний — не вставлен ({e})")
        return False

    app = wb.app.api
    wb_name = str(wb.name)
    errors: list[str] = []
    for spec in (
        _CV_CLEAR_NOTES_MACRO,
        f"Module1.{_CV_CLEAR_NOTES_MACRO}",
        f"'{wb_name}'!{_CV_CLEAR_NOTES_MACRO}",
    ):
        try:
            app.Run(spec, first_row, last_row)
            _lg(f"Эквадор: очистка примечаний VBA, строки {first_row}–{last_row}")
            return True
        except Exception as e:
            errors.append(f"{spec}: {e}")
    _lg(
        "Эквадор: макрос очистки примечаний не запустился "
        f"({errors[-1] if errors else '?'}) — Python."
    )
    return False


def _strip_range_annotations(
    sheet: object,
    *,
    first_row: int,
    last_row: int,
    first_col: int = _COL_DATA_FIRST,
    last_col: int = _COL_DATA_LAST,
    wb: object | None = None,
    log: LogFn | None = None,
) -> int:
    if (
        wb is not None
        and first_col == _COL_DATA_FIRST
        and last_col == _COL_DATA_LAST
    ):
        _run_clear_data_notes_macro(
            wb,
            first_row=first_row,
            last_row=last_row,
            log=log,
        )
    return _strip_range_annotations_python(
        sheet,
        first_row=first_row,
        last_row=last_row,
        first_col=first_col,
        last_col=last_col,
    )


def _clear_data_sheet_area(
    sheet: object,
    *,
    wb: object | None = None,
    log: LogFn | None = None,
) -> int:
    """Содержимое и примечания шаблона («0» в углу ячейки) в зоне данных."""
    addr = f"D{_DATA_FIRST_ROW}:AB{_CLEAR_LAST_ROW}"
    sheet.range(addr).clear_contents()
    api_rng = sheet.range(addr).api
    for method in ("ClearNotes", "ClearComments"):
        try:
            getattr(api_rng, method)()
        except Exception:
            pass
    return _strip_range_annotations(
        sheet,
        first_row=_DATA_FIRST_ROW,
        last_row=_CLEAR_LAST_ROW,
        wb=wb,
        log=log,
    )


def _set_cell_value(sheet: object, row: int, col: int, value: object) -> None:
    cell = sheet.api.Cells(row, col)
    _strip_cell_annotation(cell)
    sheet.range((row, col)).value = value


def _set_cell_value_ref(sheet: object, ref: str, value: object) -> None:
    rng = sheet.range(ref)
    _strip_cell_annotation(rng.api)
    rng.value = value


def _write_deal_row(sheet: object, row: int, deal: EcuadorDealRow) -> None:
    _set_cell_value(sheet, row, 4, deal.plantation)
    _set_cell_value(sheet, row, 5, deal.flower_type)
    _set_cell_value(sheet, row, 6, deal.variety)
    _set_cell_value(sheet, row, 15, deal.boxes)
    _set_cell_value(sheet, row, 16, deal.sm)
    _set_cell_value(sheet, row, 17, deal.box_type)
    _set_cell_value(sheet, row, 18, deal.total_stems)
    for col_letter, value in deal.qty_by_length_col.items():
        if value:
            _set_cell_value_ref(sheet, f"{col_letter}{row}", value)


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
            left = float(cell.Left)
            top = float(cell.Top)
            width = float(cell.Width)
            height = float(cell.Height)
            ole = sheet.api.OLEObjects().Add(
                "Forms.CommandButton.1",
                "",
                False,
                False,
                left,
                top,
                width,
                height,
            )
            btn = ole.Object
            _apply_command_button_picture(btn, img, app_api, wb=wb)
            btn.Caption = f"{prefix} {row} 0"

        color = _ZEBRA_EVEN if row % 2 == 0 else _ZEBRA_ODD
        sheet.api.Range(f"D{row}:AB{row}").Interior.Color = color


def _ensure_cv_sync_macro(wb: object) -> None:
    """Public-макрос в Module1 (как SetCommandButton, без вставки колонки A)."""
    vb = wb.api.VBProject
    ensure_vba_references(vb)
    mod = vb.VBComponents("Module1").CodeModule
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
            _CV_SYNC_MACRO,
            f"Module1.{_CV_SYNC_MACRO}",
            f"'{wb_name}'!{_CV_SYNC_MACRO}",
        ):
            try:
                app.Run(spec, first_row, last_row)
                _lg("Эквадор: чекбоксы готовы (VBA).")
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
        _lg("Эквадор: чекбоксы готовы (COM).")
    except Exception as com_err:
        tail = "; ".join(vba_errors[-2:])
        raise RuntimeError(
            "Не удалось создать чекбоксы (VBA и COM). "
            "Проверьте: макросы разрешены для автоматизации, папка шаблона в доверенном "
            "расположении Excel, bmp Red_Check/Green_Check рядом с книгой. "
            f"VBA: {tail}. COM: {com_err}"
        ) from com_err


def _ensure_cv_create_macro(wb: object) -> None:
    _ensure_cv_clear_notes_macro(wb)
    mod = wb.api.VBProject.VBComponents("Module1").CodeModule
    line_count = int(mod.CountOfLines)
    existing = mod.Lines(1, line_count) if line_count else ""
    if f"Sub {_CV_CREATE_MACRO}" in existing:
        return
    mod.InsertLines(line_count + 1, _CV_CREATE_VBA)


def _invoke_create_file_macro(
    wb: object,
    path_sheet: object,
    *,
    output_dir: Path,
    sheet_codename: str,
    log: LogFn | None = None,
) -> Path:
    """Кнопка «Создать файл» — cbCreateTimeFile_Click через public-обёртку."""
    def _lg(msg: str) -> None:
        if log is not None:
            log(msg)

    _lg("Эквадор: «Создать файл» (макрос VBA)…")
    try:
        _ensure_cv_create_macro(wb)
        _run_clear_data_notes_macro(
            wb,
            first_row=_DATA_FIRST_ROW,
            last_row=_CLEAR_LAST_ROW,
            log=log,
        )
    except Exception as e:
        raise RuntimeError(
            "Не удалось добавить макрос «Создать файл». Включите в Excel "
            "«Доверять доступ к объектной модели VBA-проекта». "
            f"({e})"
        ) from e

    save_dir = str(output_dir.resolve())
    app = wb.app.api
    wb_name = str(wb.name)
    errors: list[str] = []
    _ = sheet_codename
    for spec in (
        _CV_CREATE_MACRO,
        f"Module1.{_CV_CREATE_MACRO}",
        f"'{wb_name}'!{_CV_CREATE_MACRO}",
    ):
        try:
            app.Run(spec, save_dir)
            saved_raw = path_sheet.range("B3").value
            if saved_raw:
                out = Path(str(saved_raw)).resolve()
                if out.is_file():
                    _lg(f"Эквадор: сохранено макросом → {out}")
                    return out
            newest = _newest_ecuador_xlsm(output_dir)
            if newest is not None:
                _lg(f"Эквадор: сохранено макросом → {newest}")
                return newest
            raise RuntimeError("макрос завершился, но файл не найден")
        except Exception as e:
            errors.append(f"{spec}: {e}")

    tail = "; ".join(errors[-3:])
    raise RuntimeError(f"Не удалось вызвать «Создать файл» (VBA). ({tail})")


def _save_via_python(
    wb: object,
    *,
    output_dir: Path,
    log: LogFn | None = None,
) -> Path:
    def _lg(msg: str) -> None:
        if log is not None:
            log(msg)

    out_name = ecuador_output_basename()
    out_path = (output_dir / out_name).resolve()
    _lg(f"Эквадор: сохраняю (как «Создать файл») → {out_path}")
    _apply_create_file_ui(wb, out_name)
    wb.api.SaveAs(str(out_path))
    wb.save()
    return out_path


def ecuador_export_candidates(output_dir: Path) -> list[Path]:
    """Файлы «Эквадор …xlsm» в корне папки выгрузки (не в подпапках архива)."""
    if not output_dir.is_dir():
        return []
    found: list[Path] = []
    for entry in output_dir.iterdir():
        if not entry.is_file():
            continue
        if entry.suffix.lower() != ".xlsm":
            continue
        if not entry.name.lower().startswith("эквадор"):
            continue
        if entry.name.startswith("_"):
            continue
        found.append(entry.resolve())
    return sorted(found, key=lambda path: path.stat().st_mtime, reverse=True)


def _newest_ecuador_xlsm(output_dir: Path) -> Path | None:
    items = ecuador_export_candidates(output_dir)
    return items[0] if items else None


def archive_stale_ecuador_exports(
    output_dir: Path,
    archive_dir: Path,
    *,
    keep_path: Path,
    log: LogFn | None = None,
) -> tuple[Path | None, list[str], list[str]]:
    """
    Оставляет один файл Эквадор (keep_path), остальные переносит в папку архива.
    """
    _lg = log or (lambda _msg: None)
    if not output_dir.is_dir():
        return None, [], []

    archive_dir = archive_dir.resolve()
    output_dir = output_dir.resolve()
    keep = keep_path.resolve()
    if archive_dir == output_dir:
        raise ValueError("Папка архива Эквадор не может совпадать с папкой выгрузки.")

    candidates: list[Path] = []
    for path in ecuador_export_candidates(output_dir):
        try:
            path.resolve().relative_to(archive_dir)
        except ValueError:
            candidates.append(path)

    to_move = [path for path in candidates if path.resolve() != keep]
    if not to_move:
        _lg("Эквадор: старых файлов для архива нет")
        return None, [], []

    archive_dir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y-%m-%d_%H%M%S")
    moved: list[str] = []
    warnings: list[str] = []
    if sys.platform == "win32" and not os.access(archive_dir, os.W_OK):
        import getpass

        warnings.append(
            f"Папка архива {archive_dir}: нет записи для «{getpass.getuser()}»."
        )

    for src in sorted(to_move, key=lambda p: p.name.lower()):
        target = _archive_target_path(archive_dir, src, stamp)
        if not os.access(src, os.R_OK):
            warnings.append(f"{src.name}: пропуск (нет чтения)")
            continue
        try:
            warn = _archive_one_entry(src, target)
            moved.append(src.name)
            if warn:
                warnings.append(warn)
        except OSError as e:
            warnings.append(f"{src.name}: не удалось архивировать — {e}")

    if moved:
        _lg(f"Эквадор: в архив {len(moved)} → {archive_dir}")
    for warn in warnings:
        _lg(f"Эквадор (архив): {warn}")
    return (archive_dir if moved else None), moved, warnings


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
    use_create_file_macro: bool | None = None,
    log: LogFn | None = None,
) -> Path:
    """
    Преобразование в Python, запись в шаблон .xlsm, сохранение как «Создать файл».
    По умолчанию SaveAs из Python; VBA — только если use_create_file_macro=True в config.
    Только Windows + установленный Excel (xlwings).
    """
    if sys.platform != "win32":
        raise RuntimeError("Создание файла Эквадор доступно только на Windows с Excel.")

    biflorica_path = biflorica_path.resolve()
    if not biflorica_path.is_file():
        raise FileNotFoundError(biflorica_path)

    from cvetopt.core.runtime_settings import load_runtime_settings

    runtime = load_runtime_settings(env)
    yaml_ecuador = env.yaml_config().ecuador_create
    template = template_path or resolve_ecuador_template(env, runtime.ecuador_template_path)
    out_dir = output_dir or resolve_ecuador_output_dir(env, runtime.ecuador_output_dir)
    if use_create_file_macro is None:
        env_flag = os.environ.get("ECUADOR_USE_CREATE_FILE_MACRO", "").strip().lower()
        if env_flag in ("1", "true", "yes"):
            use_create_file_macro = True
        else:
            use_create_file_macro = yaml_ecuador.use_create_file_macro

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

    _lg(f"Эквадор: сделок {len(deals)}, шаблон {template.name}")
    for deal in deals:
        if "|" in deal.total_stems or "|" in deal.sm:
            dist = ", ".join(
                f"{col}={val}" for col, val in deal.qty_by_length_col.items() if val
            )
            _lg(
                f"Эквадор: {deal.variety!r} СМ={deal.sm!r} "
                f"ВСЕГО={deal.total_stems!r} → {dist or 'пусто'}"
            )

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
    work_copy: Path | None = None
    try:
        app = xw.App(visible=visible, add_book=False)
        _configure_excel_app(app)

        work_copy = _prepare_work_copy(template)
        assets_dir = work_copy.parent

        _lg(f"Эквадор: открываю {work_copy.name} (Path!A2=False, папка шаблона)…")
        wb = _open_workbook_quiet(app, work_copy)
        _lg(f"Эквадор: шаблон открыт — {wb.name}")

        data_sheet = wb.sheets[_SHEET_DATA]
        path_sheet = wb.sheets[_SHEET_PATH]
        sheet_codename = str(data_sheet.api.CodeName)
        path_sheet.range("A2").value = "False"

        _lg("Эквадор: заполняю строки…")
        removed_tpl = _clear_data_sheet_area(data_sheet, wb=wb, log=log)
        if removed_tpl:
            _lg(f"Эквадор: снято примечаний шаблона (Python): {removed_tpl}")

        last_row = _DATA_FIRST_ROW + len(deals) - 1
        app_api = app.api
        app_api.EnableEvents = False
        try:
            for idx, deal in enumerate(deals):
                _write_deal_row(data_sheet, _DATA_FIRST_ROW + idx, deal)
            removed_fill = _strip_range_annotations(
                data_sheet,
                first_row=_DATA_FIRST_ROW,
                last_row=last_row,
                wb=wb,
                log=log,
            )
            if removed_fill:
                _lg(f"Эквадор: снято примечаний после заполнения (Python): {removed_fill}")
        finally:
            app_api.EnableEvents = True

        path_sheet.range("A1").value = str(biflorica_path)
        path_sheet.range("B1").value = biflorica_path.name

        try:
            _sync_row_checkboxes(
                wb,
                data_sheet,
                first_row=_DATA_FIRST_ROW,
                last_row=last_row,
                assets_dir=assets_dir,
                log=log,
            )
        except Exception as e:
            _lg(f"Эквадор: чекбоксы не созданы ({e}) — продолжаю без них.")

        app_api.EnableEvents = False
        try:
            removed_after = _strip_range_annotations(
                data_sheet,
                first_row=_DATA_FIRST_ROW,
                last_row=last_row,
                wb=wb,
                log=log,
            )
            if removed_after:
                _lg(f"Эквадор: снято примечаний после чекбоксов (Python): {removed_after}")
        finally:
            app_api.EnableEvents = True

        removed_final = _strip_range_annotations(
            data_sheet,
            first_row=_DATA_FIRST_ROW,
            last_row=last_row,
            wb=wb,
            log=log,
        )
        if removed_final:
            _lg(f"Эквадор: снято примечаний перед сохранением (Python): {removed_final}")

        if use_create_file_macro:
            try:
                _ensure_cv_clear_notes_macro(wb)
                out_path = _invoke_create_file_macro(
                    wb,
                    path_sheet,
                    output_dir=out_dir,
                    sheet_codename=sheet_codename,
                    log=log,
                )
            except Exception as e:
                _lg(f"Эквадор: макрос VBA недоступен ({e!s:.200}) — SaveAs из Python.")
                out_path = _save_via_python(wb, output_dir=out_dir, log=log)
        else:
            out_path = _save_via_python(wb, output_dir=out_dir, log=log)

        _copy_checkbox_assets(template.parent, out_path.parent)

        for book in list(app.books):
            try:
                book.close()
            except Exception:
                pass
        wb = None
        _lg(f"Эквадор: файл создан → {out_path}")

        if yaml_ecuador.archive_previous_on_create:
            archive_dir = resolve_ecuador_archive_dir(
                env,
                runtime.ecuador_archive_dir,
                out_dir,
                runtime=runtime,
            )
            _lg(f"Эквадор: архив {archive_dir}")
            try:
                archive_stale_ecuador_exports(
                    out_dir,
                    archive_dir,
                    keep_path=out_path.resolve(),
                    log=log,
                )
            except Exception as e:
                _lg(f"Эквадор: архив старых файлов пропущен — {e}")

        return out_path
    finally:
        if wb is not None:
            try:
                wb.close()
            except Exception:
                pass
        if work_copy is not None:
            for attempt in range(5):
                try:
                    work_copy.unlink(missing_ok=True)
                    break
                except OSError:
                    time.sleep(0.3)
        if app is not None:
            try:
                app.quit()
            except Exception:
                pass
