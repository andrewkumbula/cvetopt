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
    order_id_from_biflorica_report,
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
Public Sub cv_Run_CreateFile(Optional aSaveDir As String = "")
    Dim aName As String
    Dim aPath As String
    Dim saveDir As String
    saveDir = aSaveDir
    If saveDir = "" Then saveDir = aPathToWrite
    aName = "Эквадор " & Format(Now(), "d/m/yy hh.nn") & ".xlsm"
    aPath = saveDir & "\\" & aName
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


def _load_picture(app_api: object, image_path: str) -> object:
    path = str(Path(image_path).resolve())
    try:
        return app_api.LoadPicture(path)
    except Exception:
        pass
    try:
        import win32com.client

        return win32com.client.Dispatch(app_api).LoadPicture(path)
    except Exception as exc:
        raise RuntimeError(f"LoadPicture({path}): {exc}") from exc


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
            btn.Picture = _load_picture(app_api, img)
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
            _CV_SYNC_MACRO,
            f"Module1.{_CV_SYNC_MACRO}",
            f"'{wb_name}'!{_CV_SYNC_MACRO}",
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


def _ensure_cv_create_macro(wb: object, sheet_codename: str) -> None:
    mod = wb.api.VBProject.VBComponents(sheet_codename).CodeModule
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

    _lg("Эквадор: «Создать файл» (макрос)…")
    try:
        _ensure_cv_create_macro(wb, sheet_codename)
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
    for spec in (
        f"'{wb_name}'!{_CV_CREATE_MACRO}",
        f"'{sheet_codename}'.{_CV_CREATE_MACRO}",
        _CV_CREATE_MACRO,
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
    _lg(f"Эквадор: сохраняю (Python SaveAs) → {out_path}")
    _apply_create_file_ui(wb, out_name)
    wb.api.SaveAs(str(out_path))
    wb.save()
    return out_path


def _newest_ecuador_xlsm(output_dir: Path) -> Path | None:
    items = [
        path
        for path in output_dir.glob("Эквадор*.xlsm")
        if path.is_file()
    ]
    if not items:
        return None
    return max(items, key=lambda path: path.stat().st_mtime)


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
    use_create_file_macro: bool = True,
    log: LogFn | None = None,
) -> Path:
    """
    Преобразование в Python, запись в шаблон .xlsm, затем «Создать файл» (VBA) или SaveAs.
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
        data_sheet.range(f"D{_DATA_FIRST_ROW}:AB{_CLEAR_LAST_ROW}").clear_contents()
        last_row = _DATA_FIRST_ROW + len(deals) - 1
        for idx, deal in enumerate(deals):
            _write_deal_row(data_sheet, _DATA_FIRST_ROW + idx, deal)

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

        if use_create_file_macro:
            try:
                out_path = _invoke_create_file_macro(
                    wb,
                    path_sheet,
                    output_dir=out_dir,
                    sheet_codename=sheet_codename,
                    log=log,
                )
            except Exception as e:
                _lg(f"Эквадор: макрос «Создать файл» не сработал ({e}) — SaveAs из Python.")
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
