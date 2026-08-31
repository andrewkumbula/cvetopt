"""
Разделение отчёта Biflorica по типу цветка (аналог макроса «Гипсофила» + «Да»).

Создаёт рядом с исходником:
  <имя> Гипсофила.xlsx
  <имя> Роза.xlsx
Исходный полный файл не удаляет.
"""

from __future__ import annotations

import re
import shutil
from collections.abc import Callable
from dataclasses import dataclass
from pathlib import Path

from openpyxl import load_workbook

from cvetopt.core.runtime_settings import BIFLORICA_DOWNLOAD_PREFIX
from cvetopt.invoice.xlsx_read import grid_by_row, read_xlsx_grid

LogFn = Callable[[str], None]

_SPLIT_SUFFIXES = ("Гипсофила", "Роза", "Прочее", "Гортензия")
_TYPE_COL = "C"
_DATA_FIRST_FALLBACK = 7


def _default_log(_msg: str) -> None:
    pass


def _norm(text: object) -> str:
    return " ".join(str(text or "").split()).strip()


def is_split_output_name(name: str) -> bool:
    """Уже разделённый файл: «… Гипсофила.xlsx» / «… Роза.xlsx»."""
    stem = Path(name).stem
    for suf in _SPLIT_SUFFIXES:
        if stem.endswith(f" {suf}") or stem.casefold().endswith(f" {suf.casefold()}"):
            return True
    return False


def classify_flower_type(type_name: str) -> str | None:
    """
    Возвращает ярлык файла-выгрузки или None (остаётся только в полном отчёте).
    Гипсофила → «Гипсофила»; любая роза → «Роза».
    """
    t = _norm(type_name).casefold()
    if not t:
        return None
    if "гипсофил" in t:
        return "Гипсофила"
    if "роз" in t:  # Роза, Крашеная роза
        return "Роза"
    return None


def _find_header_row(rows: dict[int, dict[str, str]]) -> int:
    for row_no in sorted(rows):
        row = rows[row_no]
        if row.get("B") == "ПЛАНТАЦИЯ" or row.get("A") == "ДАТА И ВРЕМЯ СДЕЛКИ":
            return row_no
    return _DATA_FIRST_FALLBACK - 1


def find_latest_biflorica_report(download_dir: Path) -> Path:
    """Самый свежий полный BiFlorica-*.xlsx в корне папки (не архив, не сплит)."""
    if not download_dir.is_dir():
        raise FileNotFoundError(f"Папка Biflorica не найдена: {download_dir}")

    files: list[Path] = []
    for path in download_dir.iterdir():
        if not path.is_file():
            continue
        if path.suffix.lower() != ".xlsx":
            continue
        if is_split_output_name(path.name):
            continue
        name = path.name
        if not (
            name.startswith(BIFLORICA_DOWNLOAD_PREFIX)
            or name.lower().startswith("biflorica")
            or re.match(r"^\d+__", path.stem)
        ):
            continue
        files.append(path)

    if not files:
        raise FileNotFoundError(
            f"В {download_dir} нет полного отчёта BiFlorica-*.xlsx. "
            "Сначала скачайте отчёт или укажите файл явно."
        )
    files.sort(key=lambda p: p.stat().st_mtime, reverse=True)
    return files[0].resolve()


@dataclass(frozen=True)
class BifloricaSplitResult:
    source: Path
    outputs: dict[str, Path]  # label → path
    counts: dict[str, int]  # label → data rows


def split_biflorica_by_type(
    source: Path,
    *,
    output_dir: Path | None = None,
    log: LogFn | None = None,
) -> BifloricaSplitResult:
    """
    Копирует source в «… Гипсофила.xlsx» / «… Роза.xlsx», оставляя в каждом
    шапку + строки нужного типа (колонка C «ТИП»).
    """
    _lg = log or _default_log
    source = source.resolve()
    if not source.is_file():
        raise FileNotFoundError(f"Файл не найден: {source}")
    if is_split_output_name(source.name):
        raise ValueError(f"Файл уже выглядит как результат сплита: {source.name}")

    out_dir = (output_dir or source.parent).resolve()
    out_dir.mkdir(parents=True, exist_ok=True)

    grid = read_xlsx_grid(source)
    rows = grid_by_row(grid)
    header_row = _find_header_row(rows)
    _lg(f"Гипсофила: источник {source.name}, шапка строка {header_row}")

    # 1-based Excel rows belonging to each label
    by_label: dict[str, list[int]] = {"Гипсофила": [], "Роза": []}
    other = 0
    for row_no in sorted(rows):
        if row_no <= header_row:
            continue
        typ = rows[row_no].get(_TYPE_COL, "")
        if not _norm(typ) and not _norm(rows[row_no].get("B", "")):
            continue
        label = classify_flower_type(typ)
        if label is None:
            other += 1
            continue
        by_label[label].append(row_no)

    outputs: dict[str, Path] = {}
    counts: dict[str, int] = {}

    for label, keep_rows in by_label.items():
        dest = out_dir / f"{source.stem} {label}.xlsx"
        if dest.exists():
            dest.unlink()
        shutil.copy2(source, dest)
        _filter_workbook_rows(dest, header_row=header_row, keep_data_rows=set(keep_rows))
        outputs[label] = dest
        counts[label] = len(keep_rows)
        _lg(f"Гипсофила: {label} → {dest.name} ({len(keep_rows)} строк)")

    if other:
        _lg(f"Гипсофила: прочих типов (только в полном файле): {other}")

    return BifloricaSplitResult(source=source, outputs=outputs, counts=counts)


def _filter_workbook_rows(
    path: Path,
    *,
    header_row: int,
    keep_data_rows: set[int],
) -> None:
    """Удаляет строки данных, не входящие в keep_data_rows (1-based)."""
    wb = load_workbook(path)
    ws = wb.worksheets[0]
    # Удаляем снизу вверх, чтобы индексы не съезжали.
    max_row = ws.max_row or header_row
    for row_no in range(max_row, header_row, -1):
        if row_no not in keep_data_rows:
            ws.delete_rows(row_no, 1)
    wb.save(path)
    wb.close()


def run_gypsophila_split(
    download_dir: Path,
    *,
    source: Path | None = None,
    log: LogFn | None = None,
) -> BifloricaSplitResult:
    """Находит свежий Biflorica (или берёт source) и пишет Гипсофила/Роза."""
    _lg = log or _default_log
    src = source.resolve() if source is not None else find_latest_biflorica_report(download_dir)
    _lg(f"Гипсофила: обрабатываю {src}")
    return split_biflorica_by_type(src, output_dir=src.parent, log=_lg)
