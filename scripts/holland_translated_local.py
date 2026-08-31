#!/usr/bin/env python3
"""
Локальный плацдарм для кнопки «Переведено» (без Windows/Excel Auto1).

Использование из корня репо:

  # сбросить work/ из pristine/ и прогнать перевод+ростовку
  .venv/bin/python scripts/holland_translated_local.py

  # только сбросить файлы (не гонять логику)
  .venv/bin/python scripts/holland_translated_local.py --reset-only

  # только проверить уже обработанные work/
  .venv/bin/python scripts/holland_translated_local.py --verify-only
"""

from __future__ import annotations

import argparse
import shutil
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
BASE = ROOT / "testdata" / "holland_translated"
PRISTINE = BASE / "pristine"
WORK = BASE / "work"
DICT = BASE / "Словарь.xls"


def _ensure_src_on_path() -> None:
    src = ROOT / "src"
    if str(src) not in sys.path:
        sys.path.insert(0, str(src))


def reset_work() -> None:
    if not PRISTINE.is_dir():
        raise SystemExit(f"Нет pristine: {PRISTINE}")
    if WORK.exists():
        shutil.rmtree(WORK)
    for slot in ("1", "2"):
        src = PRISTINE / slot
        dst = WORK / slot
        dst.mkdir(parents=True, exist_ok=True)
        files = list(src.glob("*.xls")) + list(src.glob("*.XLS")) + list(src.glob("*.xlsx"))
        if not files:
            raise SystemExit(f"В {src} нет Excel")
        for f in files:
            shutil.copy2(f, dst / f.name)
        print(f"reset: {slot}/ ← {files[0].name}")


def run_process() -> None:
    _ensure_src_on_path()
    from cvetopt.invoice.holland_mail_continue import process_mail_folders

    if not DICT.is_file():
        raise SystemExit(f"Нет словаря: {DICT}")
    print(f"словарь: {DICT}")
    r1, r2 = process_mail_folders(
        WORK / "1",
        WORK / "2",
        DICT,
        append_missing_to_dictionary=False,
        log=print,
    )
    print(
        f"итог папка1: строк={r1.total_rows} перевод={r1.translated} роз={r1.roses_updated}"
    )
    print(
        f"итог папка2: строк={r2.total_rows} перевод={r2.translated} роз={r2.roses_updated}"
    )


def _scan_folder(
    folder: Path,
    *,
    desc_col: int,
    length_col: int,
    art_col: int | None = None,
) -> bool:
    import xlrd
    from cvetopt.invoice.holland_mail_continue import (
        is_rose_description,
        length_without_cm,
    )

    files = list(folder.glob("*.xls")) + list(folder.glob("*.XLS"))
    if not files:
        raise SystemExit(f"Нет файлов в {folder}")
    path = files[0]
    book = xlrd.open_workbook(str(path))
    sh = book.sheet_by_index(0)
    roses_ok = roses_bad = non_rose_ok = non_rose_bad = 0
    samples_rose: list[str] = []
    samples_bad: list[str] = []

    for row in range(sh.nrows):
        raw = sh.cell_value(row, desc_col)
        if raw == "" or str(raw).strip().lower() == "description":
            continue
        desc = " ".join(str(raw).split()).strip()
        if not desc or desc.lower() in {"subtotal", "artnr"}:
            continue
        # Только товарные строки (есть артикул-число), не Subtotal / Declaration costs.
        if art_col is not None:
            art = sh.cell_value(row, art_col)
            art_s = str(art).strip()
            if not art_s or art_s.lower() in {"artnr", "subtotal", "total"}:
                continue
            if isinstance(art, str) and not art_s.replace(".", "", 1).isdigit():
                # текстовый артикул допустим (коды вида 128935 уже числа в xls)
                if not any(ch.isdigit() for ch in art_s):
                    continue
        length = length_without_cm(sh.cell_value(row, length_col))
        rose = is_rose_description(desc)
        tokens = desc.split()
        has_len = bool(tokens) and length and tokens[-1] == length
        if rose:
            if has_len:
                roses_ok += 1
                if len(samples_rose) < 3:
                    samples_rose.append(desc)
            else:
                roses_bad += 1
                if len(samples_bad) < 5:
                    samples_bad.append(desc)
        else:
            if has_len and length:
                non_rose_bad += 1
                if len(samples_bad) < 5:
                    samples_bad.append(f"[не-роза+длина?] {desc}")
            else:
                non_rose_ok += 1

    print(f"\n=== {folder.name}: {path.name} ===")
    print(f"розы с ростовкой: {roses_ok}, розы без: {roses_bad}")
    print(f"не-розы ок: {non_rose_ok}, не-розы с длиной в конце: {non_rose_bad}")
    for s in samples_rose:
        print(f"  ok rose: {s}")
    for s in samples_bad:
        print(f"  bad: {s}")
    return roses_bad == 0 and non_rose_bad == 0 and roses_ok > 0


def verify(*, expect_processed: bool) -> bool:
    _ensure_src_on_path()
    # папка 1: L=11, T=19, Box/нет строгого art — F=Art; папка 2: A=Artnr, F=Desc, K=S1
    ok1 = _scan_folder(WORK / "1", desc_col=11, length_col=19, art_col=5)
    ok2 = _scan_folder(WORK / "2", desc_col=5, length_col=10, art_col=0)
    if expect_processed:
        if ok1 and ok2:
            print("\nVERIFY OK")
            return True
        print("\nVERIFY FAIL")
        return False
    print("\n(verify без ожидания ростовки — после --reset-only)")
    return True


def main() -> int:
    ap = argparse.ArgumentParser(description="Локальный тест «Переведено»")
    ap.add_argument("--reset-only", action="store_true")
    ap.add_argument("--verify-only", action="store_true")
    args = ap.parse_args()

    if args.verify_only:
        return 0 if verify(expect_processed=True) else 1

    reset_work()
    if args.reset_only:
        verify(expect_processed=False)
        print(f"\nwork готов: {WORK}")
        print(f"словарь:   {DICT}")
        return 0

    run_process()
    return 0 if verify(expect_processed=True) else 1


if __name__ == "__main__":
    raise SystemExit(main())
