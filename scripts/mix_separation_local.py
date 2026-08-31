#!/usr/bin/env python3
"""Локальная проверка разбора миксов на шаблоны тест/ + архивный Biflorica."""

from __future__ import annotations

import shutil
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from cvetopt.invoice.biflorica_mixes import (  # noqa: E402
    parse_sklad_template,
    plan_mix_allocation,
    run_mix_separation,
)
from cvetopt.invoice.xlsx_read import grid_by_row, read_xlsx_grid  # noqa: E402

TEMPLATE = ROOT / "шаблоны тест" / "шаюлон 27,7,26.xlsx"
BIF_SRC = (
    ROOT
    / "data/downloads/biflorica/архив/2026-06-03_121632/10738107__2026-05-30.xlsx"
)
OUT_DIR = ROOT / "testdata" / "mixes"


def main() -> int:
    if not TEMPLATE.is_file():
        print(f"FAIL: нет шаблона {TEMPLATE}")
        return 1
    if not BIF_SRC.is_file():
        print(f"FAIL: нет Biflorica {BIF_SRC}")
        return 1

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    bif_copy = OUT_DIR / "BiFlorica-mix-test.xlsx"
    shutil.copy2(BIF_SRC, bif_copy)

    demand = parse_sklad_template(TEMPLATE)
    assert demand.totals["50"] == 100, demand.totals
    assert demand.totals["60"] == 550, demand.totals
    filled60 = [ln for ln in demand.lines if ln.qtys.get("60", 0) > 0]
    assert len(filled60) == 6, [ln.code for ln in filled60]

    plans = plan_mix_allocation(demand, bif_copy, log=print)
    assert len(plans) == 2
    by_len = {p.length: p for p in plans}
    assert abs(by_len["60"].avg_price - 0.2) < 1e-9
    assert by_len["60"].need == 550
    assert len(by_len["60"].lines) == 6
    # 50: самый дорогой Mix @ 0.22
    assert abs(by_len["50"].avg_price - 0.22) < 1e-9

    run_mix_separation(template_path=TEMPLATE, biflorica_path=bif_copy, log=print)

    rows = grid_by_row(read_xlsx_grid(bif_copy))
    new_codes = []
    for rn, cells in rows.items():
        if (cells.get("C") or "") == "Роза" and (cells.get("B") or "") == "":
            code = (cells.get("D") or "").strip()
            if code and code.lower() != "mix":
                new_codes.append(
                    (
                        code,
                        cells.get("F"),  # 50
                        cells.get("G"),  # 60
                        cells.get("O"),
                    )
                )

    codes60 = {c for c, f, g, o in new_codes if g}
    expected60 = {ln.code for ln in filled60}
    assert codes60 == expected60, (codes60, expected60)

    # остаток Mix 60: было 600, взяли 550 → 50
    mix60_left = 0
    for rn, cells in rows.items():
        if (cells.get("D") or "").strip().lower() != "mix":
            continue
        if cells.get("G"):
            mix60_left += int(float(str(cells.get("O") or "0").replace(",", ".")))
    assert mix60_left == 50, mix60_left

    print("VERIFY OK")
    print(f"  template={TEMPLATE.name} totals={demand.totals}")
    print(f"  biflorica={bif_copy}")
    print(f"  new rows={len(new_codes)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
