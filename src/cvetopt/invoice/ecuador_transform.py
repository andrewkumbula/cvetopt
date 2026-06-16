from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from cvetopt.invoice.xlsx_read import grid_by_row, read_xlsx_grid

# Biflorica: цены по длинам в E–L; на листе обработки количества в S–Z.
_LENGTHS: list[tuple[str, str, str]] = [
    ("40", "E", "S"),
    ("50", "F", "T"),
    ("60", "G", "U"),
    ("70", "H", "V"),
    ("80", "I", "W"),
    ("90", "J", "X"),
    ("100", "K", "Y"),
    ("100+", "L", "Z"),
]


@dataclass(frozen=True)
class EcuadorDealRow:
    plantation: str
    flower_type: str
    variety: str
    boxes: str
    box_type: str
    total_stems: str
    sm: str
    qty_by_length_col: dict[str, str]


def _has_price(value: str) -> bool:
    text = (value or "").strip()
    return text != "" and text != "0"


def _find_header_row(rows: dict[int, dict[str, str]]) -> int:
    for row_no in sorted(rows):
        row = rows[row_no]
        if row.get("B") == "ПЛАНТАЦИЯ" or row.get("A") == "ДАТА И ВРЕМЯ СДЕЛКИ":
            return row_no
    return 6


def transform_biflorica_deals(path: Path) -> list[EcuadorDealRow]:
    grid = read_xlsx_grid(path)
    rows = grid_by_row(grid)
    header_row = _find_header_row(rows)
    deals: list[EcuadorDealRow] = []

    for row_no in sorted(rows):
        if row_no <= header_row:
            continue
        row = rows[row_no]
        plantation = (row.get("B") or "").strip()
        if not plantation:
            continue

        priced_labels: list[str] = []
        for label, src_col, _out_col in _LENGTHS:
            if _has_price(row.get(src_col, "")):
                priced_labels.append(label)

        sm = "|".join(priced_labels)
        stems = (row.get("O") or "").strip()
        qty_cols: dict[str, str] = {out_col: "" for _label, _src, out_col in _LENGTHS}

        if priced_labels and stems:
            out_col_by_label = {label: out_col for label, _src, out_col in _LENGTHS}
            stem_parts = [p.strip() for p in stems.split("|")]
            if len(priced_labels) > 1 and len(stem_parts) == len(priced_labels):
                # «175|175» при СМ «60|70» → 175 в колонку 60, 175 в колонку 70.
                for label, part in zip(priced_labels, stem_parts):
                    if part:
                        qty_cols[out_col_by_label[label]] = part
            else:
                # Одна длина (или не делится поровну) — всё в первую длину с ценой.
                qty_cols[out_col_by_label[priced_labels[0]]] = stems

        deals.append(
            EcuadorDealRow(
                plantation=plantation,
                flower_type=(row.get("C") or "").strip(),
                variety=(row.get("D") or "").strip(),
                boxes=(row.get("M") or "").strip(),
                box_type=(row.get("N") or "").strip(),
                total_stems=stems,
                sm=sm,
                qty_by_length_col=qty_cols,
            )
        )
    return deals
