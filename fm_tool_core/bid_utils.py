from __future__ import annotations

import logging
from pathlib import Path
from typing import Any, Iterable

import openpyxl

_COLUMNS = [
    "Lane ID",
    "Origin City",
    "Orig State",
    "Orig Zip (5 or 3)",
    "Destination City",
    "Dest State",
    "Dest Zip (5 or 3)",
    "Bid Volume",
    "LH Rate",
    "Bid Miles",
    "Miles",
    "Tolls",
]

_REQUIRED = {
    "Lane ID",
    "Orig Zip (5 or 3)",
    "Dest Zip (5 or 3)",
}


def insert_bid_rows(
    wb_path: Path, rows: Iterable[dict[str, Any]], log: logging.Logger
) -> None:
    """Append valid *rows* to the BID table of ``wb_path``."""
    wb = openpyxl.load_workbook(wb_path, keep_vba=True)
    try:
        ws = wb["BID"]
    except KeyError:
        log.error("BID sheet not found in %s", wb_path)
        return

    for data in rows:
        if not _REQUIRED.issubset(k for k in data if data[k] is not None):
            continue
        ws.append([data.get(col, "") for col in _COLUMNS])

    wb.save(wb_path)
    try:
        wb.close()
    except Exception:
        pass


__all__ = ["insert_bid_rows"]
