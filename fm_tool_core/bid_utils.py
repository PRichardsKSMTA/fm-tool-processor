from __future__ import annotations

import logging
from itertools import chain
from pathlib import Path
from typing import Any, Iterable

from .constants import VISIBLE_EXCEL
from .excel_utils import pythoncom, xw

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
    """Append valid *rows* to the BID table of ``wb_path`` using COM."""
    row_iter = iter(rows)
    try:
        first = next(row_iter)
    except StopIteration:
        log.info("No BID rows to insert")
        return
    rows = chain([first], row_iter)

    data: list[list[Any]] = []
    for rec in rows:
        if _REQUIRED.issubset(k for k in rec if rec[k] is not None):
            data.append([rec.get(col, "") for col in _COLUMNS])
    if not data:
        log.info("No BID rows to insert")
        return

    if xw is None:
        log.error("xlwings is required for BID inserts")
        return

    pythoncom.CoInitialize()
    app = xw.App(visible=VISIBLE_EXCEL, add_book=False)  # type: ignore
    app.api.DisplayAlerts = False
    wb = None
    try:
        wb = app.books.open(str(wb_path))
        try:
            ws = wb.sheets["BID"]
        except Exception:
            log.error("BID sheet not found in %s", wb_path)
            return
        was_protected = bool(getattr(ws.api, "ProtectContents", False))
        if was_protected:
            ws.api.Unprotect()
        start = ws.api.Cells(ws.api.Rows.Count, 1).End(-4162).Row + 1
        chunk = 500
        idx = 0
        while idx < len(data):
            block = data[idx : idx + chunk]
            end = start + len(block) - 1
            ws.range(f"A{start}:M{end}").value = block
            start = end + 1
            idx += chunk
        if was_protected:
            ws.api.Protect()
        wb.save()
    finally:
        if wb is not None:
            try:
                wb.close()
            except Exception:
                pass
        try:
            app.kill()
        except Exception:
            pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


__all__ = ["insert_bid_rows"]
