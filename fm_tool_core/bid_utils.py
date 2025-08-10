# fm_tool_core/bid_utils.py
from __future__ import annotations

import logging
from itertools import chain
from pathlib import Path
from typing import Any, Iterable

from .constants import VISIBLE_EXCEL
from .excel_utils import pythoncom, xw

###############################################################################
# Column layout: 25 columns (A → Y)                                           #
###############################################################################
_COLUMNS = [
    "LANE_ID",
    "ORIG_CITY",
    "ORIG_ST",
    "ORIG_POSTAL_CD",
    "DEST_CITY",
    "DEST_ST",
    "DEST_POSTAL_CD",
    "BID_VOLUME",
    "LH_RATE",
    "RFP_MILES",
    "FREIGHT_TYPE",
    "TEMP_CAT",
    "BTF_FSC_PER_MILE",
    "ADHOC_INFO1",
    "ADHOC_INFO2",
    "ADHOC_INFO3",
    "ADHOC_INFO4",
    "ADHOC_INFO5",
    "ADHOC_INFO6",
    "ADHOC_INFO7",
    "ADHOC_INFO8",
    "ADHOC_INFO9",
    "ADHOC_INFO10",
    "FM_MILES",
    "FM_TOLLS",
]

_REQUIRED = {
    "LANE_ID",
    "ORIG_POSTAL_CD",
    "DEST_POSTAL_CD",
}

_TARGET_SHEET = "RFP"  # ← changed from “BID”


def insert_bid_rows(
    wb_path: Path,
    rows: Iterable[dict[str, Any]],
    log: logging.Logger,
    adhoc_headers: dict[str, str] | None = None,
) -> None:
    """
    Bulk-insert BID/RFP *rows* into the RFP sheet of *wb_path*.

    If *adhoc_headers* is provided, header labels matching its keys are
    replaced with the corresponding values before inserting any data rows.
    """
    row_iter = iter(rows)
    try:
        first = next(row_iter)
    except StopIteration:
        log.info("No RFP rows to insert")
        return

    rows = chain([first], row_iter)

    data: list[list[Any]] = []
    for rec in rows:
        if _REQUIRED.issubset(k for k in rec if rec[k] is not None):
            data.append([rec.get(col) or "" for col in _COLUMNS])

    if not data:
        log.info("No RFP rows to insert after validation")
        return

    if xw is None:
        log.error("xlwings is required for RFP inserts")
        return

    pythoncom.CoInitialize()
    app = xw.App(visible=VISIBLE_EXCEL, add_book=False)  # type: ignore
    app.api.DisplayAlerts = False
    wb = None
    try:
        wb = app.books.open(str(wb_path))
        try:
            ws = wb.sheets[_TARGET_SHEET]
        except Exception:
            log.error("%s sheet not found in %s", _TARGET_SHEET, wb_path)
            return

        if adhoc_headers:
            header_rng = ws.range((1, 1)).resize(1, len(_COLUMNS))
            values = header_rng.value
            if isinstance(values, list):
                nested = bool(values and isinstance(values[0], list))
                row = values[0] if nested else values
                mapped_row = [adhoc_headers.get(str(v), v) for v in row]
                header_rng.value = [mapped_row] if nested else mapped_row

        # First empty row in column A
        start_row = ws.api.Cells(ws.api.Rows.Count, 1).End(-4162).Row + 1
        n_rows = len(data)
        n_cols = len(_COLUMNS)

        # One-shot write
        ws.range((start_row, 1)).resize(n_rows, n_cols).value = data
        wb.save()
        log.info(
            "Wrote %d rows × %d cols to %s sheet",
            n_rows,
            n_cols,
            _TARGET_SHEET,
        )
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


__all__ = ["insert_bid_rows", "_COLUMNS"]
