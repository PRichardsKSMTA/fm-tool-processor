from __future__ import annotations

import logging
import shutil
import threading
import time
from pathlib import Path
from typing import Any

from .constants import (
    OPEN_TO,
    POLL_SLEEP,
    READY_ERR,
    READY_NAME,
    READY_OK,
    READY_TO,
    SCAC_VALIDATION_SHEET,
    VISIBLE_EXCEL,
)
from .exceptions import FlowError

# psutil may be unavailable in tests
try:  # pragma: no cover - platform specific
    import psutil
except Exception:  # pragma: no cover - missing dependency

    class _PS:
        @staticmethod
        def process_iter(attrs=None):
            return []

    psutil = _PS()

# xlwings may be unavailable in tests
try:  # pragma: no cover - heavy dependency
    import xlwings as xw
except Exception:  # pragma: no cover - missing dependency
    xw = None  # type: ignore

# win32com is Windows-only; provide no-op fallback
try:  # pragma: no cover - platform specific
    from win32com.client import pythoncom  # type: ignore
except Exception:  # pragma: no cover - non-Windows

    class _PC:
        @staticmethod
        def PumpWaitingMessages(): ...

        @staticmethod
        def CoInitialize(): ...

        @staticmethod
        def CoUninitialize(): ...

    pythoncom = _PC()


def kill_orphan_excels():
    for p in psutil.process_iter(attrs=["name"]):
        if p.info.get("name", "").lower().startswith("excel"):
            try:
                p.kill()
            except Exception:
                pass


def copy_template(src: str, dst_root: str, new_name: str, log: logging.Logger) -> Path:
    src_path = Path(src)
    if not src_path.exists():
        raise FlowError(f"Template not found: {src}", work_completed=False)

    dst_root = Path(dst_root).expanduser().resolve()
    dst_root.mkdir(parents=True, exist_ok=True)
    dst = dst_root / new_name

    try:
        shutil.copy2(src_path, dst)
    except PermissionError:
        log.warning("Destination locked; unlinking and retrying copy")
        try:
            dst.unlink()
        except Exception:
            pass
        shutil.copy2(src_path, dst)

    return dst


def wait_ready(wb, log: logging.Logger):
    start = time.time()
    last_log = -999
    while True:
        try:
            ref = wb.names[READY_NAME].refers_to
            if ref.startswith("="):
                ref = ref[1:]
            flag = ref.strip('"').strip()
        except Exception:
            flag = "<not-set>"

        elapsed = int(time.time() - start)
        if elapsed != last_log:
            log.info("Polling flag: %s (t=%ss)", flag, elapsed)
            last_log = elapsed

        if flag == READY_OK:
            return
        if flag == READY_ERR:
            raise FlowError("VBA signaled ERROR", work_completed=False)
        if elapsed >= READY_TO:
            raise FlowError("Timeout waiting for READY flag", work_completed=False)
        time.sleep(POLL_SLEEP)


def _open_excel_with_timeout(path: Path, log: logging.Logger):
    if xw is None:
        raise FlowError("xlwings is required", work_completed=False)

    log.info("Creating Excel App …")
    app = xw.App(visible=VISIBLE_EXCEL, add_book=False)  # type: ignore
    app.api.DisplayFullScreen = False
    app.api.Application.DisplayAlerts = False

    log.info("Opening workbook …")
    t0 = time.time()
    while True:
        try:
            book = app.books.open(str(path))
            return app, book
        except Exception as e:
            if time.time() - t0 > OPEN_TO:
                log.error("Excel open timed out after %s s", OPEN_TO)
                try:
                    app.kill()
                except Exception:
                    pass
                raise FlowError(f"Excel failed to open workbook: {e}", False)
            pythoncom.PumpWaitingMessages()
            time.sleep(0.5)


def _run_macro_impl(wb_path: Path, args: tuple, log: logging.Logger):
    app, wb = _open_excel_with_timeout(wb_path, log)
    try:
        log.info("Running macro PopulateAndRunReport …")
        wb.macro("PopulateAndRunReport")(*args)
        wait_ready(wb, log)
        wb.api.Application.CalculateFull()
        wb.save()
        log.info("Workbook saved")
    finally:
        for op in (wb.close, app.kill):
            try:
                op()
            except Exception:
                pass


def run_excel_macro(wb_path: Path, args: tuple, log: logging.Logger):
    exc: list[Exception] = []

    def _worker():
        try:
            pythoncom.CoInitialize()
            _run_macro_impl(wb_path, args, log)
        except Exception as e:
            exc.append(e)
        finally:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    th = threading.Thread(target=_worker, daemon=True)
    th.start()

    start = time.time()
    while th.is_alive():
        time.sleep(0.5)
        if time.time() - start > READY_TO:
            log.error("Macro exceeded %s s – killing Excel", READY_TO)
            kill_orphan_excels()
            raise FlowError("Timeout running macro", work_completed=False)

    if exc:
        raise exc[0]


def read_cell(wb_path: Path, col: str, row: str) -> Any:
    if xw is None:
        raise FlowError("xlwings is required", work_completed=False)

    app = xw.App(visible=VISIBLE_EXCEL, add_book=False)  # type: ignore
    app.api.DisplayFullScreen = False
    try:
        wb = app.books.open(str(wb_path))
        value = wb.sheets[SCAC_VALIDATION_SHEET].range(f"{col}{row}").value
    finally:
        for op in (wb.close, app.kill):  # type: ignore
            try:
                op()
            except Exception:
                pass
    return value


__all__ = [
    "kill_orphan_excels",
    "copy_template",
    "wait_ready",
    "run_excel_macro",
    "read_cell",
]
