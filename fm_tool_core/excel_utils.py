# excel_utils.py

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

# psutil for killing orphans
try:
    import psutil  # type: ignore
except ImportError:

    class _PS:
        @staticmethod
        def process_iter(attrs=None):
            return []

    psutil = _PS()

# xlwings for COM
try:
    import xlwings as xw  # type: ignore
except ImportError:
    xw = None  # type: ignore

# win32com for message pumping & CoInitialize
try:
    from win32com.client import pythoncom  # type: ignore
except ImportError:

    class _PC:
        @staticmethod
        def PumpWaitingMessages():
            pass

        @staticmethod
        def CoInitialize():
            pass

        @staticmethod
        def CoUninitialize():
            pass

    pythoncom = _PC()

# COM error codes to retry on
RPC_E_CALL_FAILED = -2147023170
RPC_E_SERVER_UNAVAILABLE = -2147023174


def kill_orphan_excels():
    """Force-kill any lingering excel.exe processes."""
    for p in psutil.process_iter(attrs=["name"]):
        name = p.info.get("name", "")
        if name and name.lower().startswith("excel"):
            try:
                p.kill()
                logging.info(f"Killed excel.exe (pid={p.pid})")
            except Exception:
                logging.warning(f"Could not kill excel.exe (pid={p.pid})")


def copy_template(src: str, dst_root: str, new_name: str, log: logging.Logger) -> Path:
    src_path = Path(src)
    if not src_path.exists():
        raise FlowError(f"Template not found: {src}", work_completed=False)
    dst_root = Path(dst_root).expanduser().resolve()
    dst_root.mkdir(parents=True, exist_ok=True)
    dst_path = dst_root / new_name
    try:
        shutil.copy2(src_path, dst_path)
    except PermissionError:
        log.warning("Destination locked; unlinking and retrying copy")
        try:
            dst_path.unlink()
        except Exception:
            pass
        shutil.copy2(src_path, dst_path)
    return dst_path


def wait_ready(wb, log: logging.Logger):
    start = time.time()
    last_log = -1
    while True:
        try:
            ref = wb.names[READY_NAME].refers_to
            if ref.startswith("="):
                ref = ref[1:]
            flag = ref.strip('"')
        except Exception:
            flag = "<not-set>"
        elapsed = int(time.time() - start)
        if elapsed != last_log:
            log.info(f"Polling flag: {flag} (t={elapsed}s)")
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
    app = xw.App(visible=True, add_book=False)  # type: ignore
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
                log.error(f"Excel open timed out after {OPEN_TO}s")
                try:
                    app.kill()
                except Exception:
                    pass
                raise FlowError(
                    f"Excel failed to open workbook: {e}", work_completed=False
                )
            pythoncom.PumpWaitingMessages()
            time.sleep(0.5)


def safe_run_macro(wb, macro_name: str, args: tuple, log: logging.Logger):
    """
    Run the given VBA macro once; errors bubble up to be caught by the outer loop.
    """
    log.info(f"Running macro {macro_name} …")
    wb.macro(macro_name)(*args)


def _run_macro_impl(wb_path: Path, args: tuple, log: logging.Logger):
    app, wb = _open_excel_with_timeout(wb_path, log)
    try:
        safe_run_macro(wb, "PopulateAndRunReport", args, log)
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
    """
    Execute the macro with retries on RPC failures.
    """
    backoff = 5.0
    retries = 3
    for attempt in range(1, retries + 1):
        kill_orphan_excels()  # ensure a clean slate each attempt
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
                log.error(f"Macro exceeded {READY_TO}s – killing Excel")
                kill_orphan_excels()
                exc.append(FlowError("Timeout running macro", work_completed=False))
                break

        if not exc:
            return  # success on this attempt

        # an exception occurred
        err = exc[0]
        hr = getattr(err, "hresult", None)
        log.error(f"Attempt {attempt}/{retries} failed: {err}")
        if hr in (RPC_E_CALL_FAILED, RPC_E_SERVER_UNAVAILABLE) and attempt < retries:
            log.warning(
                f"Retrying macro after RPC error (hresult={hr}) in {backoff}s …"
            )
            time.sleep(backoff)
            continue

        # no more retries or fatal error
        raise err


def write_home_fields(
    wb_path: Path, process_guid: str | None, customer_name: str | None
) -> None:
    """Write basic HOME sheet fields to *wb_path*."""
    if xw is None:
        raise FlowError("xlwings is required", work_completed=False)
    pythoncom.CoInitialize()
    app = xw.App(visible=VISIBLE_EXCEL, add_book=False)  # type: ignore
    app.api.DisplayFullScreen = False
    wb = None
    try:
        wb = app.books.open(str(wb_path))
        ws = wb.sheets["HOME"]
        ws.range("BID").value = process_guid
        ws.range("D8").value = customer_name
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


def read_cell(wb_path: Path, col: str, row: str) -> Any:
    if xw is None:
        raise FlowError("xlwings is required", work_completed=False)
    pythoncom.CoInitialize()
    app = xw.App(visible=VISIBLE_EXCEL, add_book=False)  # type: ignore
    app.api.DisplayFullScreen = False
    try:
        wb = app.books.open(str(wb_path))
        return wb.sheets[SCAC_VALIDATION_SHEET].range(f"{col}{row}").value
    finally:
        for op in (wb.close, app.kill):
            try:
                op()
            except Exception:
                pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


__all__ = [
    "kill_orphan_excels",
    "copy_template",
    "wait_ready",
    "run_excel_macro",
    "write_home_fields",
    "read_cell",
]
