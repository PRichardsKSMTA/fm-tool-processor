# excel_utils.py

# Allow postponing evaluation of type hints
from __future__ import annotations

# Log messages to the console or a file
import logging

# Copy files while preserving metadata
import shutil

# Run code in separate threads
import threading

# Provide timing functions such as sleep
import time

# Work with filesystem paths in an OS-agnostic way
from pathlib import Path

# Provide the Any type for flexible annotations
from typing import Any

# Import various timing and configuration constants from this package
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

# Import the custom exception type used by the flow
from .exceptions import FlowError

# Try to import psutil so we can kill stray Excel processes
try:
    import psutil
except ImportError:
    # Fallback stub when psutil is not installed
    class _PS:
        @staticmethod
        def process_iter(attrs=None):
            return []

    psutil = _PS()

# Import xlwings to drive Excel via COM automation
try:
    import xlwings as xw
except ImportError:
    # Optional dependency; keep None if missing
    xw = None  # type: ignore

# Import win32com pieces used for COM initialisation and message pumping
try:
    from win32com.client import pythoncom
except ImportError:
    # Provide a stub with the same API if win32com is missing
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

# Windows COM error codes that warrant a retry
RPC_E_CALL_FAILED = -2147023170
RPC_E_SERVER_UNAVAILABLE = -2147023174


def kill_orphan_excels():
    """Force-kill any lingering excel.exe processes."""
    # Look through all running processes
    for p in psutil.process_iter(attrs=["name"]):
        # Get the program name for each process
        name = p.info.get("name", "")
        # If it's Excel, attempt to terminate it
        if name and name.lower().startswith("excel"):
            try:
                p.kill()
                logging.info(f"Killed excel.exe (pid={p.pid})")
            except Exception:
                logging.warning(f"Could not kill excel.exe (pid={p.pid})")


def copy_template(src: str, dst_root: str, new_name: str, log: logging.Logger) -> Path:
    # Convert the source path to a Path object
    src_path = Path(src)
    # Fail if the template does not exist
    if not src_path.exists():
        raise FlowError(f"Template not found: {src}", work_completed=False)
    # Resolve the destination directory and create it if needed
    dst_root = Path(dst_root).expanduser().resolve()
    dst_root.mkdir(parents=True, exist_ok=True)
    # Build the destination file path
    dst_path = dst_root / new_name
    try:
        # Copy the template file to the new location
        shutil.copy2(src_path, dst_path)
    except PermissionError:
        # If the destination is locked, remove it and try again
        log.warning("Destination locked; unlinking and retrying copy")
        try:
            dst_path.unlink()
        except Exception:
            pass
        shutil.copy2(src_path, dst_path)
    # Return the path to the copied file
    return dst_path


def wait_ready(wb, log: logging.Logger):
    # Remember when we started waiting
    start = time.time()
    last_log = -1
    # Poll the workbook until the ready flag is set
    while True:
        try:
            # Read the named range used as the ready flag
            ref = wb.names[READY_NAME].refers_to
            if ref.startswith("="):
                ref = ref[1:]
            flag = ref.strip('"')
        except Exception:
            flag = "<not-set>"
        # Log the current flag value once per second
        elapsed = int(time.time() - start)
        if elapsed != last_log:
            log.info(f"Polling flag: {flag} (t={elapsed}s)")
            last_log = elapsed
        # Check for success or error conditions
        if flag == READY_OK:
            return
        if flag == READY_ERR:
            raise FlowError("VBA signaled ERROR", work_completed=False)
        if elapsed >= READY_TO:
            raise FlowError("Timeout waiting for READY flag", work_completed=False)
        # Sleep briefly before polling again
        time.sleep(POLL_SLEEP)


def _open_excel_with_timeout(path: Path, log: logging.Logger):
    # Ensure xlwings is available
    if xw is None:
        raise FlowError("xlwings is required", work_completed=False)
    # Start a new Excel application
    log.info("Creating Excel App …")
    app = xw.App(visible=True, add_book=False)  # type: ignore
    app.api.DisplayFullScreen = False
    app.api.Application.DisplayAlerts = False
    log.info("Opening workbook …")
    t0 = time.time()
    # Keep trying to open the workbook until success or timeout
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
            # Allow COM messages to be processed
            pythoncom.PumpWaitingMessages()
            time.sleep(0.5)


def safe_run_macro(wb, macro_name: str, args: tuple, log: logging.Logger):
    """Run the given VBA macro once."""
    # Inform the log which macro is running
    log.info(f"Running macro {macro_name} …")
    # Invoke the macro with the provided arguments
    wb.macro(macro_name)(*args)


def _run_macro_impl(wb_path: Path, args: tuple, log: logging.Logger):
    # Open Excel and the workbook
    app, wb = _open_excel_with_timeout(wb_path, log)
    try:
        # Call the macro that does all the work
        safe_run_macro(wb, "PopulateAndRunReport", args, log)
        # Wait until the VBA code signals completion
        wait_ready(wb, log)
        # Recalculate all formulas
        wb.api.Application.CalculateFull()
        # Save the workbook
        wb.save()
        log.info("Workbook saved")
    finally:
        # Always close the workbook and kill Excel
        for op in (wb.close, app.kill):
            try:
                op()
            except Exception:
                pass


def run_excel_macro(wb_path: Path, args: tuple, log: logging.Logger):
    """Execute the macro with retries on RPC failures."""
    backoff = 5.0
    retries = 3
    for attempt in range(1, retries + 1):
        # Remove any existing Excel processes before starting
        kill_orphan_excels()  # ensure a clean slate each attempt
        exc: list[Exception] = []

        def _worker():
            # Runs in a background thread to execute the macro
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

        # An exception occurred while running the macro
        err = exc[0]
        hr = getattr(err, "hresult", None)
        log.error(f"Attempt {attempt}/{retries} failed: {err}")
        if hr in (RPC_E_CALL_FAILED, RPC_E_SERVER_UNAVAILABLE) and attempt < retries:
            log.warning(
                f"Retrying macro after RPC error (hresult={hr}) in {backoff}s …"
            )
            time.sleep(backoff)
            continue

        # No more retries or fatal error
        raise err


def read_cell(wb_path: Path, col: str, row: str) -> Any:
    # Read a single cell value from a workbook
    if xw is None:
        raise FlowError("xlwings is required", work_completed=False)
    app = xw.App(visible=VISIBLE_EXCEL, add_book=False)  # type: ignore
    app.api.DisplayFullScreen = False
    try:
        # Open the workbook and return the requested cell value
        wb = app.books.open(str(wb_path))
        return wb.sheets[SCAC_VALIDATION_SHEET].range(f"{col}{row}").value
    finally:
        # Close the workbook and quit Excel
        for op in (wb.close, app.kill):
            try:
                op()
            except Exception:
                pass


# Symbols exported when importing * from this module
__all__ = [
    "kill_orphan_excels",
    "copy_template",
    "wait_ready",
    "run_excel_macro",
    "read_cell",
]
