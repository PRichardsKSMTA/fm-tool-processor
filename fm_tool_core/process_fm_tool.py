#!/usr/bin/env python3
"""
process_fm_tool.py   –   2025-06-25

Stable Excel automation for FM-Tool processing
──────────────────────────────────────────────
• Unique working filename per run
• Step-level logging, open/macro time-outs
• Ctrl-C responsive
• READY-flag normalization (=READY → READY) so polling exits correctly
"""

from __future__ import annotations

import argparse
import json
import logging
import os
import shutil
import sys
import threading
import time
import uuid
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List
from urllib.parse import urlparse

import psutil
import xlwings as xw
from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext

# win32com is Windows-only; provide no-op fallback for tests / non-Windows
try:
    from win32com.client import pythoncom  # type: ignore
except Exception:  # pragma: no cover
    class _PC:
        @staticmethod
        def PumpWaitingMessages(): ...
        @staticmethod
        def CoInitialize(): ...
        @staticmethod
        def CoUninitialize(): ...
    pythoncom = _PC()

# ---------- CONSTANTS ---------------------------------------------------- #
READY_NAME, READY_OK, READY_ERR = "PY_READY_FLAG", "READY", "ERROR"
READY_TO, OPEN_TO, POLL_SLEEP, RETRY_SLEEP = 600, 60, 0.25, 2  # seconds
LOG_DIR = Path(os.getenv("LOG_DIR", "./logs")).resolve()
SP_USERNAME, SP_PASSWORD = os.getenv("SP_USERNAME"), os.getenv("SP_PASSWORD")
# The site collection under which our service account lives
ROOT_SP_SITE = "https://ksmcpa.sharepoint.com/teams/ksmta"
# Sheet that holds the validation cells
SCAC_VALIDATION_SHEET = "HOME"
# Show the Excel UI when FM_SHOW_EXCEL=1
VISIBLE_EXCEL = os.getenv("FM_SHOW_EXCEL", "0") == "1"
# ------------------------------------------------------------------------ #


class FlowError(Exception):
    """Domain-specific exception returned to Power Automate Cloud Flow."""

    def __init__(self, msg: str, *, work_completed: bool = False) -> None:
        super().__init__(msg)
        self.work_completed = work_completed

    def __str__(self) -> str:
        return self.args[0]


# -------------------- PROCESS HELPERS ----------------------------------- #
def kill_orphan_excels():
    for p in psutil.process_iter(attrs=["name"]):
        if p.info["name"] and p.info["name"].lower().startswith("excel"):
            try:
                p.kill()
            except Exception:
                pass


def copy_template(src: str, dst_root: str, new_name: str,
                  log: logging.Logger) -> Path:
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


def wait_ready(wb: xw.Book, log: logging.Logger):
    start = time.time()
    last_log = -999
    while True:
        try:
            ref = wb.names[READY_NAME].refers_to
            if ref.startswith("="):  # e.g. ="READY" or =READY
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


def _open_excel_with_timeout(path: Path, log: logging.Logger
                             ) -> tuple[xw.App, xw.Book]:
    log.info("Creating Excel App …")
    app = xw.App(visible=VISIBLE_EXCEL, add_book=False)
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


def run_macro(wb_path: Path, args: tuple, log: logging.Logger):
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
    app = xw.App(visible=VISIBLE_EXCEL, add_book=False)
    app.api.DisplayFullScreen = False
    try:
        wb = app.books.open(str(wb_path))
        value = wb.sheets[SCAC_VALIDATION_SHEET].range(f"{col}{row}").value
    finally:
        for op in (wb.close, app.kill):
            try:
                op()
            except Exception:
                pass
    return value


def sp_ctx(_: str = None):
    if not (SP_USERNAME and SP_PASSWORD):
        raise FlowError("SharePoint credentials missing", work_completed=False)
    return ClientContext(ROOT_SP_SITE).with_credentials(
        UserCredential(SP_USERNAME, SP_PASSWORD)
    )


def sp_exists(ctx, rel_url: str) -> bool:
    try:
        ctx.web.get_file_by_server_relative_url(rel_url).get().execute_query()
        return True
    except Exception:
        return False


def sp_upload(ctx, folder: str, fname: str, local: Path):
    tgt = ctx.web.get_folder_by_server_relative_url(folder)
    with local.open("rb") as f:
        content = f.read()
    tgt.upload_file(fname, content).execute_query()


# -------------------- ROW PROCESSOR -------------------------------------- #
def process_row(row: Dict[str, Any], upload: bool, root: str,
                run_id: str, log: logging.Logger):
    dst_name = f"{Path(row['NEW_EXCEL_FILENAME']).stem}_{run_id}.xlsm"
    dst_path = copy_template(row["TOOL_TEMPLATE_FILEPATH"], root, dst_name, log)
    log.info("Template copied to %s", dst_path)

    run_macro(
        dst_path,
        (row["SCAC_OPP"], row["WEEK_CT"], row["PROCESSING_WEEK"]),
        log,
    )

    log.info("Reading validation cells …")
    op_val = read_cell(dst_path, row["SCAC_VALIDATION_COLUMN"], row["SCAC_VALIDATION_ROW"])
    oa_val = read_cell(dst_path, row["ORDERAREAS_VALIDATION_COLUMN"], row["ORDERAREAS_VALIDATION_ROW"])

    log.info("Validation: %s=%s, %s=%s",
             f"{row['SCAC_VALIDATION_COLUMN']}{row['SCAC_VALIDATION_ROW']}", op_val,
             f"{row['ORDERAREAS_VALIDATION_COLUMN']}{row['ORDERAREAS_VALIDATION_ROW']}", oa_val)

    if op_val != row["SCAC_OPP"]:
        raise FlowError(f"Validation failed – expected {row['SCAC_OPP']} got {op_val}", work_completed=False)
    if oa_val == row["ORDERAREAS_VALIDATION_VALUE"]:
        raise FlowError("Validation failed – ORDER/AREA unchanged", work_completed=False)

    if upload:
        ctx = sp_ctx()
        site_path = urlparse(row["CLIENT_DEST_SITE"]).path
        folder_name = row["CLIENT_DEST_FOLDER_PATH"].lstrip("/")
        # build a forward-slash server-relative URL
        folder_rel = (Path(site_path) / folder_name).as_posix()
        rel_file = f"{folder_rel}/{row['NEW_EXCEL_FILENAME']}"

        log.info("DEBUG: Uploading to server-relative folder %s", folder_rel)
        if sp_exists(ctx, rel_file):
            log.warning("File exists on SharePoint – skip upload")
        else:
            sp_upload(ctx, folder_rel, row['NEW_EXCEL_FILENAME'], dst_path)
            log.info("Uploaded %s", rel_file)

    dst_path.unlink(missing_ok=True)
    log.info("Local file deleted")


# -------------------- MAIN RUNNER ---------------------------------------- #
def run_flow(payload: Dict[str, Any]) -> Dict[str, Any]:
    if "parameters" in payload and isinstance(payload["parameters"], dict):
        payload = payload["parameters"]

    max_retry = int(payload.get("item/In_intMaxRetry", 3))
    root_folder = payload["item/In_strDestinationProcessingFolder"]
    rows: List[Dict] = payload["item/In_dtInputData"]
    enable_upload = payload.get("item/In_boolEnableSharePointUpload", True)

    run_id = uuid.uuid4().hex[:8]
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    log_file = LOG_DIR / f"{datetime.utcnow():%Y-%m-%d-%H-%M-%S}_{run_id}.log"

    log = logging.getLogger("fm_tool")
    log.setLevel(logging.INFO)  
    log.handlers.clear()
    for h in (logging.StreamHandler(),
              logging.FileHandler(log_file, encoding="utf-8")):
        h.setFormatter(logging.Formatter(
            "%(asctime)s | %(levelname)s | %(message)s", "%Y-%m-%dT%H:%M:%SZ"))
        log.addHandler(h)

    log.info("----- FM Tool run %s -----", run_id)
    kill_orphan_excels()

    try:
        for row in rows:
            attempts = 0
            while True:
                attempts += 1
                try:
                    process_row(row, enable_upload, root_folder, run_id, log)
                    break
                except Exception as e:
                    if attempts >= max_retry:
                        raise
                    log.warning("Retry %s/%s after %s: %s", 
                                attempts, max_retry, e.__class__.__name__, e)
                    kill_orphan_excels()
                    time.sleep(RETRY_SLEEP)

        log.info("SUCCESS")
        return {"Out_strWorkExceptionMessage": "", "Out_boolWorkcompleted": True}

    except KeyboardInterrupt:
        log.error("Interrupted by user")
        return {"Out_strWorkExceptionMessage": "Interrupted by user",
                "Out_boolWorkcompleted": False}

    except FlowError as fe:
        log.exception("FlowError encountered")
        return {"Out_strWorkExceptionMessage": str(fe), "Out_boolWorkcompleted": fe.work_completed}

    except Exception as ex:
        log.exception("Unexpected exception")
        return {"Out_strWorkExceptionMessage": f"Unexpected error: {ex}",
                "Out_boolWorkcompleted": False}

    finally:
        kill_orphan_excels()
        log.info("Log saved to %s", log_file)
        log.info("----- FM Tool run %s ended -----", run_id)


# -------------------- CLI ------------------------------------------------ #
def _cli() -> None:
    ap = argparse.ArgumentParser(description="Run FM Tool processor")
    ap.add_argument("json_file", help="Payload file path or '-' for stdin")
    args = ap.parse_args()

    raw_json = (sys.stdin.read() if args.json_file == "-"
                else Path(args.json_file).read_text())
    result = run_flow(json.loads(raw_json))
    print(json.dumps(result, indent=2))


if __name__ == "__main__":
    _cli()
