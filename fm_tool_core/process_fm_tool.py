#!/usr/bin/env python3
"""
process_fm_tool.py  –  2025-06-24

Stable Excel automation.
• Unique working filename per run.
• Step-level logging, open/macro time-outs.
• Proper FlowError signature.
"""

from __future__ import annotations

import argparse
import json
import logging
import os
import shutil
import sys
import time
import uuid
from datetime import datetime
from pathlib import Path
from types import SimpleNamespace
from typing import Any, Dict, List

import psutil
import xlwings as xw
from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext

try:  # pythoncom is Windows-only; provide noop fallback for tests
    from win32com.client import pythoncom  # type: ignore
except Exception:  # pragma: no cover - triggered on non-Windows
    pythoncom = SimpleNamespace(PumpWaitingMessages=lambda: None)

# ---------- CONSTANTS ----------------------------------------------------- #
READY_NAME, READY_OK, READY_ERR = "PY_READY_FLAG", "READY", "ERROR"
READY_TO, OPEN_TO, POLL, RETRY_SLEEP = 600, 60, 0.25, 2
MACRO_TO = READY_TO
LOG_DIR = Path(os.getenv("LOG_DIR", "./logs")).resolve()
SP_USERNAME, SP_PASSWORD = os.getenv("SP_USERNAME"), os.getenv("SP_PASSWORD")
SP_CHUNK = int(os.getenv("SP_CHUNK_MB", "10")) * 1024 * 1024
# ------------------------------------------------------------------------- #


class FlowError(Exception):
    """Controlled exception with PAD-style flag."""

    def __init__(self, msg: str, *, work_completed: bool = False) -> None:
        super().__init__(msg)
        self.work_completed = work_completed

    def __str__(self) -> str:
        return self.args[0]


# -------------------- HELPERS -------------------------------------------- #
def kill_orphan_excels():
    for p in psutil.process_iter(attrs=["name"]):
        if p.info["name"] and p.info["name"].lower().startswith("excel"):
            try:
                p.kill()
            except Exception:
                pass


def copy_template(src: str, root: str, name: str, lg: logging.Logger) -> Path:
    src_path = Path(src)
    if not src_path.exists():
        raise FlowError(f"Template not found: {src}", work_completed=False)

    root_path = Path(root).expanduser().resolve()
    root_path.mkdir(parents=True, exist_ok=True)
    dst = root_path / name

    try:
        shutil.copy2(src_path, dst)
    except PermissionError:
        lg.warning("Destination locked; unlinking and retrying copy")
        try:
            dst.unlink()
        except Exception:
            pass
        shutil.copy2(src_path, dst)

    return dst


def wait_ready(wb: xw.Book, lg: logging.Logger):
    start = time.time()
    while True:
        try:
            flag = wb.names[READY_NAME].refers_to.replace('"', "")
        except Exception:
            flag = "<not-set>"

        elapsed = int(time.time() - start)
        lg.info("Polling flag: %s (t=%ss)", flag, elapsed)

        if flag == READY_OK:
            return
        if flag == READY_ERR:
            raise FlowError("VBA signalled ERROR", work_completed=False)
        if elapsed >= READY_TO:
            raise FlowError(
                "Timeout waiting for READY flag",
                work_completed=False,
            )
        time.sleep(POLL)


def open_with_timeout(
    path: Path,
    lg: logging.Logger,
) -> tuple[xw.App, xw.Book]:
    lg.info("Creating Excel App …")
    app = xw.App(visible=False, add_book=False)
    app.api.Application.DisplayAlerts = False

    lg.info("Opening workbook …")
    t0 = time.time()
    while True:
        try:
            book = app.books.open(str(path))
            return app, book
        except Exception as e:
            if time.time() - t0 > OPEN_TO:
                lg.error("Excel open timed-out after %s s", OPEN_TO)
                try:
                    app.kill()
                except Exception:
                    pass
                raise FlowError(
                    f"Excel failed to open workbook: {e}", work_completed=False
                )
            pythoncom.PumpWaitingMessages()
            time.sleep(0.5)


def _run_macro_impl(dst: Path, args: tuple, lg: logging.Logger):
    app, wb = open_with_timeout(dst, lg)
    try:
        lg.info("Running macro PopulateAndRunReport …")
        wb.macro("PopulateAndRunReport")(*args)
        wait_ready(wb, lg)
        wb.save()
        lg.info("Workbook saved")
    finally:
        for op in (wb.close, app.kill):
            try:
                op()
            except Exception:
                pass


def run_macro(dst: Path, args: tuple, lg: logging.Logger):
    """Run macro in a thread so KeyboardInterrupt can interrupt."""
    exc: List[Exception] = []

    def _worker():
        try:
            pythoncom.CoInitialize()
            _run_macro_impl(dst, args, lg)
        except Exception as e:  # pragma: no cover - worker errors
            exc.append(e)
        finally:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    th = threading.Thread(target=_worker, daemon=True)
    th.start()
    th.join(MACRO_TO)
    if th.is_alive():
        lg.error("Macro timed-out after %s s", MACRO_TO)
        kill_orphan_excels()
        raise FlowError("Timeout running macro", work_completed=False)
    if exc:
        raise exc[0]


# Backwards compatibility with older code/tests
def run_excel_macro(dst: Path, args: tuple, lg: logging.Logger):
    """Alias maintained for legacy unit tests."""
    return run_macro(dst, args, lg)


def read_cell(path: Path, col: str, row: str):
    app = xw.App(visible=False, add_book=False)
    try:
        wb = app.books.open(str(path))
        val = wb.sheets[0].range(f"{col}{row}").value
    finally:
        for op in (wb.close, app.kill):
            try:
                op()
            except Exception:
                pass
    return val


def sp_ctx(site: str):
    if not (SP_USERNAME and SP_PASSWORD):
        raise FlowError("SharePoint credentials missing", work_completed=False)
    return ClientContext(site).with_credentials(
        UserCredential(SP_USERNAME, SP_PASSWORD)
    )


def sp_exists(ctx, rel):
    try:
        ctx.web.get_file_by_server_relative_url(rel).get().execute_query()
        return True
    except Exception:
        return False


# Backwards compatibility for unit tests
def sharepoint_file_exists(ctx, rel):
    """Alias for :func:`sp_exists`."""
    return sp_exists(ctx, rel)


def sp_upload(ctx, folder, fname, local: Path):
    tgt = ctx.web.get_folder_by_server_relative_url(folder)
    with local.open("rb") as f:
        size = os.fstat(f.fileno()).st_size
        if size <= SP_CHUNK:
            tgt.upload_file(fname, f.read()).execute_query()
        else:
            sess = tgt.files.create_upload_session(fname, size).execute_query()
            off = 0
            while off < size:
                chunk = f.read(SP_CHUNK)
                sess.upload_chunk(chunk, off, len(chunk)).execute_query()
                off += len(chunk)


def sharepoint_upload(ctx, folder, fname, local: Path):
    """Alias for :func:`sp_upload`."""
    return sp_upload(ctx, folder, fname, local)


# -------------------- WORK UNIT ------------------------------------------ #
def process_row(
    it: Dict[str, Any],
    upload: bool,
    root: str,
    run_id: str,
    lg: logging.Logger,
):
    unique_name = f"{Path(it['NEW_EXCEL_FILENAME']).stem}_{run_id}.xlsm"
    dst = copy_template(it["TOOL_TEMPLATE_FILEPATH"], root, unique_name, lg)
    lg.info("Template copied to %s", dst)

    run_excel_macro(
        dst,
        (
            it["SCAC_OPP"],
            it["WEEK_CT"],
            it["PROCESSING_WEEK"],
        ),
        lg,
    )

    lg.info("Reading validation cells …")
    op = read_cell(
        dst,
        it["SCAC_VALIDATION_COLUMN"],
        it["SCAC_VALIDATION_ROW"],
    )
    oa = read_cell(
        dst,
        it["ORDERAREAS_VALIDATION_COLUMN"],
        it["ORDERAREAS_VALIDATION_ROW"],
    )
    cell_ref = "".join(
        [
            f"{it['ORDERAREAS_VALIDATION_COLUMN']}",
            f"{it['ORDERAREAS_VALIDATION_ROW']}",
        ]
    )
    lg.info(
        "Validation: %s=%s, %s=%s",
        f"{it['SCAC_VALIDATION_COLUMN']}{it['SCAC_VALIDATION_ROW']}",
        op,
        cell_ref,
        oa,
    )
    if op != it["SCAC_OPP"]:
        raise FlowError(
            f"Validation failed – expected {it['SCAC_OPP']} got {op}",
            work_completed=False,
        )
    if oa == it["ORDERAREAS_VALIDATION_VALUE"]:
        raise FlowError("Validation failed – ORDER/AREA unchanged", False)

    if upload:
        ctx = sp_ctx(it["CLIENT_DEST_SITE"])
        rel_folder = Path(it["CLIENT_DEST_FOLDER_PATH"]).as_posix().lstrip("/")
        rel_file = f"/sites/{rel_folder}/{it['NEW_EXCEL_FILENAME']}"
        if sp_exists(ctx, rel_file):
            lg.warning("File exists on SharePoint – skip upload")
        else:
            sp_upload(
                ctx,
                f"/sites/{rel_folder}",
                it["NEW_EXCEL_FILENAME"],
                dst,
            )
            lg.info("Uploaded %s", rel_file)

    dst.unlink(missing_ok=True)
    lg.info("Local file deleted")


# -------------------- ORCHESTRATOR --------------------------------------- #
def run_flow(payload: Dict[str, Any]) -> Dict[str, Any]:
    if "parameters" in payload and isinstance(payload["parameters"], dict):
        payload = payload["parameters"]

    max_try = int(payload.get("item/In_intMaxRetry", 3))
    root = payload["item/In_strDestinationProcessingFolder"]
    rows: List[Dict] = payload["item/In_dtInputData"]
    upload = payload.get("item/In_boolEnableSharePointUpload", True)

    run_id = uuid.uuid4().hex[:8]
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    log_file = LOG_DIR / f"{datetime.utcnow():%Y-%m-%d-%H-%M-%S}_{run_id}.log"

    lg = logging.getLogger("fm_tool")
    lg.setLevel(logging.INFO)
    lg.handlers.clear()
    for h in (
        logging.StreamHandler(),
        logging.FileHandler(log_file, encoding="utf-8"),
    ):
        h.setFormatter(
            logging.Formatter(
                "%(asctime)s | %(levelname)s | %(message)s",
                "%Y-%m-%dT%H:%M:%SZ",
            )
        )
        lg.addHandler(h)

    lg.info("----- FM Tool run %s -----", run_id)
    kill_orphan_excels()

    try:
        for row in rows:
            attempt = 0
            while True:
                attempt += 1
                try:
                    process_row(row, upload, root, run_id, lg)
                    break
                except Exception as ex:
                    if attempt >= max_try:
                        raise
                    lg.warning(
                        "Retry %s/%s after %s: %s",
                        attempt,
                        max_try,
                        ex.__class__.__name__,
                        ex,
                    )
                    kill_orphan_excels()
                    time.sleep(RETRY_SLEEP)

        lg.info("SUCCESS")
        return {
            "Out_strWorkExceptionMessage": "",
            "Out_boolWorkcompleted": True,
        }

    except KeyboardInterrupt:
        lg.error("Interrupted by user")
        return {
            "Out_strWorkExceptionMessage": "Interrupted by user",
            "Out_boolWorkcompleted": False,
        }

    except FlowError as fe:
        lg.exception("FlowError")
        return {
            "Out_strWorkExceptionMessage": str(fe),
            "Out_boolWorkcompleted": fe.work_completed,
        }

    except Exception as ex:
        lg.exception("Unexpected")
        return {
            "Out_strWorkExceptionMessage": f"Unexpected error: {ex}",
            "Out_boolWorkcompleted": False,
        }

    finally:
        kill_orphan_excels()
        lg.info("Log saved to %s", log_file)


# -------------------- CLI ------------------------------------------------- #
def _cli():
    ap = argparse.ArgumentParser()
    ap.add_argument("json_file")
    p = ap.parse_args()
    if p.json_file == "-":
        raw = sys.stdin.read()
    else:
        raw = Path(p.json_file).read_text()
    print(json.dumps(run_flow(json.loads(raw)), indent=2))


if __name__ == "__main__":
    _cli()
