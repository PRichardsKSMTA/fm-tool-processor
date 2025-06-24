#!/usr/bin/env python3
"""
process_fm_tool.py
==================

Replicates—and extends—the legacy Power Automate Desktop (PAD) flow
for FM Tool (.xlsm) generation & SharePoint upload.

* Can run stand-alone:
      python process_fm_tool.py payload.json
* Or be imported by the Azure Function wrapper in fm_tool_processor/__init__.py

Author: 2025-06-24
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
from contextlib import contextmanager
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Tuple

import psutil                # process management
import xlwings as xw         # Excel COM automation
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential
from tenacity import (
    retry,
    retry_if_exception_type,
    stop_after_attempt,
    wait_fixed,
)

# --------------------------- CONFIGURATION --------------------------- #

# Environment variables (set in VM, Azure Functions App Settings, or .env)
SP_USERNAME     = os.getenv("SP_USERNAME")     # e.g. user@tenant.onmicrosoft.com
SP_PASSWORD     = os.getenv("SP_PASSWORD")     # app-password or AAD creds
SP_CHUNK_SIZE   = int(os.getenv("SP_CHUNK_MB" , "10")) * 1024 * 1024  # upload chunks
LOG_DIR         = Path(os.getenv("LOG_DIR"     , "./logs")).resolve()

# -------------------------------------------------------------------- #


class FlowError(Exception):
    """Domain-specific error that bubbles up to Cloud Flow."""
    def __init__(self, message: str, *, work_completed: bool = False) -> None:
        super().__init__(message)
        self.work_completed = work_completed
        self.message = message


@contextmanager
def excel_app():
    """Context manager: start hidden Excel instance, guarantee kill."""
    app = xw.App(visible=False, add_book=False)
    try:
        yield app
    finally:
        # Close all workbooks gracefully first
        for wb in list(app.books):
            try:
                wb.close()
            except Exception:  # noqa: BLE001
                pass
        app.kill()


def kill_orphan_excels() -> None:
    """Terminate all stray Excel processes before we begin."""
    for proc in psutil.process_iter(attrs=["name"]):
        if proc.info["name"] and proc.info["name"].lower().startswith("excel"):
            try:
                proc.kill()
            except Exception:  # noqa: BLE001
                pass


def copy_and_rename_template(template_path: str,
                             dest_folder: str,
                             new_name: str) -> Path:
    """Copy template to processing folder and rename/overwrite."""
    dest_folder = Path(dest_folder).expanduser().resolve()
    dest_folder.mkdir(parents=True, exist_ok=True)

    new_path = dest_folder / new_name
    shutil.copy2(template_path, new_path)  # overwrite OK
    return new_path


def run_excel_macro(wb_path: Path,
                    macro_name: str,
                    *macro_args) -> xw.Book:
    """Open workbook, run macro, save, return Book object (for validation)."""
    with excel_app() as app:
        wb = app.books.open(str(wb_path))
        time.sleep(2)
        try:
            wb.macro(macro_name)(*macro_args)
            time.sleep(2)
            wb.save()
            return wb
        finally:
            wb.close()


def read_cell(wb: xw.Book, col: str, row: str) -> Any:
    """Read a single cell by column letter and row number (as str)."""
    sheet = wb.sheets[0]   # assuming first sheet
    return sheet.range(f"{col}{row}").value


def get_sp_context(site_url: str) -> ClientContext:
    if not SP_USERNAME or not SP_PASSWORD:
        raise FlowError("SharePoint credentials not set in environment",
                        work_completed=False)
    creds = UserCredential(SP_USERNAME, SP_PASSWORD)
    return ClientContext(site_url).with_credentials(creds)


def sharepoint_file_exists(ctx: ClientContext, relative_url: str) -> bool:
    """Return True if file exists, False otherwise."""
    try:
        ctx.web.get_file_by_server_relative_url(relative_url).get().execute_query()
        return True
    except Exception:  # noqa: BLE001
        return False


def sharepoint_upload(ctx: ClientContext,
                      folder_url: str,
                      filename: str,
                      local_path: Path) -> None:
    """
    Upload file in chunks.  `folder_url` is server-relative,
    e.g. '/sites/KSMTA_HUMD/Client  Downloads/Pricing Tools'
    """
    target_folder = ctx.web.get_folder_by_server_relative_url(folder_url)
    with local_path.open("rb") as fd:
        size = os.fstat(fd.fileno()).st_size
        if size <= SP_CHUNK_SIZE:
            target_folder.upload_file(filename, fd.read()).execute_query()
        else:
            uploaded = target_folder.files.create_upload_session(
                filename, size).execute_query()
            start = 0
            while start < size:
                chunk = fd.read(SP_CHUNK_SIZE)
                uploaded.upload_chunk(chunk, start, len(chunk)).execute_query()
                start += len(chunk)


# ----------------------- RETRYABLE WORK UNIT ------------------------ #

@retry(
    stop=stop_after_attempt(lambda args: args[0]["max_retry"]),
    wait=wait_fixed(3),
    retry=retry_if_exception_type(FlowError),
    reraise=True,
)
def process_single_item(current_item: Dict[str, Any],
                        max_retry: int,
                        enable_sp_upload: bool,
                        processing_root: str,
                        logger: logging.Logger) -> None:

    template_path             = current_item["TOOL_TEMPLATE_FILEPATH"]
    new_excel_filename        = current_item["NEW_EXCEL_FILENAME"]
    scac_opp                  = current_item["SCAC_OPP"]
    week_ct                   = current_item["WEEK_CT"]
    processing_week           = current_item["PROCESSING_WEEK"]
    scac_col                  = current_item["SCAC_VALIDATION_COLUMN"]
    scac_row                  = current_item["SCAC_VALIDATION_ROW"]
    oa_col                    = current_item["ORDERAREAS_VALIDATION_COLUMN"]
    oa_row                    = current_item["ORDERAREAS_VALIDATION_ROW"]
    oa_expected               = current_item["ORDERAREAS_VALIDATION_VALUE"]

    # --- Stage 1: Copy & rename template
    path = copy_and_rename_template(template_path, processing_root,
                                    new_excel_filename)
    logger.info("Template copied to %s", path)

    # --- Stage 2: Run macro
    wb = run_excel_macro(
        path,
        "PopulateAndRunReport",
        scac_opp,
        week_ct,
        processing_week,
    )
    logger.info("Macro PopulateAndRunReport executed")

    # --- Stage 3: Validation
    client_operation = read_cell(wb, scac_col, scac_row)
    validate_oa      = read_cell(wb, oa_col  , oa_row)
    logger.info("Validation cells read: %s=%s, %s=%s",
                f"{scac_col}{scac_row}", client_operation,
                f"{oa_col}{oa_row}", validate_oa)

    if client_operation != scac_opp:
        raise FlowError(f"Excel validation failed – expected "
                        f"{scac_opp} in {scac_col}{scac_row}, "
                        f"got {client_operation}", work_completed=False)

    if validate_oa == oa_expected:
        raise FlowError(f"Excel validation failed – "
                        f"found '{oa_expected}' in {oa_col}{oa_row}.",
                        work_completed=False)

    # --- Stage 4: SharePoint upload (optional)
    if enable_sp_upload:
        sp_site      = current_item["CLIENT_DEST_SITE"]
        sp_folder    = current_item["CLIENT_DEST_FOLDER_PATH"]
        rel_folder   = Path(sp_folder).as_posix().lstrip("/")
        ctx          = get_sp_context(sp_site)

        server_rel_file = f"{rel_folder}/{new_excel_filename}"
        if sharepoint_file_exists(ctx, f"/sites/{server_rel_file}"):
            logger.warning("File already exists on SharePoint – skipping upload")
        else:
            sharepoint_upload(ctx, f"/sites/{rel_folder}",
                              new_excel_filename, path)
            logger.info("Uploaded %s to SharePoint", server_rel_file)

    # --- Stage 5: Cleanup
    path.unlink(missing_ok=True)
    logger.info("Local file %s deleted", path)

    # If no exception raised, success!
    return


# ------------------------ MAIN ENTRYPOINT --------------------------- #

def run_flow(payload: Dict[str, Any]) -> Dict[str, Any]:
    """Main routine called by stand-alone CLI *or* Azure Function."""
    # Map PAD input → Python vars
    max_retry                = int(payload.get("item/In_intMaxRetry", 3))
    processing_root          = payload["item/In_strDestinationProcessingFolder"]
    input_rows: List[Dict]   = payload["item/In_dtInputData"]
    enable_sp_upload         = payload.get("item/In_boolEnableSharePointUpload", True)

    run_id = uuid.uuid4().hex[:8]
    LOG_DIR.mkdir(exist_ok=True, parents=True)
    log_path = LOG_DIR / f"{datetime.utcnow():%Y-%m-%d}_{run_id}.log"

    # Configure logger
    logger = logging.getLogger("fm_tool")
    logger.setLevel(logging.INFO)
    logger.handlers.clear()
    stream_h = logging.StreamHandler(sys.stdout)
    file_h   = logging.FileHandler(log_path, encoding="utf-8")
    for h in (stream_h, file_h):
        h.setFormatter(logging.Formatter(
            "%(asctime)s | %(levelname)s | %(message)s",
            "%Y-%m-%dT%H:%M:%SZ"))
        logger.addHandler(h)

    logger.info("----- FM Tool run %s started -----", run_id)
    kill_orphan_excels()

    try:
        for row in input_rows:
            # Sidebar:  tenacity pulls arguments from the retry-decorated func
            process_single_item(row,
                                max_retry=max_retry,
                                enable_sp_upload=enable_sp_upload,
                                processing_root=processing_root,
                                logger=logger)
        logger.info("All items processed successfully")
        return {
            "Out_strWorkExceptionMessage": "",
            "Out_boolWorkcompleted": True,
        }

    except FlowError as e:
        # Already descriptive – bubble up to caller
        logger.exception("FlowError encountered")
        return {
            "Out_strWorkExceptionMessage": str(e),
            "Out_boolWorkcompleted": e.work_completed,
        }

    except Exception as e:  # noqa: BLE001
        logger.exception("Unexpected exception")
        return {
            "Out_strWorkExceptionMessage": f"Unexpected error: {e}",
            "Out_boolWorkcompleted": False,
        }

    finally:
        kill_orphan_excels()
        logger.info("Log saved to %s", log_path)
        logger.info("----- FM Tool run %s ended -----", run_id)


# -------------------------- CLI helper ------------------------------ #

def _cli() -> None:
    parser = argparse.ArgumentParser(description="Run FM Tool processor")
    parser.add_argument("json_file",
                        help="Path to Power Automate payload JSON file "
                             "or '-' to read from stdin")
    args = parser.parse_args()

    data = (sys.stdin.read() if args.json_file == "-" else
            Path(args.json_file).read_text(encoding="utf-8"))
    payload = json.loads(data)
    result  = run_flow(payload)
    print(json.dumps(result, indent=2))


if __name__ == "__main__":
    _cli()
