#!/usr/bin/env python3
"""
FM Tool processing entry points.
"""
from __future__ import annotations

import argparse
import json
import logging
import sys
import time
import uuid
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List
from urllib.parse import urlparse

import psutil

from .constants import LOG_DIR, RETRY_SLEEP
from .excel_utils import (
    copy_template,
    kill_orphan_excels,
    read_cell,
    run_excel_macro,
)
from .exceptions import FlowError
from .sharepoint_utils import (
    sp_ctx,
    sp_exists,
    sp_upload,
)

def wait_for_cpu(
    max_percent: float = 80.0,
    check_interval: float = 1.0,
    backoff: float = 5.0,
) -> None:
    while True:
        cpu = psutil.cpu_percent(interval=check_interval)
        if cpu < max_percent:
            logging.getLogger("fm_tool").info(f"CPU load at {cpu}% < {max_percent}%, proceeding")
            return
        logging.getLogger("fm_tool").warning(f"High CPU load ({cpu}%), waiting {backoff}s…")
        time.sleep(backoff)


def process_row(
    row: Dict[str, Any],
    upload: bool,
    root: str,
    run_id: str,
    log: logging.Logger,
):
    dst_name = f"{Path(row['NEW_EXCEL_FILENAME']).stem}_{run_id}.xlsm"
    dst_path = copy_template(row["TOOL_TEMPLATE_FILEPATH"], root, dst_name, log)
    log.info(f"Copying template from {row['TOOL_TEMPLATE_FILEPATH']}")

    log.info("Template copied to %s", dst_path)

    log.info("Waiting for CPU to drop")
    wait_for_cpu()

    kill_orphan_excels()

    run_excel_macro(
        dst_path,
        (row["SCAC_OPP"], row["WEEK_CT"], row["PROCESSING_WEEK"]),
        log,
    )

    log.info("Reading validation …")
    op_val = read_cell(dst_path, row["SCAC_VALIDATION_COLUMN"], row["SCAC_VALIDATION_ROW"])
    oa_val = read_cell(
        dst_path,
        row["ORDERAREAS_VALIDATION_COLUMN"],
        row["ORDERAREAS_VALIDATION_ROW"],
    )

    log.info("Validation: %s=%s, %s=%s",
             f"{row['SCAC_VALIDATION_COLUMN']}{row['SCAC_VALIDATION_ROW']}", op_val,
             f"{row['ORDERAREAS_VALIDATION_COLUMN']}{row['ORDERAREAS_VALIDATION_ROW']}", oa_val)

    if op_val != row["SCAC_OPP"]:
        raise FlowError(f"Validation failed – expected {row['SCAC_OPP']} got {op_val}", work_completed=False)
    if oa_val == row["ORDERAREAS_VALIDATION_VALUE"]:
        raise FlowError("Validation failed – ORDER/AREA unchanged", work_completed=False)

    if upload:
        ctx = sp_ctx(row["CLIENT_DEST_SITE"])
        site_path = urlparse(row["CLIENT_DEST_SITE"]).path
        folder_name = row["CLIENT_DEST_FOLDER_PATH"].lstrip("/")
        folder_rel = (Path(site_path) / folder_name).as_posix()
        rel_file = f"{folder_rel}/{row['NEW_EXCEL_FILENAME']}"

        log.info("Uploading to %s", rel_file)
        if sp_exists(ctx, rel_file):
            log.warning("SharePoint file exists – skipping upload")
        else:
            sp_upload(ctx, folder_rel, row["NEW_EXCEL_FILENAME"], dst_path)
            log.info("Uploaded %s", rel_file)

    dst_path.unlink(missing_ok=True)
    log.info("Local file deleted")


def run_flow(payload: Dict[str, Any]) -> Dict[str, Any]:
    if "parameters" in payload and isinstance(payload["parameters"], dict):
        payload = payload["parameters"]

    max_retry = int(payload.get("item/In_intMaxRetry", 3))
    root_folder = payload["item/In_strDestinationProcessingFolder"]
    rows: List[Dict[str, Any]] = payload["item/In_dtInputData"]
    enable_upload = payload.get("item/In_boolEnableSharePointUpload", True)

    run_id = uuid.uuid4().hex[:8]
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    log_file = LOG_DIR / f"{datetime.utcnow():%Y-%m-%d-%H-%M-%S}_{run_id}.log"

    # ----- KEY CHANGE HERE: send StreamHandler to stderr only ----- #
    log = logging.getLogger("fm_tool")
    log.setLevel(logging.INFO)
    log.handlers.clear()
    fmt = logging.Formatter("%(asctime)s | %(levelname)s | %(message)s", "%Y-%m-%dT%H:%M:%SZ")

    # All logging to stderr
    stderr_handler = logging.StreamHandler(stream=sys.stderr)
    stderr_handler.setFormatter(fmt)
    log.addHandler(stderr_handler)

    # Still write a file for your archive
    file_handler = logging.FileHandler(log_file, encoding="utf-8")
    file_handler.setFormatter(fmt)
    log.addHandler(file_handler)
    # -------------------------------------------------------------- #

    log.info("----- FM Tool run %s -----", run_id)
    kill_orphan_excels()

    try:
        for row in rows:
            kill_orphan_excels()
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
                                attempts, max_retry, type(e).__name__, e)
                    kill_orphan_excels()
                    time.sleep(RETRY_SLEEP)

        log.info("SUCCESS")
        return {
            "Out_strWorkExceptionMessage": "",
            "Out_boolWorkcompleted": True,
            "Out_strLogPath": str(log_file),
        }

    except FlowError as fe:
        log.exception("FlowError encountered")
        return {
            "Out_strWorkExceptionMessage": str(fe),
            "Out_boolWorkcompleted": fe.work_completed,
            "Out_strLogPath": str(log_file),
        }

    except Exception as ex:
        log.exception("Unexpected exception")
        return {
            "Out_strWorkExceptionMessage": f"Unexpected error: {ex}",
            "Out_boolWorkcompleted": False,
            "Out_strLogPath": str(log_file),
        }

    finally:
        kill_orphan_excels()
        log.info("Log saved to %s", log_file)
        log.info("----- FM Tool run %s ended -----", run_id)


def _cli() -> None:
    ap = argparse.ArgumentParser(description="Run FM Tool processor")
    ap.add_argument("json_file", help="Payload file or '-' for stdin")
    args = ap.parse_args()

    raw_json = (sys.stdin.read() if args.json_file == "-" 
                else Path(args.json_file).read_text())
    result = run_flow(json.loads(raw_json))
    print(json.dumps(result))


if __name__ == "__main__":
    _cli()
