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


# -------------------- ROW PROCESSOR -------------------------------------- #


def process_row(
    row: Dict[str, Any], upload: bool, root: str, run_id: str, log: logging.Logger
):
    dst_name = f"{Path(row['NEW_EXCEL_FILENAME']).stem}_{run_id}.xlsm"
    dst_path = copy_template(row["TOOL_TEMPLATE_FILEPATH"], root, dst_name, log)
    log.info("Template copied to %s", dst_path)

    run_excel_macro(
        dst_path,
        (row["SCAC_OPP"], row["WEEK_CT"], row["PROCESSING_WEEK"]),
        log,
    )

    log.info("Reading validation cells …")
    op_val = read_cell(
        dst_path, row["SCAC_VALIDATION_COLUMN"], row["SCAC_VALIDATION_ROW"]
    )
    oa_val = read_cell(
        dst_path, row["ORDERAREAS_VALIDATION_COLUMN"], row["ORDERAREAS_VALIDATION_ROW"]
    )

    log.info(
        "Validation: %s=%s, %s=%s",
        f"{row['SCAC_VALIDATION_COLUMN']}{row['SCAC_VALIDATION_ROW']}",
        op_val,
        f"{row['ORDERAREAS_VALIDATION_COLUMN']}{row['ORDERAREAS_VALIDATION_ROW']}",
        oa_val,
    )

    if op_val != row["SCAC_OPP"]:
        raise FlowError(
            f"Validation failed – expected {row['SCAC_OPP']} got {op_val}",
            work_completed=False,
        )
    if oa_val == row["ORDERAREAS_VALIDATION_VALUE"]:
        raise FlowError(
            "Validation failed – ORDER/AREA unchanged", work_completed=False
        )

    if upload:
        ctx = sp_ctx(row["CLIENT_DEST_SITE"])
        site_path = urlparse(row["CLIENT_DEST_SITE"]).path
        folder_name = row["CLIENT_DEST_FOLDER_PATH"].lstrip("/")
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
    """
    Processes a batch of rows, capturing individual results so that one failure
    does not stop the entire process. Returns a summary with per-row statuses.
    """
    if "parameters" in payload and isinstance(payload["parameters"], dict):
        payload = payload["parameters"]

    max_retry = int(payload.get("item/In_intMaxRetry", 3))
    root_folder = payload["item/In_strDestinationProcessingFolder"]
    rows: List[Dict[str, Any]] = payload["item/In_dtInputData"]
    enable_upload = payload.get("item/In_boolEnableSharePointUpload", True)

    run_id = uuid.uuid4().hex[:8]
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    log_file = LOG_DIR / f"{datetime.utcnow():%Y-%m-%d-%H-%M-%S}_{run_id}.log"

    log = logging.getLogger("fm_tool")
    log.setLevel(logging.INFO)
    log.handlers.clear()
    for h in (logging.StreamHandler(), logging.FileHandler(log_file, encoding="utf-8")):
        h.setFormatter(
            logging.Formatter(
                "%(asctime)s | %(levelname)s | %(message)s", "%Y-%m-%dT%H:%M:%SZ"
            )
        )
        log.addHandler(h)

    log.info("----- FM Tool run %s -----", run_id)
    kill_orphan_excels()

    results: List[Dict[str, Any]] = []

    for row in rows:
        identifier = row.get('NEW_EXCEL_FILENAME', 'unknown')
        row_result: Dict[str, Any] = {
            'file': identifier,
            'status': 'running',
            'error': ''
        }
        results.append(row_result)
        attempts = 0
        try:
            while True:
                attempts += 1
                try:
                    process_row(row, enable_upload, root_folder, run_id, log)
                    row_result['status'] = 'success'
                    break
                except Exception as e:
                    if attempts >= max_retry:
                        raise
                    log.warning(
                        "Retry %s/%s after %s: %s",
                        attempts,
                        max_retry,
                        e.__class__.__name__,
                        e,
                    )
                    kill_orphan_excels()
                    time.sleep(RETRY_SLEEP)
        except FlowError as fe:
            log.error("Row %s failed: %s", identifier, fe)
            row_result['status'] = 'failed'
            row_result['error'] = str(fe)
        except Exception as ex:
            log.error("Unexpected error on row %s: %s", identifier, ex)
            row_result['status'] = 'failed'
            row_result['error'] = f"Unexpected error: {ex}"

    # Determine overall completion: true if all succeeded
    work_completed = all(r['status'] == 'success' for r in results)
    # Build summary message
    if work_completed:
        summary_msg = ''
        log.info("All rows completed successfully.")
    else:
        summary_msg = 'One or more rows failed; check Out_dtRowResults for details'
        log.warning(summary_msg)

    # Final cleanup
    kill_orphan_excels()
    log.info("Log saved to %s", log_file)
    log.info("----- FM Tool run %s ended -----", run_id)

    return {
        'Out_strWorkExceptionMessage': summary_msg,
        'Out_boolWorkcompleted': work_completed,
        'Out_dtRowResults': results,
    }


# -------------------- CLI ------------------------------------------------ #

def _cli() -> None:
    ap = argparse.ArgumentParser(description="Run FM Tool processor")
    ap.add_argument("json_file", help="Payload file path or '-' for stdin")
    args = ap.parse_args()

    raw_json = (
        sys.stdin.read() if args.json_file == "-" else Path(args.json_file).read_text()
    )
    result = run_flow(json.loads(raw_json))
    print(json.dumps(result, indent=2))


if __name__ == "__main__":  # pragma: no cover - manual execution
    _cli()
