#!/usr/bin/env python3
"""
FM Tool processing entry point – verbose logging & SQL status tracking
=====================================================================

* Detailed INFO-level log lines (template copy, CPU wait %, validation, upload).
* BEGIN / COMPLETE / RESET stored-procedure calls via env-vars
  (SQL_SERVER, SQL_DATABASE, SQL_USERNAME, SQL_PASSWORD).
  Override proc names with SQL_UPDATE_PROC / SQL_RESET_PROC.
* Duplicate SharePoint uploads are INFO-level only.
* `pyodbc` optional – if missing, SQL is skipped but run continues.
"""
from __future__ import annotations

import argparse
import json
import logging
import os
import sys
import time
import uuid
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List
from urllib.parse import urlparse

import psutil  # type: ignore  # noqa: E401

# ────────────────────────────── pyodbc (optional) ──────────────────────────
try:
    import pyodbc  # type: ignore
except ImportError as _e:  # pragma: no cover
    pyodbc = None  # type: ignore
    logging.basicConfig(level=logging.WARNING)
    logging.warning("pyodbc missing – SQL disabled (%s)", _e)

from .constants import LOG_DIR, RETRY_SLEEP
from .excel_utils import copy_template, kill_orphan_excels, read_cell, run_excel_macro
from .exceptions import FlowError
from .sharepoint_utils import sp_ctx, sp_exists, sp_upload

# ───────────────────────────── SQL HELPERS ─────────────────────────────────
_SQL_CONN_STR: str | None = None


def _sql_conn_str() -> str:
    """Build (and cache) an ODBC connection string."""
    global _SQL_CONN_STR
    if _SQL_CONN_STR:
        return _SQL_CONN_STR
    srv, db, usr, pwd = (os.getenv(k) for k in ("SQL_SERVER",
                                                "SQL_DATABASE",
                                                "SQL_USERNAME",
                                                "SQL_PASSWORD"))
    if not all((srv, db, usr, pwd)):
        raise RuntimeError("SQL connection env vars missing")
    _SQL_CONN_STR = (
        f"DRIVER={{ODBC Driver 18 for SQL Server}};SERVER={srv};DATABASE={db};"
        f"UID={usr};PWD={pwd};Encrypt=yes;TrustServerCertificate=yes;"
    )
    return _SQL_CONN_STR


def _exec_proc(proc: str,
               params: tuple[Any, ...],
               log: logging.Logger) -> None:
    """Execute a stored procedure (no-op if pyodbc is unavailable)."""
    if pyodbc is None:
        log.info("(SQL disabled) would EXEC %s %s", proc, params)
        return
    conn = None
    try:
        conn = pyodbc.connect(_sql_conn_str(), timeout=10)
        with conn.cursor() as cur:
            log.info("EXEC %s %s", proc, ", ".join(repr(p) for p in params))
            cur.execute(f"EXEC {proc} " + ", ".join("?" for _ in params), params)
            conn.commit()
    except Exception as exc:
        log.warning("Stored procedure %s failed: %s", proc, exc)
    finally:
        if conn:
            try:
                conn.close()
            except Exception:
                pass


def _update_status(scac: str, state: str,
                   log: logging.Logger) -> None:
    _exec_proc(os.getenv("SQL_UPDATE_PROC",
                         "dbo.UPDATE_CLIENT_UPLOAD_STATUS"),
               (scac, state), log)


def _reset_status(scac: str, state: str,
                  log: logging.Logger) -> None:
    _exec_proc(os.getenv("SQL_RESET_PROC",
                         "dbo.RESET_CLIENT_PROCESSING_STATUS"),
               (scac, state), log)

# ───────────────────────────── HELPER FUNCTIONS ────────────────────────────
def _fifo_sort(rows: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Keep FIFO order based on first timestamp-like key, else input order."""
    key = next((k for k in ("QUEUE_TS", "CREATED_UTC", "CREATED_AT")
                if all(k in r for r in rows)), None)
    if not key:
        return rows

    def _ts(v: Any) -> float:
        if isinstance(v, datetime):
            return v.timestamp()
        if isinstance(v, (int, float)):
            return float(v)
        if isinstance(v, str):
            try:
                return datetime.fromisoformat(v.rstrip("Z")).timestamp()
            except ValueError:
                pass
        return time.time()

    return sorted(rows, key=lambda r: _ts(r[key]))


def wait_for_cpu(max_percent: float = 80.0,
                 interval: float = 1.0,
                 backoff: float = 5.0,
                 log: logging.Logger | None = None) -> None:
    """Block until overall CPU utilisation drops below *max_percent*."""
    while True:
        cpu = psutil.cpu_percent(interval)
        if cpu < max_percent:
            if log:
                log.info("CPU load at %.1f%% < %.1f%%, proceeding",
                         cpu, max_percent)
            return
        if log:
            log.warning("High CPU (%.1f%%) – sleeping %ss",
                        cpu, backoff)
        time.sleep(backoff)

# ───────────────────────────── ROW WORKER ──────────────────────────────────
def process_row(row: Dict[str, Any],
                upload: bool,
                root: str,
                run_id: str,
                log: logging.Logger) -> bool:
    """Process one FM payload row. Returns True on success."""
    op_code = row["SCAC_OPP"]
    template_src = row["TOOL_TEMPLATE_FILEPATH"]

    log.info("Copying template from %s", template_src)
    dst_path = copy_template(template_src, root,
                             f"{Path(row['NEW_EXCEL_FILENAME']).stem}_{run_id}.xlsm", log)
    log.info("Template copied to %s", dst_path)

    log.info("Waiting for CPU to drop")
    wait_for_cpu(log=log)
    kill_orphan_excels()

    log.info("Creating Excel App …")
    try:
        log.info("Opening workbook …")
        run_excel_macro(dst_path,
                        (row["SCAC_OPP"], row["WEEK_CT"], row["PROCESSING_WEEK"]),
                        log)
        log.info("Running macro PopulateAndRunReport …")

        log.info("Reading validation …")
        op_val = read_cell(dst_path,
                           row["SCAC_VALIDATION_COLUMN"],
                           row["SCAC_VALIDATION_ROW"])
        oa_val = read_cell(dst_path,
                           row["ORDERAREAS_VALIDATION_COLUMN"],
                           row["ORDERAREAS_VALIDATION_ROW"])
        log.info("Validation: %s=%s, %s=%s",
                 f"{row['SCAC_VALIDATION_COLUMN']}{row['SCAC_VALIDATION_ROW']}",
                 op_val,
                 f"{row['ORDERAREAS_VALIDATION_COLUMN']}{row['ORDERAREAS_VALIDATION_ROW']}",
                 oa_val)

        if op_val != op_code or \
           oa_val == row["ORDERAREAS_VALIDATION_VALUE"]:
            raise FlowError("Validation failed", work_completed=False)

        if upload:
            ctx = sp_ctx(row["CLIENT_DEST_SITE"])
            site_path = urlparse(row["CLIENT_DEST_SITE"]).path
            folder = (Path(site_path) /
                      row["CLIENT_DEST_FOLDER_PATH"].lstrip("/")).as_posix()
            rel_file = f"{folder}/{row['NEW_EXCEL_FILENAME']}"
            log.info("Uploading to %s", rel_file)
            if sp_exists(ctx, rel_file):
                log.info("SharePoint file exists – skipping upload "
                         "(not an error)")
            else:
                sp_upload(ctx, folder,
                          row["NEW_EXCEL_FILENAME"], dst_path)
                log.info("Uploaded %s", rel_file)

        log.info("Local file deleted")
        return True
    except Exception:
        log.exception("process_row failure")
        return False
    finally:
        dst_path.unlink(missing_ok=True)
        kill_orphan_excels()

# ───────────────────────────── RUN FLOW ────────────────────────────────────
def run_flow(payload: Dict[str, Any]) -> Dict[str, Any]:
    """Main entry when invoked by the PowerShell wrapper."""
    if "parameters" in payload and isinstance(payload["parameters"], dict):
        payload = payload["parameters"]

    rows = _fifo_sort(payload["item/In_dtInputData"])
    root_folder = payload["item/In_strDestinationProcessingFolder"]
    enable_upload = payload.get("item/In_boolEnableSharePointUpload", True)
    max_retry = int(payload.get("item/In_intMaxRetry", 1))

    run_id = uuid.uuid4().hex[:8]
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    log_file = LOG_DIR / f"{datetime.utcnow():%Y-%m-%d-%H-%M-%S}_{run_id}.log"

    log = logging.getLogger("fm_tool"); log.handlers.clear()
    log.setLevel(logging.INFO)
    fmt = logging.Formatter("%(asctime)s | %(levelname)s | %(message)s",
                            "%Y-%m-%dT%H:%M:%SZ")
    for h in (logging.StreamHandler(sys.stderr),
              logging.FileHandler(log_file, encoding="utf-8")):
        h.setFormatter(fmt)
        log.addHandler(h)

    log.info("----- FM Tool run %s -----", run_id)

    op_code = rows[0]["SCAC_OPP"]
    scac = op_code.split("_", 1)[0].upper()

    # BEGIN status
    _update_status(scac, "NIT-BEGIN", log)
    _update_status(scac, "PIT-BEGIN", log)

    success = False
    try:
        for row in rows:
            attempts = 0
            while True:
                attempts += 1
                if process_row(row, enable_upload,
                               root_folder, run_id, log):
                    success = True
                    break
                if attempts >= max_retry:
                    raise RuntimeError("Max retries reached")
                time.sleep(RETRY_SLEEP)

        # COMPLETE status
        _update_status(scac, "NIT-COMPLETE", log)
        _update_status(scac, "PIT-COMPLETE", log)
        log.info("SUCCESS")
        return {
            "Out_strWorkExceptionMessage": "",
            "Out_boolWorkcompleted": True,
            "Out_strLogPath": str(log_file),
        }

    except Exception as exc:
        log.exception("Run failed")
        _reset_status(scac, "NIT", log)
        return {
            "Out_strWorkExceptionMessage": str(exc),
            "Out_boolWorkcompleted": success,
            "Out_strLogPath": str(log_file),
        }
    finally:
        kill_orphan_excels()
        log.info("Log saved to %s", log_file)


# ------------------------------ CLI ---------------------------------------

def _cli() -> None:
    ap = argparse.ArgumentParser(description="Run FM Tool processor")
    ap.add_argument("json_file", help="Payload file or '-' for stdin")
    args = ap.parse_args()

    raw = sys.stdin.read() if args.json_file == "-" else Path(args.json_file).read_text(encoding="utf-8")
    print(json.dumps(run_flow(json.loads(raw)), indent=2))


if __name__ == "__main__":
    _cli()
