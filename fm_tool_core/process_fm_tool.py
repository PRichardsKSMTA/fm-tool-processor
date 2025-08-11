#!/usr/bin/env python3
"""
FM Tool processing entry point – verbose logging & SQL status tracking
=====================================================================

Key behaviour (July 2025)
────────────────────────
* **Payload-type detection**
    • We inspect the `FM_TOOL` column of the first payload row.
      ─ `"NIT"`  → NIT run
      ─ `"PIT"` *or anything else / missing* → PIT run (default)
* **Mutually-exclusive status procs** – exactly one `…-BEGIN` at start and one
  `…-COMPLETE` in the `finally` block.
* **RESET removed** – `dbo.RESET_CLIENT_PROCESSING_STATUS` exists for legacy
  reasons but is **never called**.
* Failed runs are still surfaced to Power Automate, yet marked COMPLETE so
  they do not auto-retry.

Excel, SharePoint, logging and CLI behaviour remain unchanged.
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

# ───────────────────────────── psutil (optional) ───────────────────────────
try:
    import psutil  # type: ignore
except Exception as _e:  # pragma: no cover
    psutil = None  # type: ignore
    logging.basicConfig(level=logging.WARNING)
    logging.warning("psutil missing – CPU checks disabled (%s)", _e)

# ────────────────────────────── pyodbc (optional) ──────────────────────────
try:
    import pyodbc  # type: ignore
except ImportError as _e:  # pragma: no cover
    pyodbc = None  # type: ignore
    logging.basicConfig(level=logging.WARNING)
    logging.warning("pyodbc missing – SQL disabled (%s)", _e)

try:
    from .bid_utils import _COLUMNS
except Exception as _e:  # pragma: no cover
    _COLUMNS = []  # type: ignore
    logging.basicConfig(level=logging.WARNING)
    logging.warning("BID utils unavailable: %s", _e)
from .constants import LOG_DIR, RETRY_SLEEP
from .excel_utils import (
    copy_template,
    kill_orphan_excels,
    read_cell,
    run_excel_macro,
    write_home_fields,
)
from .exceptions import FlowError
from .sharepoint_utils import sp_ctx, sharepoint_file_exists, sharepoint_upload

# ───────────────────────────── SQL HELPERS ─────────────────────────────────
_SQL_CONN_STR: str | None = None


def _sql_conn_str() -> str:
    """Build (and cache) an ODBC connection string."""
    global _SQL_CONN_STR
    if _SQL_CONN_STR:
        return _SQL_CONN_STR
    srv, db, usr, pwd = (
        os.getenv(k)
        for k in ("SQL_SERVER", "SQL_DATABASE", "SQL_USERNAME", "SQL_PASSWORD")
    )
    if not all((srv, db, usr, pwd)):
        raise RuntimeError("SQL connection env vars missing")
    _SQL_CONN_STR = (
        f"DRIVER={{ODBC Driver 18 for SQL Server}};SERVER={srv};DATABASE={db};"
        f"UID={usr};PWD={pwd};Encrypt=yes;TrustServerCertificate=yes;"
    )
    return _SQL_CONN_STR


def _exec_proc(proc: str, params: tuple[Any, ...], log: logging.Logger) -> None:
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


def _update_status(scac: str, state: str, log: logging.Logger) -> None:
    _exec_proc(
        os.getenv("SQL_UPDATE_PROC", "dbo.UPDATE_CLIENT_UPLOAD_STATUS"),
        (scac, state),
        log,
    )


# RETAINED for completeness – **not used**
def _reset_status(scac: str, state: str, log: logging.Logger) -> None:  # noqa: F401
    _exec_proc(
        os.getenv("SQL_RESET_PROC", "dbo.RESET_CLIENT_PROCESSING_STATUS"),
        (scac, state),
        log,
    )


# ───────────────────────────── HELPER FUNCTIONS ────────────────────────────
def _fifo_sort(rows: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Keep FIFO order based on first timestamp-like key, else input order."""
    key = next(
        (
            k
            for k in ("QUEUE_TS", "CREATED_UTC", "CREATED_AT")
            if all(k in r for r in rows)
        ),
        None,
    )
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


def wait_for_cpu(
    max_percent: float = 80.0,
    interval: float = 1.0,
    backoff: float = 5.0,
    log: logging.Logger | None = None,
) -> None:
    """Block until overall CPU utilisation drops below *max_percent*."""
    if psutil is None:
        if log:
            log.info("(psutil missing – skipping CPU wait)")
        return
    while True:
        cpu = psutil.cpu_percent(interval)
        if cpu < max_percent:
            if log:
                log.info("CPU load at %.1f%% < %.1f%%, proceeding", cpu, max_percent)
            return
        if log:
            log.warning("High CPU (%.1f%%) – sleeping %ss", cpu, backoff)
        time.sleep(backoff)


def _detect_payload_type(rows: List[Dict[str, Any]]) -> str:
    """
    Determine whether the payload is PIT or NIT.

    Logic: read `FM_TOOL` (case-insensitive) from the first row:
        * "NIT" → NIT
        * "PIT" *or missing / anything else* → PIT (default)
    """
    if not rows:
        return "PIT"
    val = str(rows[0].get("FM_TOOL", "")).strip().upper()
    return "NIT" if val == "NIT" else "PIT"


def _fetch_bid_rows(process_guid: str, log: logging.Logger) -> List[Dict[str, Any]]:
    """
    Return BID rows for *process_guid*.

    The SELECT **order must match** _COLUMNS in bid_utils.py.
    """
    if pyodbc is None:
        log.info("(SQL disabled) would fetch BID rows for %s", process_guid)
        return []

    conn = None
    try:
        conn = pyodbc.connect(_sql_conn_str(), timeout=10)
        with conn.cursor() as cur:
            cur.execute(
                """
                SELECT LANE_ID, ORIG_CITY, ORIG_ST, ORIG_POSTAL_CD,
                       DEST_CITY, DEST_ST, DEST_POSTAL_CD,
                       BID_VOLUME, LH_RATE, RFP_MILES,
                       FREIGHT_TYPE, TEMP_CAT, BTF_FSC_PER_MILE,
                       ADHOC_INFO1, ADHOC_INFO2, ADHOC_INFO3, ADHOC_INFO4, ADHOC_INFO5,
                       ADHOC_INFO6, ADHOC_INFO7, ADHOC_INFO8, ADHOC_INFO9, ADHOC_INFO10,
                       FM_MILES, FM_TOLLS
                FROM dbo.RFP_OBJECT_DATA
                WHERE PROCESS_GUID = ?
                """,
                (process_guid,),
            )
            return [dict(zip(_COLUMNS, row)) for row in cur.fetchall()]
    except Exception as exc:
        log.warning("Failed to fetch BID rows: %s", exc)
        return []
    finally:
        if conn:
            try:
                conn.close()
            except Exception:
                pass


def _fetch_adhoc_headers(process_guid: str, log: logging.Logger) -> Dict[str, str]:
    """Return ADHOC_INFO* labels for *process_guid*."""
    if pyodbc is None:
        log.info("(SQL disabled) would fetch ad-hoc headers for %s", process_guid)
        return {}

    conn = None
    try:
        conn = pyodbc.connect(_sql_conn_str(), timeout=10)
        with conn.cursor() as cur:
            cur.execute(
                "SELECT PROCESS_JSON FROM dbo.MAPPING_AGENT_PROCESSES "
                "WHERE PROCESS_GUID = ?",
                (process_guid,),
            )
            row = cur.fetchone()
            if not row or not row[0]:
                return {}
            try:
                data = json.loads(row[0])
            except Exception as exc:
                log.warning("Malformed PROCESS_JSON: %s", exc)
                return {}
            return {
                k: v
                for k, v in (
                    (f"ADHOC_INFO{i}", data.get(f"ADHOC_INFO{i}")) for i in range(1, 11)
                )
                if isinstance(v, str)
            }
    except Exception as exc:
        log.warning("Failed to fetch ad-hoc headers: %s", exc)
        return {}
    finally:
        if conn:
            try:
                conn.close()
            except Exception:
                pass


def _fetch_customer_ids(process_guid: str, log: logging.Logger) -> List[str]:
    """Return up to five CUSTOMER_IDs for *process_guid*."""
    if pyodbc is None:
        log.info("(SQL disabled) would fetch CUSTOMER_ID for %s", process_guid)
        return []

    conn = None
    try:
        conn = pyodbc.connect(_sql_conn_str(), timeout=10)
        with conn.cursor() as cur:
            cur.execute(
                "SELECT TOP 1 CUSTOMER_ID FROM dbo.RFP_OBJECT_DATA "
                "WHERE PROCESS_GUID = ?",
                (process_guid,),
            )
            row = cur.fetchone()
            if not row or not row[0]:
                return []
            raw = str(row[0])
            parts = [p.strip() for p in raw.replace(";", ",").split(",")]
            return [p for p in parts if p][:5]
    except Exception as exc:
        log.warning("Failed to fetch CUSTOMER_ID: %s", exc)
        return []
    finally:
        if conn:
            try:
                conn.close()
            except Exception:
                pass


# ───────────────────────────── ROW WORKER ──────────────────────────────────
def process_row(
    row: Dict[str, Any],
    upload: bool,
    root: str,
    run_id: str,
    log: logging.Logger,
    bid_guid: str | None = None,
) -> bool:
    """Process one FM payload row. Returns True on success."""
    op_code = row["SCAC_OPP"]
    template_src = row["TOOL_TEMPLATE_FILEPATH"]

    log.info("Copying template from %s", template_src)
    dst_path = copy_template(
        template_src, root, f"{Path(row['NEW_EXCEL_FILENAME']).stem}_{run_id}.xlsm", log
    )
    log.info("Template copied to %s", dst_path)

    cust_ids: List[str] | None = None
    if bid_guid is not None:
        cust_ids = _fetch_customer_ids(bid_guid, log)
    write_home_fields(dst_path, bid_guid, row.get("CUSTOMER_NAME"), cust_ids)

    log.info("Waiting for CPU to drop")
    wait_for_cpu(log=log)
    kill_orphan_excels()

    try:
        log.info("Opening workbook …")
        macro_args = (
            row["SCAC_OPP"],
            row["WEEK_CT"],
            row["PROCESSING_WEEK"],
        )
        if bid_guid is not None:
            macro_args += (bid_guid,)
        run_excel_macro(dst_path, macro_args, log)
        log.info("Running macro PopulateAndRunReport …")

        log.info("Reading validation …")
        op_val = read_cell(
            dst_path, row["SCAC_VALIDATION_COLUMN"], row["SCAC_VALIDATION_ROW"]
        )
        oa_val = read_cell(
            dst_path,
            row["ORDERAREAS_VALIDATION_COLUMN"],
            row["ORDERAREAS_VALIDATION_ROW"],
        )
        log.info(
            "Validation: %s=%s, %s=%s",
            f"{row['SCAC_VALIDATION_COLUMN']}{row['SCAC_VALIDATION_ROW']}",
            op_val,
            f"{row['ORDERAREAS_VALIDATION_COLUMN']}{row['ORDERAREAS_VALIDATION_ROW']}",
            oa_val,
        )

        if op_val != op_code or oa_val == row["ORDERAREAS_VALIDATION_VALUE"]:
            raise FlowError("Validation failed", work_completed=False)

        if upload:
            ctx = sp_ctx(row["CLIENT_DEST_SITE"])
            site_path = urlparse(row["CLIENT_DEST_SITE"]).path
            folder = (
                Path(site_path) / row["CLIENT_DEST_FOLDER_PATH"].lstrip("/")
            ).as_posix()
            rel_file = f"{folder}/{row['NEW_EXCEL_FILENAME']}"
            log.info("Uploading to %s", rel_file)
            if sharepoint_file_exists(ctx, rel_file):
                log.info("SharePoint file exists – skipping upload " "(not an error)")
            else:
                sharepoint_upload(ctx, folder, row["NEW_EXCEL_FILENAME"], dst_path)
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
    """Entry point when invoked by the PowerShell wrapper."""
    if "parameters" in payload and isinstance(payload["parameters"], dict):
        payload = payload["parameters"]

    rows = _fifo_sort(payload["item/In_dtInputData"])
    root_folder = payload["item/In_strDestinationProcessingFolder"]
    enable_upload = payload.get("item/In_boolEnableSharePointUpload", True)
    max_retry = int(payload.get("item/In_intMaxRetry", 1))
    bid_guid = payload.get("BID-Payload")

    run_id = uuid.uuid4().hex[:8]
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    log_file = LOG_DIR / f"{datetime.utcnow():%Y-%m-%d-%H-%M-%S}_{run_id}.log"

    log = logging.getLogger("fm_tool")
    log.handlers.clear()
    log.setLevel(logging.INFO)
    fmt = logging.Formatter(
        "%(asctime)s | %(levelname)s | %(message)s", "%Y-%m-%dT%H:%M:%SZ"
    )
    for h in (
        logging.StreamHandler(sys.stderr),
        logging.FileHandler(log_file, encoding="utf-8"),
    ):
        h.setFormatter(fmt)
        log.addHandler(h)

    log.info("----- FM Tool run %s -----", run_id)

    op_code = rows[0]["SCAC_OPP"]
    scac = op_code.split("_", 1)[0].upper()
    payload_type = _detect_payload_type(rows)  # 'PIT' or 'NIT'
    log.info("Detected payload type: %s", payload_type)

    # BEGIN status (single proc)
    _update_status(scac, f"{payload_type}-BEGIN", log)

    success = False
    try:
        for row in rows:
            attempts = 0
            while True:
                attempts += 1
                if process_row(row, enable_upload, root_folder, run_id, log, bid_guid):
                    success = True
                    break
                if attempts >= max_retry:
                    raise RuntimeError("Max retries reached")
                time.sleep(RETRY_SLEEP)

        log.info("SUCCESS")
        return {
            "Out_strWorkExceptionMessage": "",
            "Out_boolWorkcompleted": True,
            "Out_strLogPath": str(log_file),
        }

    except Exception as exc:
        log.exception("Run failed")
        return {
            "Out_strWorkExceptionMessage": str(exc),
            "Out_boolWorkcompleted": success,
            "Out_strLogPath": str(log_file),
        }

    finally:
        # COMPLETE status – always executed
        try:
            _update_status(scac, f"{payload_type}-COMPLETE", log)
        except Exception:
            log.exception("Failed to mark %s-COMPLETE in SQL", payload_type)
        kill_orphan_excels()
        log.info("Log saved to %s", log_file)


# ------------------------------ CLI ---------------------------------------


def _cli() -> None:
    ap = argparse.ArgumentParser(description="Run FM Tool processor")
    ap.add_argument("json_file", help="Payload file or '-' for stdin")
    args = ap.parse_args()

    raw = (
        sys.stdin.read()
        if args.json_file == "-"
        else Path(args.json_file).read_text(encoding="utf-8")
    )
    print(json.dumps(run_flow(json.loads(raw)), indent=2))


if __name__ == "__main__":
    _cli()
