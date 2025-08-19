from __future__ import annotations

import os
from pathlib import Path

READY_NAME = "PY_READY_FLAG"
READY_OK = "READY"
READY_ERR = "ERROR"

READY_TO = 600  # seconds
OPEN_TO = 60
POLL_SLEEP = 0.25
RETRY_SLEEP = 2

LOG_DIR = Path(os.getenv("LOG_DIR", "../Logs")).resolve()

SP_USERNAME = os.getenv("SP_USERNAME")
SP_PASS = os.getenv("SP_PASS")
ROOT_SP_SITE = "https://ksmcpa.sharepoint.com/teams/ksmta"

SCAC_VALIDATION_SHEET = "HOME"
VISIBLE_EXCEL = os.getenv("FM_SHOW_EXCEL", "0") == "1"

SMTP_SERVER = os.getenv("SMTP_SERVER")
SMTP_PORT = os.getenv("SMTP_PORT")
SMTP_USERNAME = os.getenv("SMTP_USERNAME")
SMTP_PASSWORD = os.getenv("SMTP_PASSWORD")
SMTP_FROM = os.getenv("SMTP_FROM")

__all__ = [
    "READY_NAME",
    "READY_OK",
    "READY_ERR",
    "READY_TO",
    "OPEN_TO",
    "POLL_SLEEP",
    "RETRY_SLEEP",
    "LOG_DIR",
    "SP_USERNAME",
    "SP_PASS",
    "ROOT_SP_SITE",
    "SCAC_VALIDATION_SHEET",
    "VISIBLE_EXCEL",
    "SMTP_SERVER",
    "SMTP_PORT",
    "SMTP_USERNAME",
    "SMTP_PASSWORD",
    "SMTP_FROM",
]
