# Enable postponed evaluation of type annotations
from __future__ import annotations

# Access environment variables
import os

# Use Path objects for filesystem paths
from pathlib import Path

# Name of the Excel named range used as a ready flag
READY_NAME = "PY_READY_FLAG"
# Expected value of the ready flag when work is done
READY_OK = "READY"
# Expected value of the ready flag when VBA reports an error
READY_ERR = "ERROR"

# How long to wait for the VBA ready flag (in seconds)
READY_TO = 600  # seconds
# Timeout for opening Excel
OPEN_TO = 60
# Delay between polls for the ready flag
POLL_SLEEP = 0.25
# Pause before retrying failed work
RETRY_SLEEP = 2

# Directory where log files are written
LOG_DIR = Path(os.getenv("LOG_DIR", "./logs")).resolve()

# Credentials for SharePoint access
SP_USERNAME = os.getenv("SP_USERNAME")
SP_PASSWORD = os.getenv("SP_PASSWORD")
# Base URL for SharePoint
ROOT_SP_SITE = "https://ksmcpa.sharepoint.com/teams/ksmta"

# Worksheet used to read validation values
SCAC_VALIDATION_SHEET = "HOME"
# Show Excel when reading validation if env var is set to "1"
VISIBLE_EXCEL = os.getenv("FM_SHOW_EXCEL", "0") == "1"

# Names exported when `from fm_tool_core.constants import *` is used
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
    "SP_PASSWORD",
    "ROOT_SP_SITE",
    "SCAC_VALIDATION_SHEET",
    "VISIBLE_EXCEL",
]
