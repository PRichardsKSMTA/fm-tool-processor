#!/usr/bin/env python3
"""
Run a single FM-Tool payload that was dropped to disk by Power Automate.

How it fits into the pipeline
─────────────────────────────
1. **Power Automate Cloud** writes each incoming JSON payload to a file.
2. **PowerShell wrapper (process_new_payloads.ps1)** finds the file and
   invokes *this* script.
3. We load environment variables (.env), import the heavy-lifting
   `fm_tool_core.process_fm_tool.run_flow` function, and hand the JSON to it.
4. The resulting JSON is printed to **stdout** so the caller can decide
   whether the run succeeded and where to archive the payload file.

The script is intentionally thin: all Excel / SharePoint / SQL logic lives in
`fm_tool_core` and is unit-tested separately.
"""

from __future__ import annotations

import argparse       # robust CLI parsing
import json           # read / write JSON
import sys            # access sys.path & printing
from pathlib import Path
from typing import Any

from dotenv import load_dotenv  # read secrets from .env files

# ───────────────────────────── CONFIGURATION ──────────────────────────────
# 1) Tell Python where the *package code* lives so an editable checkout or a
#    Function App deployment both just work. Adjust if your folder differs.
sys.path.append("C:\\Tasks\\ExcelAutomation\\PythonScript\\fm-tool-processor")

# 2) Load .env **once** so every downstream import can rely on the vars.
load_dotenv(Path("C:\\Tasks\\ExcelAutomation\\PythonScript\\fm-tool-processor") / ".env")

# 3) Import the one public entry point of the FM-Tool engine.
from fm_tool_core.process_fm_tool import run_flow  # noqa: E402  (after sys.path!)

# ────────────────────────────── MAIN PROGRAM ──────────────────────────────
def main(argv: list[str] | None = None) -> None:
    """Parse CLI args, read payload JSON, call `run_flow`, echo result."""
    parser = argparse.ArgumentParser(
        description="Run FM Tool for the given JSON payload file"
    )
    parser.add_argument(
        "-i",
        "--input-file",
        required=True,
        help="Absolute path to the JSON payload created by Power Automate",
    )

    args = parser.parse_args(argv)

    # Read the payload (UTF-8 so everything is cross-platform safe)
    payload_path = Path(args.input_file)
    with payload_path.open("r", encoding="utf-8") as f:
        payload: dict[str, Any] = json.load(f)

    # Execute the heavy lifting.
    result = run_flow(payload)

    # Emit JSON so the PowerShell wrapper can parse success / failure.
    print(json.dumps(result))


# Standard Python entry-point guard.
if __name__ == "__main__":
    main()
