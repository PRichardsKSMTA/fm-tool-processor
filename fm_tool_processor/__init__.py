"""
HTTP-triggered Azure Function wrapper for process_fm_tool.run_flow.
Deploy on a Windows Premium/Dedicated plan where Excel is installed.
"""

import json
import logging
from pathlib import Path

import azure.functions as func

from process_fm_tool import run_flow

def main(req: func.HttpRequest) -> func.HttpResponse:  # noqa: N802 – Azure sig
    logging.info("fm_tool_processor function triggered")

    try:
        payload = req.get_json()
    except ValueError:
        return func.HttpResponse("Invalid JSON body", status_code=400)

    result = run_flow(payload)
    status  = 200 if result["Out_boolWorkcompleted"] else 500
    return func.HttpResponse(
        json.dumps(result, ensure_ascii=False),
        status_code=status,
        mimetype="application/json",
    )
