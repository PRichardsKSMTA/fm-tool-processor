"""Tests for success/failure email notifications."""

from __future__ import annotations

from types import SimpleNamespace
from unittest.mock import ANY, patch

import pytest

import fm_tool_core as core
from fm_tool_core.exceptions import FlowError


class _FakeWorkbook:
    """Minimal stub used by run_flow tests."""

    sheets = [SimpleNamespace(range=lambda _: SimpleNamespace(value="HUMD_VAN"))]

    def save(self): ...

    def close(self): ...


@pytest.fixture
def payload(tmp_path):
    folder = str(tmp_path)
    return {
        "item/In_intMaxRetry": 1,
        "item/In_strDestinationProcessingFolder": folder,
        "item/In_dtInputData": [
            {
                "SCAC_OPP": "HUMD_VAN",
                "CLIENT_DEST_SITE": "https://example.sharepoint.com",
                "CLIENT_DEST_FOLDER_PATH": "/NA",
                "FM_TOOL": "PIT",
                "TOOL_TEMPLATE_FILEPATH": __file__,
                "NEW_EXCEL_FILENAME": "dummy.xlsm",
                "WEEK_CT": "12",
                "PROCESSING_WEEK": "2025-06-14",
                "SCAC_VALIDATION_COLUMN": "A",
                "SCAC_VALIDATION_ROW": "1",
                "ORDERAREAS_VALIDATION_COLUMN": "B",
                "ORDERAREAS_VALIDATION_ROW": "1",
                "ORDERAREAS_VALIDATION_VALUE": "Input <> Order/Area",
                "CUSTOMER_NAME": "ACME",
            }
        ],
        "item/In_boolEnableSharePointUpload": True,
    }


def test_success_email_sent(monkeypatch, payload):
    """send_success_email called with filename and SharePoint URL."""

    monkeypatch.setenv("NOTIFY_EMAIL", "notify@example.com")
    with patch(
        "fm_tool_core.process_fm_tool.run_excel_macro", return_value=_FakeWorkbook()
    ), patch(
        "fm_tool_core.process_fm_tool.read_cell", side_effect=["HUMD_VAN", "ok"]
    ), patch(
        "fm_tool_core.process_fm_tool.sp_ctx"
    ), patch(
        "fm_tool_core.process_fm_tool.sharepoint_file_exists", return_value=False
    ), patch(
        "fm_tool_core.process_fm_tool.sharepoint_upload"
    ), patch(
        "fm_tool_core.process_fm_tool.write_home_fields"
    ), patch(
        "fm_tool_core.process_fm_tool.wait_for_cpu"
    ), patch(
        "fm_tool_core.process_fm_tool.kill_orphan_excels"
    ), patch(
        "fm_tool_core.process_fm_tool.send_success_email"
    ) as success, patch(
        "fm_tool_core.process_fm_tool.send_failure_email"
    ) as failure:
        core.run_flow(payload)

    success.assert_called_once_with(
        "notify@example.com",
        "dummy.xlsm",
        "https://example.sharepoint.com/NA/dummy.xlsm",
        ANY,
    )
    failure.assert_not_called()


def test_failure_email_sent(monkeypatch, payload):
    """send_failure_email called when FlowError raised."""

    monkeypatch.setenv("NOTIFY_EMAIL", "notify@example.com")
    with patch(
        "fm_tool_core.process_fm_tool.process_row",
        side_effect=FlowError("boom", work_completed=False),
    ), patch("fm_tool_core.process_fm_tool.kill_orphan_excels"), patch(
        "fm_tool_core.process_fm_tool.send_success_email"
    ) as success, patch(
        "fm_tool_core.process_fm_tool.send_failure_email"
    ) as failure:
        core.run_flow(payload)

    failure.assert_called_once()
    args = failure.call_args[0]
    assert args[0] == "notify@example.com"
    assert "boom" in args[1]
    success.assert_not_called()
