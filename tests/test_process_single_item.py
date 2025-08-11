"""High-level smoke test for run_flow.

Mocks heavy external dependencies (Excel & SharePoint) so we can assert that:
  * The function returns the expected schema
  * Validation logic branches correctly
"""

from types import SimpleNamespace
from unittest.mock import ANY, patch

import pytest

import fm_tool_core as core

# -------------------- helpers -------------------- #


def _fake_xlwings_macro(*args, **kwargs):  # noqa: D401
    """Pretend macro runs successfully"""
    return None


class _FakeWorkbook:
    sheets = [SimpleNamespace(range=lambda x: SimpleNamespace(value="HUMD_VAN"))]

    def save(self): ...

    def close(self): ...


# ------------------------------------------------- #


@pytest.fixture
def payload(tmp_path):
    tmp_folder = str(tmp_path)
    return {
        "item/In_intMaxRetry": 1,
        "item/In_strDestinationProcessingFolder": tmp_folder,
        "item/In_dtInputData": [
            {
                "SCAC_OPP": "HUMD_VAN",
                "CLIENT_SCAC": "HUMD",
                "KSMTA_DEST_SITE": ("https://example." "sharepoint.com"),
                "KSMTA_DEST_FOLDER_PATH": "/NA",
                "CLIENT_DEST_SITE": ("https://example." "sharepoint.com"),
                "CLIENT_DEST_FOLDER_PATH": "/",
                "FM_TOOL": "PIT",
                "TOOL_TEMPLATE_FILEPATH": __file__,  # any file
                "NEW_EXCEL_FILENAME": "dummy.xlsm",
                "WEEK_CT": "12",
                "PROCESSING_WEEK": "2025-06-14",
                "SCAC_VALIDATION_COLUMN": "A",
                "SCAC_VALIDATION_ROW": "1",
                "ORDERAREAS_VALIDATION_COLUMN": "B",
                "ORDERAREAS_VALIDATION_ROW": "1",
                "ORDERAREAS_VALIDATION_VALUE": ("Input <> " "Order/Area"),
                "CUSTOMER_NAME": "ACME",
            }
        ],
        "item/In_boolEnableSharePointUpload": False,
        "BID-Payload": "123e4567-e89b-12d3-a456-426614174000",
    }


def test_run_flow_success(payload):
    """All validations pass -> Out_boolWorkcompleted=True"""
    with patch(
        "fm_tool_core.process_fm_tool.run_excel_macro",
        return_value=_FakeWorkbook(),
    ) as macro, patch(
        "fm_tool_core.process_fm_tool.read_cell",
        return_value="HUMD_VAN",
    ), patch(
        "fm_tool_core.process_fm_tool.sharepoint_upload"
    ), patch(
        "fm_tool_core.process_fm_tool.sharepoint_file_exists",
        return_value=False,
    ), patch(
        "fm_tool_core.process_fm_tool._fetch_bid_rows",
    ) as fetch_mock, patch(
        "fm_tool_core.process_fm_tool._fetch_customer_ids",
        return_value=["i1", "i2"],
    ) as cid_mock, patch(
        "fm_tool_core.process_fm_tool.write_home_fields",
    ) as write_mock:
        result = core.run_flow(payload)
    macro.assert_called_once()
    args_tuple = macro.call_args[0][1]
    assert len(args_tuple) == 4
    assert args_tuple[-1] == payload["BID-Payload"]
    fetch_mock.assert_not_called()
    cid_mock.assert_called_once_with(payload["BID-Payload"], ANY)
    write_mock.assert_called_once_with(
        ANY, payload["BID-Payload"], "ACME", ["i1", "i2"]
    )
    assert result["Out_boolWorkcompleted"] is True
    assert result["Out_strWorkExceptionMessage"] == ""


def test_run_flow_without_bid_payload(payload):
    """run_excel_macro only receives three args when BID-Payload missing"""

    payload.pop("BID-Payload")
    with patch(
        "fm_tool_core.process_fm_tool.run_excel_macro",
        return_value=_FakeWorkbook(),
    ) as macro, patch(
        "fm_tool_core.process_fm_tool.read_cell",
        return_value="HUMD_VAN",
    ), patch(
        "fm_tool_core.process_fm_tool.sharepoint_upload"
    ), patch(
        "fm_tool_core.process_fm_tool.sharepoint_file_exists",
        return_value=False,
    ), patch(
        "fm_tool_core.process_fm_tool._fetch_customer_ids"
    ) as cid_mock, patch(
        "fm_tool_core.process_fm_tool.write_home_fields",
    ) as write_mock:
        result = core.run_flow(payload)
    macro.assert_called_once()
    assert len(macro.call_args[0][1]) == 3
    write_mock.assert_called_once_with(ANY, None, "ACME", None)
    cid_mock.assert_not_called()
    assert result["Out_boolWorkcompleted"] is True
