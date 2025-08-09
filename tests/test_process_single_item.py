"""High-level smoke test for run_flow.

Mocks heavy external dependencies (Excel & SharePoint) so we can assert that:
  * The function returns the expected schema
  * Validation logic branches correctly
"""

from types import SimpleNamespace
from unittest.mock import patch

<<<<<<< HEAD
import logging
=======
>>>>>>> 0626e50 (Update VM with current version)
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
            }
        ],
        "item/In_boolEnableSharePointUpload": False,
        "BID-Payload": "123e4567-e89b-12d3-a456-426614174000",
    }


def test_run_flow_success(payload):
    """All validations pass -> Out_boolWorkcompleted=True"""

    bid_rows = [
        {"Lane ID": 1, "Orig Zip (5 or 3)": "12345", "Dest Zip (5 or 3)": "54321"}
    ]
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
        return_value=bid_rows,
    ), patch(
<<<<<<< HEAD
        "fm_tool_core.process_fm_tool._fetch_adhoc_headers",
        return_value={},
    ), patch(
=======
<<<<<<< HEAD
>>>>>>> aa4e400 (resolve merge conflict)
        "fm_tool_core.process_fm_tool.insert_bid_rows",
    ):
=======
        "fm_tool_core.process_fm_tool.insert_bid_rows"
    ), patch(
        "fm_tool_core.process_fm_tool.write_named_cell"
    ) as write_cell:
>>>>>>> 0626e50 (Update VM with current version)
        result = core.run_flow(payload)
    macro.assert_called_once()
    args_tuple = macro.call_args[0][1]
    assert len(args_tuple) == 4
    assert args_tuple[-1] == payload["BID-Payload"]
<<<<<<< HEAD
=======
    write_cell.assert_called_once()
    assert write_cell.call_args[0] == (
        macro.call_args[0][0],
        "BID",
        payload["BID-Payload"],
    )
>>>>>>> 0626e50 (Update VM with current version)
    assert result["Out_boolWorkcompleted"] is True
    assert result["Out_strWorkExceptionMessage"] == ""


<<<<<<< HEAD
def test_run_flow_inserts_bid_rows(payload, caplog):
=======
def test_run_flow_inserts_bid_rows(payload):
>>>>>>> 0626e50 (Update VM with current version)
    """run_flow fetches BID rows and inserts them once"""

    bid_rows = [
        {"Lane ID": 1, "Orig Zip (5 or 3)": "11111", "Dest Zip (5 or 3)": "22222"}
    ]
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
        return_value=bid_rows,
    ), patch(
        "fm_tool_core.process_fm_tool._fetch_adhoc_headers",
        return_value={"ADHOC_INFO1": "X1"},
    ), patch(
        "fm_tool_core.process_fm_tool.insert_bid_rows"
<<<<<<< HEAD
    ) as insert_mock:
        with caplog.at_level(logging.INFO):
            result = core.run_flow(payload)
    insert_mock.assert_called_once()
    assert insert_mock.call_args[0][3] == {"ADHOC_INFO1": "X1"}
    macro.assert_called_once()
    assert len(macro.call_args[0][1]) == 4
    assert any("Fetched 1 BID rows in" in rec.message for rec in caplog.records)
    assert any("Batch inserted 1 BID rows in" in rec.message for rec in caplog.records)
    assert result["Out_boolWorkcompleted"] is True


def test_run_flow_skips_insert_when_no_rows(payload, caplog):
    """insert_bid_rows is not called when no BID rows fetched"""

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
        return_value=[],
    ), patch(
        "fm_tool_core.process_fm_tool._fetch_adhoc_headers",
        return_value={},
    ), patch(
        "fm_tool_core.process_fm_tool.insert_bid_rows"
    ) as insert_mock:
        with caplog.at_level(logging.INFO):
            core.run_flow(payload)
    insert_mock.assert_not_called()
    macro.assert_called_once()
    assert any("Fetched 0 BID rows in" in rec.message for rec in caplog.records)
    assert not any("Batch inserted" in rec.message for rec in caplog.records)


=======
    ) as insert_mock, patch(
        "fm_tool_core.process_fm_tool.write_named_cell"
    ) as write_cell:
        result = core.run_flow(payload)
    insert_mock.assert_called_once()
    macro.assert_called_once()
    assert len(macro.call_args[0][1]) == 4
    write_cell.assert_called_once()
    assert write_cell.call_args[0] == (
        macro.call_args[0][0],
        "BID",
        payload["BID-Payload"],
    )
    assert result["Out_boolWorkcompleted"] is True


>>>>>>> 0626e50 (Update VM with current version)
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
<<<<<<< HEAD
    ):
        result = core.run_flow(payload)
    macro.assert_called_once()
    assert len(macro.call_args[0][1]) == 3
=======
    ), patch(
        "fm_tool_core.process_fm_tool.write_named_cell"
    ) as write_cell:
        result = core.run_flow(payload)
    macro.assert_called_once()
    assert len(macro.call_args[0][1]) == 3
    write_cell.assert_not_called()
>>>>>>> 0626e50 (Update VM with current version)
    assert result["Out_boolWorkcompleted"] is True
