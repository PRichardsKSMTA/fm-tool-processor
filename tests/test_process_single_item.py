"""High-level smoke test for run_flow.

Mocks heavy external dependencies (Excel & SharePoint) so we can assert that:
  * The function returns the expected schema
  * Validation logic branches correctly
"""

<<<<<<< HEAD
from types import SimpleNamespace
from unittest.mock import patch

<<<<<<< HEAD
import logging
=======
>>>>>>> 0626e50 (Update VM with current version)
=======
import logging
from types import SimpleNamespace
from unittest.mock import ANY, patch

>>>>>>> refs/remotes/origin/main
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


def test_run_flow_success(payload, caplog):
    """All validations pass -> Out_boolWorkcompleted=True"""
    with caplog.at_level(logging.INFO, logger="fm_tool"), patch(
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
<<<<<<< HEAD
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
=======
    ) as fetch_mock, patch(
        "fm_tool_core.process_fm_tool._fetch_customer_ids",
        return_value=["i1", "i2"],
    ) as cid_mock, patch(
        "fm_tool_core.process_fm_tool.write_home_fields",
    ) as write_mock, patch(
        "fm_tool_core.process_fm_tool._fetch_adhoc_headers",
        return_value={"ADHOC_INFO1": "Origin (Live/Drop)"},
    ) as adhoc_mock, patch(
        "fm_tool_core.process_fm_tool.update_adhoc_headers",
    ) as upd_mock:
>>>>>>> refs/remotes/origin/main
        result = core.run_flow(payload)
    macro.assert_called_once()
    args_tuple = macro.call_args[0][1]
    assert len(args_tuple) == 4
    assert args_tuple[-1] == payload["BID-Payload"]
<<<<<<< HEAD
<<<<<<< HEAD
=======
    write_cell.assert_called_once()
    assert write_cell.call_args[0] == (
        macro.call_args[0][0],
        "BID",
        payload["BID-Payload"],
    )
>>>>>>> 0626e50 (Update VM with current version)
=======
    fetch_mock.assert_not_called()
    cid_mock.assert_called_once_with(payload["BID-Payload"], ANY)
    write_mock.assert_called_once_with(
        ANY,
        payload["BID-Payload"],
        "ACME",
        ["i1", "i2"],
        {"ADHOC_INFO1": "Origin (Live/Drop)"},
    )
    adhoc_mock.assert_called_once_with(payload["BID-Payload"], ANY)
    upd_mock.assert_called_once_with(
        ANY,
        {"ADHOC_INFO1": "Origin (Live/Drop)"},
        ANY,
    )
    assert "Applying ad-hoc headers" in caplog.text
>>>>>>> refs/remotes/origin/main
    assert result["Out_boolWorkcompleted"] is True
    assert result["Out_strWorkExceptionMessage"] == ""


<<<<<<< HEAD
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
=======
>>>>>>> refs/remotes/origin/main
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
=======
    ), patch(
        "fm_tool_core.process_fm_tool._fetch_customer_ids"
    ) as cid_mock, patch(
        "fm_tool_core.process_fm_tool.write_home_fields",
    ) as write_mock, patch(
        "fm_tool_core.process_fm_tool._fetch_adhoc_headers",
    ) as adhoc_mock, patch(
        "fm_tool_core.process_fm_tool.update_adhoc_headers",
    ) as upd_mock:
        result = core.run_flow(payload)
    macro.assert_called_once()
    assert len(macro.call_args[0][1]) == 3
    write_mock.assert_called_once_with(ANY, None, "ACME", None, None)
    cid_mock.assert_not_called()
    adhoc_mock.assert_not_called()
    upd_mock.assert_not_called()
>>>>>>> refs/remotes/origin/main
    assert result["Out_boolWorkcompleted"] is True


def test_run_flow_nit_ignores_bid_payload(payload):
    """BID-Payload is ignored for NIT runs"""

    payload["item/In_dtInputData"][0]["FM_TOOL"] = "NIT"
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
        "fm_tool_core.process_fm_tool.write_home_fields"
    ) as write_mock, patch(
        "fm_tool_core.process_fm_tool._fetch_adhoc_headers"
    ) as adhoc_mock, patch(
        "fm_tool_core.process_fm_tool.update_adhoc_headers"
    ) as upd_mock:
        result = core.run_flow(payload)
    macro.assert_called_once()
    assert len(macro.call_args[0][1]) == 3
    write_mock.assert_called_once_with(ANY, None, "ACME", None, None)
    cid_mock.assert_not_called()
    adhoc_mock.assert_not_called()
    upd_mock.assert_not_called()
    assert result["Out_boolWorkcompleted"] is True


def test_upload_strips_path_from_filename(payload):
    """SharePoint upload uses only the final path component"""

    payload["item/In_boolEnableSharePointUpload"] = True
    payload["item/In_dtInputData"][0]["NEW_EXCEL_FILENAME"] = "x/y/dummy.xlsm"
    with patch(
        "fm_tool_core.process_fm_tool.run_excel_macro",
        return_value=_FakeWorkbook(),
    ), patch(
        "fm_tool_core.process_fm_tool.read_cell",
        return_value="HUMD_VAN",
    ), patch(
        "fm_tool_core.process_fm_tool.sp_ctx",
        return_value=object(),
    ), patch(
        "fm_tool_core.process_fm_tool.sharepoint_file_exists",
        return_value=False,
    ) as exists_mock, patch(
        "fm_tool_core.process_fm_tool.sharepoint_upload"
    ) as upload_mock, patch(
        "fm_tool_core.process_fm_tool._fetch_bid_rows",
    ), patch(
        "fm_tool_core.process_fm_tool._fetch_customer_ids",
        return_value=["i1"],
    ), patch(
        "fm_tool_core.process_fm_tool.write_home_fields",
    ), patch(
        "fm_tool_core.process_fm_tool._fetch_adhoc_headers",
        return_value={},
    ), patch(
        "fm_tool_core.process_fm_tool.update_adhoc_headers",
    ):
        result = core.run_flow(payload)
    upload_mock.assert_called_once()
    assert upload_mock.call_args[0][2] == "dummy.xlsm"
    rel_arg = exists_mock.call_args[0][1]
    assert rel_arg.endswith("/dummy.xlsm")
    assert result["Out_boolWorkcompleted"] is True


def test_run_flow_raises_last_error_message(payload):
    """Final error message includes the last process_row failure"""

    payload["item/In_intMaxRetry"] = 2
    with patch(
        "fm_tool_core.process_fm_tool.process_row",
        side_effect=[RuntimeError("first"), RuntimeError("second")],
    ), patch("fm_tool_core.process_fm_tool.time.sleep"):
        result = core.run_flow(payload)
    assert result["Out_boolWorkcompleted"] is False
    msg = result["Out_strWorkExceptionMessage"]
    assert "second" in msg and "first" not in msg
