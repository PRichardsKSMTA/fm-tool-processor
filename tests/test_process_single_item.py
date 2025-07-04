"""High-level smoke test for run_flow.

Mocks heavy external dependencies (Excel & SharePoint) so we can assert that:
  * The function returns the expected schema
  * Validation logic branches correctly
"""

from types import SimpleNamespace
from unittest.mock import patch

import fm_tool_core as core

# -------------------- helpers -------------------- #


def _fake_xlwings_macro(*args, **kwargs):  # noqa: D401
    """Pretend macro runs successfully"""
    return None


class _FakeWorkbook:
    sheets = [
        SimpleNamespace(range=lambda x: SimpleNamespace(value="HUMD_VAN")),
    ]

    def save(self): ...
    def close(self): ...


# ------------------------------------------------- #


def test_run_flow_success(tmp_path, monkeypatch):
    """All validations pass -> Out_boolWorkcompleted=True"""

    # -- Patch xlwings + SharePoint functions --
    with patch(
        "fm_tool_core.process_fm_tool.run_excel_macro",
        return_value=_FakeWorkbook(),
    ):
        with patch(
            "fm_tool_core.process_fm_tool.read_cell",
            return_value="HUMD_VAN",
        ):
            with patch("fm_tool_core.process_fm_tool.sharepoint_upload"):
                with patch(
                    "fm_tool_core.process_fm_tool.sharepoint_file_exists",
                    return_value=False,
                ):
                    tmp_folder = str(tmp_path)
                    payload = {
                        "item/In_intMaxRetry": 1,
                        "item/In_strDestinationProcessingFolder": tmp_folder,
                        "item/In_dtInputData": [
                            {
                                "SCAC_OPP": "HUMD_VAN",
                                "CLIENT_SCAC": "HUMD",
                                "KSMTA_DEST_SITE": (
                                    "https://example." "sharepoint.com"
                                ),
                                "KSMTA_DEST_FOLDER_PATH": "/NA",
                                "CLIENT_DEST_SITE": (
                                    "https://example." "sharepoint.com"
                                ),
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
                                "ORDERAREAS_VALIDATION_VALUE": (
                                    "Input <> " "Order/Area"
                                ),
                            }
                        ],
                        "item/In_boolEnableSharePointUpload": False,
                    }

                    result = core.run_flow(payload)
                    assert result["Out_boolWorkcompleted"] is True
                    assert result["Out_strWorkExceptionMessage"] == ""
