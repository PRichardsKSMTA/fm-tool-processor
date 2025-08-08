from pathlib import Path
from types import SimpleNamespace
from unittest.mock import MagicMock

from fm_tool_core import excel_utils


def test_read_cell_initializes_and_uninitializes(monkeypatch):
    pc = SimpleNamespace(CoInitialize=MagicMock(), CoUninitialize=MagicMock())
    monkeypatch.setattr(excel_utils, "pythoncom", pc)

    m_range = MagicMock()
    m_range.value = "v"
    m_sheet = MagicMock()
    m_sheet.range.return_value = m_range
    m_wb = MagicMock()
    m_wb.sheets = {excel_utils.SCAC_VALIDATION_SHEET: m_sheet}
    m_app = MagicMock()
    m_app.api = MagicMock()
    m_app.books.open.return_value = m_wb

    xw_mock = SimpleNamespace(App=MagicMock(return_value=m_app))
    monkeypatch.setattr(excel_utils, "xw", xw_mock)

    result = excel_utils.read_cell(Path("dummy.xlsx"), "A", "1")
    assert result == "v"
    pc.CoInitialize.assert_called_once_with()
    pc.CoUninitialize.assert_called_once_with()
