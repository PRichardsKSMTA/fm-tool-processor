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


def test_write_home_fields(monkeypatch, tmp_path):
    pc = SimpleNamespace(CoInitialize=MagicMock(), CoUninitialize=MagicMock())
    monkeypatch.setattr(excel_utils, "pythoncom", pc)

    bid_rng = SimpleNamespace(value=None)
    d8_rng = SimpleNamespace(value=None)

    def range_side_effect(addr):
        if addr == "BID":
            return bid_rng
        if addr == "D8":
            return d8_rng
        raise AssertionError

    sheet = SimpleNamespace(range=MagicMock(side_effect=range_side_effect))
    wb = SimpleNamespace(sheets={"HOME": sheet}, save=MagicMock(), close=MagicMock())
    app = SimpleNamespace(
        api=SimpleNamespace(),
        books=SimpleNamespace(open=MagicMock(return_value=wb)),
        kill=MagicMock(),
    )

    xw_mock = SimpleNamespace(App=MagicMock(return_value=app))
    monkeypatch.setattr(excel_utils, "xw", xw_mock)

    excel_utils.write_home_fields(tmp_path / "wb.xlsx", "pg", "cust")
    assert bid_rng.value == "pg"
    assert d8_rng.value == "cust"
    wb.save.assert_called_once_with()
    wb.close.assert_called_once_with()
    app.kill.assert_called_once_with()
    pc.CoInitialize.assert_called_once_with()
    pc.CoUninitialize.assert_called_once_with()
