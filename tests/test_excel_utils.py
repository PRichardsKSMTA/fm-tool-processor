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

    validation_obj = object()
    cust_range = SimpleNamespace(value=None)
    cust_range.api = SimpleNamespace(Validation=validation_obj)

    cells = {
        "BID": SimpleNamespace(value=None),
        "D8:H8": cust_range,
        "D10": SimpleNamespace(value=None),
        "E10": SimpleNamespace(value=None),
        "F10": SimpleNamespace(value=None),
        "G10": SimpleNamespace(value=None),
        "H10": SimpleNamespace(value=None),
        "AR36": SimpleNamespace(value=None),
        "AR37": SimpleNamespace(value=None),
        "AR38": SimpleNamespace(value=None),
        "AR39": SimpleNamespace(value=None),
        "AR40": SimpleNamespace(value=None),
        "AR41": SimpleNamespace(value=None),
        "AR42": SimpleNamespace(value=None),
        "AR43": SimpleNamespace(value=None),
        "AR44": SimpleNamespace(value=None),
        "AR45": SimpleNamespace(value=None),
    }

    def range_side_effect(addr):
        return cells[addr]

    sheet = SimpleNamespace(range=MagicMock(side_effect=range_side_effect))
    wb = SimpleNamespace(sheets={"HOME": sheet}, save=MagicMock(), close=MagicMock())
    app = SimpleNamespace(
        api=SimpleNamespace(),
        books=SimpleNamespace(open=MagicMock(return_value=wb)),
        kill=MagicMock(),
    )

    xw_mock = SimpleNamespace(App=MagicMock(return_value=app))
    monkeypatch.setattr(excel_utils, "xw", xw_mock)

    ids = ["c1", "c2", "c3", "c4", "c5"]
    headers = {"ADHOC_INFO1": "A1", "ADHOC_INFO5": "A5", "ADHOC_INFO10": "A10"}
    excel_utils.write_home_fields(tmp_path / "wb.xlsx", "pg", "cust", ids, headers)
    assert cells["BID"].value == "pg"
    assert cells["D8:H8"].value == "cust"
    assert cells["D8:H8"].api.Validation is validation_obj
    assert cells["D10"].value == "c1"
    assert cells["E10"].value == "c2"
    assert cells["F10"].value == "c3"
    assert cells["G10"].value == "c4"
    assert cells["H10"].value == "c5"
    assert cells["AR36"].value == "A1"
    assert cells["AR40"].value == "A5"
    assert cells["AR45"].value == "A10"
    for addr in [
        "AR37",
        "AR38",
        "AR39",
        "AR41",
        "AR42",
        "AR43",
        "AR44",
    ]:
        assert cells[addr].value is None
    wb.save.assert_called_once_with()
    wb.close.assert_called_once_with()
    app.kill.assert_called_once_with()
    pc.CoInitialize.assert_called_once_with()
    pc.CoUninitialize.assert_called_once_with()
