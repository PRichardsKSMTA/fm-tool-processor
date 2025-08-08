import logging
import sys
import types


def test_insert_bid_rows_early_return(monkeypatch, tmp_path, caplog):
    called = False

    def fake_load_workbook(*args, **kwargs):
        nonlocal called
        called = True
        raise AssertionError("load_workbook should not be called")

    dummy = types.SimpleNamespace(load_workbook=fake_load_workbook)
    monkeypatch.setitem(sys.modules, "openpyxl", dummy)

    from fm_tool_core import bid_utils

    log = logging.getLogger("test")
    with caplog.at_level(logging.INFO):
        bid_utils.insert_bid_rows(tmp_path / "wb.xlsx", [], log)
    assert not called
    assert "No BID rows to insert" in caplog.text
