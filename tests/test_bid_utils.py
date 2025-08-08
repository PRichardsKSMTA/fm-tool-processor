import logging
import types


def test_insert_bid_rows_early_return(monkeypatch, tmp_path, caplog):
    called = False

    def fake_app(*args, **kwargs):
        nonlocal called
        called = True
        raise AssertionError("App should not be called")

    from fm_tool_core import bid_utils

    monkeypatch.setattr(
        bid_utils, "xw", types.SimpleNamespace(App=fake_app), raising=False
    )

    log = logging.getLogger("test")
    with caplog.at_level(logging.INFO):
        bid_utils.insert_bid_rows(tmp_path / "wb.xlsx", [], log)
    assert not called
    assert "No BID rows to insert" in caplog.text
