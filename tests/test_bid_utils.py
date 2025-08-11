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
    assert "No RFP rows to insert" in caplog.text


def test_insert_bid_rows_writes_rows(monkeypatch, tmp_path, caplog):
    from fm_tool_core import bid_utils

    wb_path = tmp_path / "wb.xlsx"
    wb_path.touch()
    calls: list[str] = []

    class FakeApi:
        def __init__(self):
            self.Rows = types.SimpleNamespace(Count=1)

        def Cells(self, _row, _col):
            end = lambda _dir: types.SimpleNamespace(Row=1)
            return types.SimpleNamespace(End=end)

    class FakeRange:
        def __init__(self):
            self._value = None

        def resize(self, _r, _c):
            return self

        @property
        def value(self):
            return self._value

        @value.setter
        def value(self, val):
            calls.append("write")
            self._value = val

    class FakeSheet:
        def __init__(self):
            self.api = FakeApi()

        def range(self, _addr):
            return FakeRange()

    class FakeBook:
        def __init__(self):
            self.sheets = {"RFP": FakeSheet()}

        def save(self):
            pass

        def close(self):
            pass

    def fake_open(_path):
        return FakeBook()

    class FakeBooks:
        def open(self, path):
            return fake_open(path)

    class FakeApp:
        def __init__(self, *args, **kwargs):
            self.api = types.SimpleNamespace(DisplayAlerts=False)
            self.books = FakeBooks()

        def kill(self):
            pass

    monkeypatch.setattr(
        bid_utils,
        "xw",
        types.SimpleNamespace(App=FakeApp),
        raising=False,
    )
    monkeypatch.setattr(
        bid_utils,
        "pythoncom",
        types.SimpleNamespace(
            CoInitialize=lambda: None,
            CoUninitialize=lambda: None,
        ),
        raising=False,
    )

    rows = [
        {
            "LANE_ID": "1",
            "ORIG_POSTAL_CD": "11111",
            "DEST_POSTAL_CD": "22222",
        }
    ]
    log = logging.getLogger("test")
    with caplog.at_level(logging.INFO):
        bid_utils.insert_bid_rows(wb_path, rows, log)
    assert calls == ["write"]
    assert any("RFP sheet" in r.message for r in caplog.records)


def test_insert_bid_rows_custom_headers(monkeypatch, tmp_path, caplog):
    from fm_tool_core import bid_utils

    wb_path = tmp_path / "wb.xlsx"
    wb_path.touch()
    calls: list[str] = []

    class FakeApi:
        def __init__(self):
            self.Rows = types.SimpleNamespace(Count=1)

        def Cells(self, _row, _col):
            end = lambda _dir: types.SimpleNamespace(Row=1)
            return types.SimpleNamespace(End=end)

    class FakeHeaderRange:
        def __init__(self, sheet):
            self.sheet = sheet

        def resize(self, _r, _c):
            return self

        def expand(self, _dir):
            return self

        @property
        def value(self):
            return (tuple(self.sheet.headers),)

        @value.setter
        def value(self, val):
            self.sheet.headers = val[0] if val and isinstance(val[0], list) else val

        def get_address(self, *_args):
            return "A1"

    class FakeDataRange:
        def __init__(self):
            self._value = None

        def resize(self, _r, _c):
            return self

        @property
        def value(self):
            return self._value

        @value.setter
        def value(self, val):
            calls.append("write")
            self._value = val

    class FakeCell:
        def __init__(self, addr):
            self.addr = addr

        def get_address(self, *_args):
            row, col = self.addr
            return f"{chr(ord('A') + col - 1)}{row}"

    class FakeSheet:
        def __init__(self):
            self.api = FakeApi()
            self.headers = bid_utils._COLUMNS.copy()
            self.headers[13] = " adhoc_info1 "
            self.headers[15] = "AdHoC_Info3  "

        def range(self, addr):
            if addr == (1, 1):
                return FakeHeaderRange(self)
            if addr[0] == 1:
                return FakeCell(addr)
            return FakeDataRange()

    sheet = FakeSheet()

    class FakeBook:
        def __init__(self):
            self.sheets = {"RFP": sheet}

        def save(self):
            pass

        def close(self):
            pass

    def fake_open(_path):
        return FakeBook()

    class FakeBooks:
        def open(self, path):
            return fake_open(path)

    class FakeApp:
        def __init__(self, *args, **kwargs):
            self.api = types.SimpleNamespace(DisplayAlerts=False)
            self.books = FakeBooks()

        def kill(self):
            pass

    monkeypatch.setattr(
        bid_utils,
        "xw",
        types.SimpleNamespace(App=FakeApp),
        raising=False,
    )
    monkeypatch.setattr(
        bid_utils,
        "pythoncom",
        types.SimpleNamespace(
            CoInitialize=lambda: None,
            CoUninitialize=lambda: None,
        ),
        raising=False,
    )

    rows = [
        {
            "LANE_ID": "1",
            "ORIG_POSTAL_CD": "11111",
            "DEST_POSTAL_CD": "22222",
        }
    ]
    log = logging.getLogger("test")
    caplog.set_level(logging.DEBUG)
    bid_utils.insert_bid_rows(
        wb_path,
        rows,
        log,
        adhoc_headers={" AdHoC inFo1 ": "X1", "adhocinfo3": "X3"},
    )
    assert calls == ["write"]
    assert sheet.headers[13] == "X1"
    assert sheet.headers[15] == "X3"
    assert "Received custom headers" in caplog.text
    assert "Replacing  adhoc_info1  with X1" in caplog.text
    assert "Replacing AdHoC_Info3   with X3" in caplog.text


def test_update_adhoc_headers(monkeypatch, tmp_path, caplog):
    from fm_tool_core import bid_utils

    wb_path = tmp_path / "wb.xlsx"
    wb_path.touch()

    class FakeHeaderRange:
        def __init__(self, sheet):
            self.sheet = sheet

        def resize(self, _r, _c):
            return self

        def expand(self, _dir):
            return self

        @property
        def value(self):
            return (tuple(self.sheet.headers),)

        @value.setter
        def value(self, val):
            self.sheet.headers = val[0] if val and isinstance(val[0], list) else val

        def get_address(self, *_args):
            return "A1"

    class FakeCell:
        def __init__(self, addr):
            self.addr = addr

        def get_address(self, *_args):
            row, col = self.addr
            return f"{chr(ord('A') + col - 1)}{row}"

    class FakeSheet:
        def __init__(self):
            self.headers = bid_utils._COLUMNS.copy()
            self.headers[13] = "adhoc_info1"
            self.headers[14] = "ADHOC_INFO2"

        def range(self, addr):
            if addr == (1, 1):
                return FakeHeaderRange(self)
            assert addr[0] == 1
            return FakeCell(addr)

    sheet = FakeSheet()

    class FakeBook:
        def __init__(self):
            self.sheets = {"RFP": sheet}

        def save(self):
            pass

        def close(self):
            pass

    def fake_open(_path):
        return FakeBook()

    class FakeBooks:
        def open(self, path):
            return fake_open(path)

    class FakeApp:
        def __init__(self, *args, **kwargs):
            self.api = types.SimpleNamespace(DisplayAlerts=False)
            self.books = FakeBooks()

        def kill(self):
            pass

    monkeypatch.setattr(
        bid_utils,
        "xw",
        types.SimpleNamespace(App=FakeApp),
        raising=False,
    )
    monkeypatch.setattr(
        bid_utils,
        "pythoncom",
        types.SimpleNamespace(
            CoInitialize=lambda: None,
            CoUninitialize=lambda: None,
        ),
        raising=False,
    )

    log = logging.getLogger("test")
    caplog.set_level(logging.DEBUG)
    bid_utils.update_adhoc_headers(
        wb_path,
        {"adhoc info1": "X1", "ADHOCINFO2": "X2", "ADHOCINFO11": "Z"},
        log,
    )
    assert sheet.headers[13] == "X1"
    assert sheet.headers[14] == "X2"
    assert "Received custom headers" in caplog.text
    assert "Examining N1: adhoc_info1" in caplog.text
    assert "Replacing adhoc_info1 with X1" in caplog.text
    assert "Examining O1: ADHOC_INFO2" in caplog.text
    assert "Replacing ADHOC_INFO2 with X2" in caplog.text
    assert "No matching column for custom header ADHOCINFO11" in caplog.text
