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


def test_insert_bid_rows_writes_rows(monkeypatch, tmp_path):
    from fm_tool_core import bid_utils

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
    bid_utils.insert_bid_rows(tmp_path / "wb.xlsx", rows, log)
    assert calls == ["write"]


def test_insert_bid_rows_custom_headers(monkeypatch, tmp_path):
    from fm_tool_core import bid_utils

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

        @property
        def value(self):
            return self.sheet.headers

        @value.setter
        def value(self, val):
            self.sheet.headers = val

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

    class FakeSheet:
        def __init__(self):
            self.api = FakeApi()
            self.headers = bid_utils._COLUMNS.copy()

        def range(self, addr):
            if addr == (1, 1):
                return FakeHeaderRange(self)
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
    bid_utils.insert_bid_rows(
        tmp_path / "wb.xlsx",
        rows,
        log,
        adhoc_headers={"ADHOC_INFO1": "X1", "ADHOC_INFO3": "X3"},
    )
    assert calls == ["write"]
    assert sheet.headers[13] == "X1"
    assert sheet.headers[15] == "X3"
