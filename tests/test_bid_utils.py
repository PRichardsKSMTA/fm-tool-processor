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
            return types.SimpleNamespace(
                End=lambda _dir: types.SimpleNamespace(Row=1),
            )

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

    class FakeApiRFP:
        def __init__(self):
            self.Rows = types.SimpleNamespace(Count=1)

        def Cells(self, _row, _col):
            return types.SimpleNamespace(
                End=lambda _dir: types.SimpleNamespace(Row=1),
            )

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

    class FakeSheetRFP:
        def __init__(self):
            self.api = FakeApiRFP()

        def range(self, _addr):
            return FakeDataRange()

    class FakeHeaderApi:
        def __init__(self, last_col):
            self.Columns = types.SimpleNamespace(Count=last_col)
            self._last_col = last_col

        def Cells(self, _row, _col):
            return types.SimpleNamespace(
                End=lambda _dir: types.SimpleNamespace(Column=self._last_col),
            )

    class FakeRow6Range:
        def __init__(self, sheet):
            self.sheet = sheet

        def resize(self, _r, _c):
            return self

        @property
        def value(self):
            return (tuple(self.sheet.row6),)

        def get_address(self, *_args):
            return "A6"

    class FakeRow7Cell:
        def __init__(self, sheet, col):
            self.sheet = sheet
            self.col = col

        @property
        def value(self):
            return self.sheet.row7[self.col - 1]

        @value.setter
        def value(self, val):
            self.sheet.row7[self.col - 1] = val

    class FakeCell:
        def __init__(self, addr):
            self.addr = addr

        def get_address(self, *_args):
            row, col = self.addr
            return f"{chr(ord('A') + col - 1)}{row}"

    class HeaderSheet:
        def __init__(self):
            self.row6 = ["", "Ad Hoc Info 1", "x", "Ad Hoc Info 3"]
            self.row7 = ["", "", "", ""]
            self.api = FakeHeaderApi(len(self.row6))

        def range(self, addr):
            row, col = addr
            if row == 6 and col == 1:
                return FakeRow6Range(self)
            if row == 6:
                return FakeCell(addr)
            if row == 7:
                return FakeRow7Cell(self, col)
            return FakeCell(addr)

    header_sheet = HeaderSheet()
    data_sheet = FakeSheetRFP()

    class FakeBook:
        def __init__(self):
            self.sheets = {"RFP": data_sheet, "BID": header_sheet}

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
        types.SimpleNamespace(
            App=FakeApp,
            constants=types.SimpleNamespace(
                Direction=types.SimpleNamespace(xlToLeft=1)
            ),
        ),
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
        adhoc_headers={
            " AdHoC inFo1 ": "X1",
            "adhocinfo3": "X3",
            "ADHOCINFO11": "Z",
        },
    )
    assert calls == ["write"]
    assert header_sheet.row7[1] == "X1"
    assert header_sheet.row7[3] == "X3"
    assert header_sheet.row7[0] == ""
    assert "Received custom headers" in caplog.text
    assert "Writing X1 to row 7 column 2" in caplog.text
    assert "Writing X3 to row 7 column 4" in caplog.text
    assert "No matching column for custom header ADHOCINFO11" in caplog.text


def test_update_adhoc_headers(monkeypatch, tmp_path, caplog):
    from fm_tool_core import bid_utils

    wb_path = tmp_path / "wb.xlsx"
    wb_path.touch()

    class FakeHeaderApi:
        def __init__(self, last_col):
            self.Columns = types.SimpleNamespace(Count=last_col)
            self._last_col = last_col

        def Cells(self, _row, _col):
            return types.SimpleNamespace(
                End=lambda _dir: types.SimpleNamespace(Column=self._last_col),
            )

    class FakeRow6Range:
        def __init__(self, sheet):
            self.sheet = sheet

        def resize(self, _r, _c):
            return self

        @property
        def value(self):
            return (tuple(self.sheet.row6),)

        def get_address(self, *_args):
            return "A6"

    class FakeRow7Cell:
        def __init__(self, sheet, col):
            self.sheet = sheet
            self.col = col

        @property
        def value(self):
            return self.sheet.row7[self.col - 1]

        @value.setter
        def value(self, val):
            self.sheet.row7[self.col - 1] = val

    class FakeCell:
        def __init__(self, addr):
            self.addr = addr

        def get_address(self, *_args):
            row, col = self.addr
            return f"{chr(ord('A') + col - 1)}{row}"

    class HeaderSheet:
        def __init__(self):
            self.row6 = ["Ad Hoc Info 1", "Ad Hoc Info 2", "foo"]
            self.row7 = ["", "", ""]
            self.api = FakeHeaderApi(len(self.row6))

        def range(self, addr):
            row, col = addr
            if row == 6 and col == 1:
                return FakeRow6Range(self)
            if row == 6:
                return FakeCell(addr)
            if row == 7:
                return FakeRow7Cell(self, col)
            return FakeCell(addr)

    sheet = HeaderSheet()

    class FakeBook:
        def __init__(self):
            self.sheets = {"BID": sheet}

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
        types.SimpleNamespace(
            App=FakeApp,
            constants=types.SimpleNamespace(
                Direction=types.SimpleNamespace(xlToLeft=1)
            ),
        ),
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
    assert sheet.row7[0] == "X1"
    assert sheet.row7[1] == "X2"
    assert sheet.row7[2] == ""
    assert "Received custom headers" in caplog.text
    assert "Writing X1 to row 7 column 1" in caplog.text
    assert "Writing X2 to row 7 column 2" in caplog.text
    assert "No matching column for custom header ADHOCINFO11" in caplog.text
