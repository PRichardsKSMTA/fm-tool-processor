import logging
import types

from fm_tool_core import process_fm_tool as mod


def _setup_pyodbc(monkeypatch, json_str):
    class Cur:
        def __enter__(self):
            return self

        def __exit__(self, *_):
            return False

        def execute(self, *_args, **_kwargs):
            pass

        def fetchone(self):
            return (json_str,)

    class Conn:
        def cursor(self):
            return Cur()

        def close(self):
            pass

    monkeypatch.setattr(
        mod,
        "pyodbc",
        types.SimpleNamespace(connect=lambda *a, **k: Conn()),
    )
    monkeypatch.setattr(mod, "_sql_conn_str", lambda: "conn")


def test_fetch_adhoc_headers_success(monkeypatch):
    _setup_pyodbc(monkeypatch, '{"ADHOC_INFO1": "A", "ADHOC_INFO2": "B"}')
    log = logging.getLogger("test")
    res = mod._fetch_adhoc_headers("guid", log)
    assert res == {"ADHOC_INFO1": "A", "ADHOC_INFO2": "B"}


def test_fetch_adhoc_headers_no_pyodbc(monkeypatch, caplog):
    monkeypatch.setattr(mod, "pyodbc", None)
    log = logging.getLogger("test")
    with caplog.at_level(logging.INFO):
        res = mod._fetch_adhoc_headers("guid", log)
    assert res == {}
    assert "SQL disabled" in caplog.text


def test_fetch_adhoc_headers_bad_json(monkeypatch, caplog):
    _setup_pyodbc(monkeypatch, "not-json")
    log = logging.getLogger("test")
    with caplog.at_level(logging.WARNING):
        res = mod._fetch_adhoc_headers("guid", log)
    assert res == {}
    assert "Malformed PROCESS_JSON" in caplog.text
