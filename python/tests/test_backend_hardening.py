"""Regression tests for the backend-robustness fixes (server loop + io)."""
import datetime
import json
from decimal import Decimal
from types import SimpleNamespace

import pytest

import server
from xlide.excel_io import _parse_cell
from xlide.vba_io import _signature_dropped


def test_json_default_serializes_dates_and_decimal():
    assert server._json_default(datetime.date(2024, 1, 2)) == "2024-01-02"
    assert server._json_default(Decimal("1.5")) == 1.5


def test_write_response_serializes_datetime_without_crashing(capsys):
    # The blocker: a date cell value in a result must serialize, not raise and
    # take the whole server loop down.
    server._write_response(
        {"jsonrpc": "2.0", "id": 1, "result": {"data": [[datetime.datetime(2024, 1, 2, 3, 4)]]}}
    )
    parsed = json.loads(capsys.readouterr().out.strip())
    assert parsed["id"] == 1
    assert parsed["result"]["data"][0][0].startswith("2024-01-02")


def test_write_response_falls_back_to_error_for_unserializable(capsys):
    recursive: list = []
    recursive.append(recursive)
    server._write_response({"jsonrpc": "2.0", "id": 7, "result": recursive})
    parsed = json.loads(capsys.readouterr().out.strip())
    assert parsed["id"] == 7
    assert parsed["error"]["code"] == -32603


def test_handle_rejects_non_dict_request():
    assert server._handle([1, 2, 3])["error"]["code"] == -32600


def test_handle_rejects_non_dict_params():
    resp = server._handle({"id": 3, "method": "listModules", "params": [1, 2]})
    assert resp["error"]["code"] == -32602
    assert resp["id"] == 3


def _warn(category, message):
    return SimpleNamespace(category=category, message=message)


def test_signature_dropped_matches_only_signature_warnings():
    assert _signature_dropped([_warn(UserWarning, "Dropped stale VBA digital signature")]) is True
    assert _signature_dropped([_warn(UserWarning, "unrelated deprecation notice")]) is False
    assert _signature_dropped([_warn(DeprecationWarning, "signature")]) is False
    assert _signature_dropped([]) is False


def test_parse_cell_raises_clear_error_on_bad_reference():
    with pytest.raises(ValueError, match="Invalid cell reference"):
        _parse_cell("not-a-cell")
