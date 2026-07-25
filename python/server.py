"""
XLIDE Python backend - JSON-RPC 2.0 server over stdio.

Each request is a newline-terminated JSON object:
  {"jsonrpc":"2.0","id":1,"method":"listModules","params":{"path":"..."}}

Each response is a newline-terminated JSON object:
  {"jsonrpc":"2.0","id":1,"result":{...}}
  {"jsonrpc":"2.0","id":1,"error":{"code":-32000,"message":"..."}}
"""
from __future__ import annotations

import datetime
import json
import sys
from decimal import Decimal
from typing import Any

from xlide.vba_io import (
    list_modules,
    list_subs,
    read_module,
    read_modules,
    write_module,
    rename_module,
    delete_module,
    get_protection_info,
    validate_workbook,
    create_workbook,
)
from xlide.excel_io import get_workbook_info, list_sheets, read_cells, read_formulas, write_cells, run_openpyxl
from xlide.env_info import get_package_versions

_HANDLERS: dict[str, Any] = {
    "listModules": list_modules,
    "getPackageVersions": get_package_versions,
    "listSubs": list_subs,
    "readModule": read_module,
    "readModules": read_modules,
    "writeModule": write_module,
    "renameModule": rename_module,
    "deleteModule": delete_module,
    "listSheets": list_sheets,
    "getWorkbookInfo": get_workbook_info,
    "getProtectionInfo": get_protection_info,
    "validateWorkbook": validate_workbook,
    "createWorkbook": create_workbook,
    "readCells": read_cells,
    "readFormulas": read_formulas,
    "writeCells": write_cells,
    "runOpenpyxl": run_openpyxl,
}


def _json_default(obj: Any) -> Any:
    """Fallback encoder for values json.dumps cannot serialize natively.

    openpyxl returns datetime/date/time and Decimal for ordinary cells, so
    without this a read of any date column would raise TypeError. Anything else
    unexpected degrades to its string form rather than crashing the server.
    """
    if isinstance(obj, (datetime.datetime, datetime.date, datetime.time)):
        return obj.isoformat()
    if isinstance(obj, Decimal):
        return float(obj)
    if isinstance(obj, (bytes, bytearray)):
        return obj.decode("utf-8", "replace")
    return str(obj)


def _write_response(resp: dict[str, Any]) -> None:
    """Serialize and emit one response line.

    A serialization failure becomes an error response for that request rather
    than an exception that would escape the read loop and take the whole backend
    down for every subsequent request.
    """
    try:
        line = json.dumps(resp, default=_json_default)
    except (TypeError, ValueError) as exc:
        line = json.dumps(
            {
                "jsonrpc": "2.0",
                "id": resp.get("id"),
                "error": {"code": -32603, "message": f"Result not serializable: {exc}"},
            }
        )
    sys.stdout.write(line + "\n")
    sys.stdout.flush()


def _handle(req: Any) -> dict[str, Any]:
    if not isinstance(req, dict):
        return {
            "jsonrpc": "2.0",
            "id": None,
            "error": {"code": -32600, "message": "Invalid Request: expected a JSON object"},
        }
    rpc_id = req.get("id")
    method = req.get("method", "")
    params = req.get("params")
    if params is None:
        params = {}
    elif not isinstance(params, dict):
        return {
            "jsonrpc": "2.0",
            "id": rpc_id,
            "error": {"code": -32602, "message": "Invalid params: expected an object"},
        }

    handler = _HANDLERS.get(method)
    if handler is None:
        return {
            "jsonrpc": "2.0",
            "id": rpc_id,
            "error": {"code": -32601, "message": f"Method not found: {method}"},
        }

    try:
        result = handler(**params)
        return {"jsonrpc": "2.0", "id": rpc_id, "result": result}
    except Exception as exc:  # noqa: BLE001
        return {
            "jsonrpc": "2.0",
            "id": rpc_id,
            "error": {"code": -32000, "message": str(exc)},
        }


def main() -> None:
    # Force UTF-8 on the stdio pipes so non-ASCII payloads (accented text, CJK,
    # emoji in cell data / VBA source / module names) round-trip correctly
    # regardless of the OS default encoding. Windows defaults to a legacy code
    # page, which otherwise silently mangles non-ASCII stdin.
    for _stream in (sys.stdin, sys.stdout):
        reconfigure = getattr(_stream, "reconfigure", None)
        if callable(reconfigure):
            try:
                reconfigure(encoding="utf-8")
            except (ValueError, OSError):
                pass

    # Signal to the TypeScript host that all imports are done and we are ready
    # to handle requests.  The bridge holds queued calls until it sees this.
    sys.stdout.write('{"ready":true}\n')
    sys.stdout.flush()

    for raw in sys.stdin:
        raw = raw.strip()
        if not raw:
            continue

        try:
            req = json.loads(raw)
        except json.JSONDecodeError as exc:
            _write_response(
                {
                    "jsonrpc": "2.0",
                    "id": None,
                    "error": {"code": -32700, "message": f"Parse error: {exc}"},
                }
            )
            continue

        try:
            resp = _handle(req)
        except Exception as exc:  # noqa: BLE001  - last resort: never let one request kill the loop
            rpc_id = req.get("id") if isinstance(req, dict) else None
            resp = {
                "jsonrpc": "2.0",
                "id": rpc_id,
                "error": {"code": -32603, "message": f"Internal error: {exc}"},
            }
        _write_response(resp)


if __name__ == "__main__":
    main()
