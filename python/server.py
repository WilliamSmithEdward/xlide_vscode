"""
XLIDE Python backend - JSON-RPC 2.0 server over stdio.

Each request is a newline-terminated JSON object:
  {"jsonrpc":"2.0","id":1,"method":"listModules","params":{"path":"..."}}

Each response is a newline-terminated JSON object:
  {"jsonrpc":"2.0","id":1,"result":{...}}
  {"jsonrpc":"2.0","id":1,"error":{"code":-32000,"message":"..."}}
"""
from __future__ import annotations

import json
import sys
from typing import Any

from xlide.vba_io import (
    list_modules,
    list_subs,
    read_module,
    write_module,
    rename_module,
    delete_module,
    get_protection_info,
    validate_workbook,
    create_workbook,
)
from xlide.excel_io import get_workbook_info, list_sheets, read_cells, read_formulas, write_cells, run_openpyxl

_HANDLERS: dict[str, Any] = {
    "listModules": list_modules,
    "listSubs": list_subs,
    "readModule": read_module,
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


def _handle(req: dict[str, Any]) -> dict[str, Any]:
    rpc_id = req.get("id")
    method = req.get("method", "")
    params = req.get("params") or {}

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
            resp: dict[str, Any] = {
                "jsonrpc": "2.0",
                "id": None,
                "error": {"code": -32700, "message": f"Parse error: {exc}"},
            }
            sys.stdout.write(json.dumps(resp) + "\n")
            sys.stdout.flush()
            continue

        resp = _handle(req)
        sys.stdout.write(json.dumps(resp) + "\n")
        sys.stdout.flush()


if __name__ == "__main__":
    main()
