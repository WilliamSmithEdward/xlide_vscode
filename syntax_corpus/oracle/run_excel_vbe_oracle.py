"""Run optional Excel/VBE oracle fixtures.

This coordinator keeps COM automation out of Python dependencies by launching a
small PowerShell worker for each fixture. It is intentionally not part of normal
CI: Excel/VBE is a developer oracle for empirical compatibility checks only.
"""
from __future__ import annotations

import argparse
import json
import os
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import Any


ROOT = Path(__file__).resolve().parent
DEFAULT_CASES = ROOT / "vbe_oracle_cases.json"
WORKER = ROOT / "excel_vbe_oracle_worker.ps1"


def load_cases(path: Path) -> list[dict[str, Any]]:
    data = json.loads(path.read_text(encoding="utf-8-sig"))
    cases = data.get("cases")
    if not isinstance(cases, list):
        raise ValueError(f"{path} does not contain a cases array")
    return cases


def expected_matches(expected: str | None, outcome: str) -> bool:
    return expected in (None, "", "observe") or expected == outcome


def kill_recorded_excel(pid_path: Path) -> None:
    try:
        raw = pid_path.read_text(encoding="ascii").strip()
        if not raw:
            return
        pid = int(raw)
    except (OSError, ValueError):
        return
    try:
        subprocess.run(
            ["taskkill", "/PID", str(pid), "/T", "/F"],
            capture_output=True,
            text=True,
            timeout=10,
            check=False,
        )
    except (OSError, subprocess.TimeoutExpired):
        return


def read_dialog_result(case: dict[str, Any], dialog_path: Path) -> dict[str, Any] | None:
    try:
        raw = dialog_path.read_text(encoding="utf-8")
    except OSError:
        return None
    try:
        dialog = json.loads(raw)
    except json.JSONDecodeError:
        return None
    message = str(dialog.get("message") or "VBE showed a compile error dialog.")
    return {
        "caseId": case.get("id"),
        "outcome": "rejected" if dialog.get("kind") == "vbe_compile_dialog" else "accepted",
        "stage": "compile_dialog" if dialog.get("kind") == "vbe_compile_dialog" else "vbe_dialog_after_compile",
        "message": message,
        "hresult": None,
        "dialog": dialog,
    }


def stage_dialog_fallback(case: dict[str, Any], stage: str, timeout: int) -> dict[str, Any] | None:
    if stage == "compile_dialog":
        return {
            "caseId": case.get("id"),
            "outcome": "rejected",
            "stage": "compile_dialog",
            "message": f"VBE compile dialog observed; worker timed out after {timeout} seconds before returning dialog text",
            "hresult": None,
        }
    if stage == "vbe_dialog":
        return {
            "caseId": case.get("id"),
            "outcome": "accepted",
            "stage": "vbe_dialog_after_compile",
            "message": f"VBE non-compile dialog observed; worker timed out after {timeout} seconds before returning dialog text",
            "hresult": None,
        }
    return None


def run_case(case: dict[str, Any], timeout: int) -> dict[str, Any]:
    with tempfile.TemporaryDirectory(prefix="xlide-vbe-oracle-") as tmp:
        tmp_dir = Path(tmp)
        case_path = tmp_dir / "case.json"
        pid_path = tmp_dir / "excel.pid"
        stage_path = tmp_dir / "stage.txt"
        dialog_path = tmp_dir / "vbe_dialog.json"
        case_path.write_text(json.dumps(case), encoding="utf-8")

        dialog_watch_seconds = max(1, min(3, timeout - 3))

        cmd = [
            "powershell.exe",
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            str(WORKER),
            "-CasePath",
            str(case_path),
            "-PidPath",
            str(pid_path),
            "-StagePath",
            str(stage_path),
            "-DialogPath",
            str(dialog_path),
            "-DialogWatchSeconds",
            str(dialog_watch_seconds),
        ]
        try:
            completed = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                timeout=timeout,
                check=False,
            )
        except subprocess.TimeoutExpired:
            try:
                stage = stage_path.read_text(encoding="ascii").strip() or "timeout"
            except OSError:
                stage = "timeout"
            dialog_result = read_dialog_result(case, dialog_path)
            if dialog_result:
                kill_recorded_excel(pid_path)
                return dialog_result
            stage_result = stage_dialog_fallback(case, stage, timeout)
            if stage_result:
                kill_recorded_excel(pid_path)
                return stage_result
            kill_recorded_excel(pid_path)
            return {
                "caseId": case.get("id"),
                "outcome": "timeout",
                "stage": stage,
                "message": f"Timed out after {timeout} seconds",
                "hresult": None,
            }

        stdout = completed.stdout.strip()
        if not stdout:
            dialog_result = read_dialog_result(case, dialog_path)
            if dialog_result:
                return dialog_result
            return {
                "caseId": case.get("id"),
                "outcome": "worker_error",
                "stage": "worker",
                "message": completed.stderr.strip() or f"PowerShell exited {completed.returncode}",
                "hresult": None,
            }
        try:
            result = json.loads(stdout.splitlines()[-1])
        except json.JSONDecodeError:
            dialog_result = read_dialog_result(case, dialog_path)
            if dialog_result:
                return dialog_result
            return {
                "caseId": case.get("id"),
                "outcome": "worker_error",
                "stage": "worker",
                "message": stdout,
                "stderr": completed.stderr.strip(),
                "hresult": None,
            }
        if completed.stderr.strip():
            result["stderr"] = completed.stderr.strip()
        return result


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--cases", type=Path, default=DEFAULT_CASES)
    parser.add_argument("--case", dest="case_ids", action="append", default=[])
    parser.add_argument("--timeout", type=int, default=20)
    parser.add_argument("--json", action="store_true", help="Emit machine-readable JSON")
    parser.add_argument(
        "--strict",
        action="store_true",
        help="Exit non-zero when observed outcomes do not match fixture expectations",
    )
    args = parser.parse_args(argv)

    if os.name != "nt":
        print("Excel/VBE oracle tests require Windows.", file=sys.stderr)
        return 2

    cases = load_cases(args.cases)
    if args.case_ids:
        wanted = set(args.case_ids)
        cases = [case for case in cases if case.get("id") in wanted]
        missing = wanted.difference(str(case.get("id")) for case in cases)
        if missing:
            print(f"Unknown oracle case(s): {', '.join(sorted(missing))}", file=sys.stderr)
            return 2

    results: list[dict[str, Any]] = []
    failures = 0
    for case in cases:
        result = run_case(case, args.timeout)
        expected = str(case.get("expected", "observe"))
        result["expected"] = expected
        result["matched"] = expected_matches(expected, str(result.get("outcome", "")))
        result["description"] = case.get("description", "")
        if not result["matched"]:
            failures += 1
        results.append(result)

    if args.json:
        print(json.dumps({"results": results, "failureCount": failures}, indent=2))
    else:
        for result in results:
            marker = "PASS" if result["matched"] else "FAIL"
            print(
                f"{marker} {result['caseId']}: outcome={result['outcome']} "
                f"expected={result['expected']} stage={result.get('stage', '')}"
            )
            if result.get("message"):
                print(f"  {result['message']}")
        print(f"\n{len(results)} oracle case(s), {failures} expectation failure(s).")

    return 1 if args.strict and failures else 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
