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
EVIDENCE_OUTCOMES = {"accepted", "rejected"}
DEFAULT_TIMEOUT_RETRIES = 2


def load_cases(path: Path) -> list[dict[str, Any]]:
    return load_cases_document(path)["cases"]


def load_cases_document(path: Path) -> dict[str, Any]:
    data = json.loads(path.read_text(encoding="utf-8-sig"))
    cases = data.get("cases")
    if not isinstance(cases, list):
        raise ValueError(f"{path} does not contain a cases array")
    return data


def expected_matches(expected: str | None, outcome: str) -> bool:
    if outcome not in EVIDENCE_OUTCOMES:
        return False
    return expected in (None, "", "observe") or expected == outcome


def is_oracle_infrastructure_failure(outcome: str) -> bool:
    return outcome not in EVIDENCE_OUTCOMES


def evidence_phase_for_case(case: dict[str, Any]) -> str:
    return "runtime" if str(case.get("mode", "compile")) == "run" else "compile"


def evidence_phase_for_result(case: dict[str, Any], result: dict[str, Any]) -> str:
    result_phase = str(result.get("evidencePhase") or "")
    if result_phase in ("compile", "runtime"):
        return result_phase
    stage = str(result.get("stage") or "")
    if stage in ("run", "runtime_dialog"):
        return "runtime"
    return evidence_phase_for_case(case)


def diagnostic_meaning_for_case(case: dict[str, Any], expected: str) -> str:
    if expected == "observe":
        return "observation"
    phase = evidence_phase_for_case(case)
    if phase == "runtime":
        return "runtime-error" if expected == "rejected" else "runtime-valid"
    return "compile-error" if expected == "rejected" else "compile-valid"


def diagnostic_meaning_for_result(
    case: dict[str, Any],
    result: dict[str, Any],
    expected: str,
) -> str:
    if expected == "observe":
        return "observation"
    phase = evidence_phase_for_result(case, result)
    if phase == "runtime":
        return "runtime-error" if expected == "rejected" else "runtime-valid"
    return "compile-error" if expected == "rejected" else "compile-valid"


def promote_observed_cases(
    document: dict[str, Any],
    results: list[dict[str, Any]],
) -> tuple[int, list[str]]:
    result_by_id = {
        str(result.get("caseId")): result
        for result in results
        if str(result.get("caseId") or "")
    }
    cases = document.get("cases", [])
    if not isinstance(cases, list):
        return 0, ["cases is not an array"]

    errors: list[str] = []
    cases_by_id: dict[str, dict[str, Any]] = {}
    for case in cases:
        case_id = str(case.get("id") or "")
        if case_id:
            cases_by_id[case_id] = case

    for case_id, result in result_by_id.items():
        case = cases_by_id.get(case_id)
        if case is None:
            errors.append(f"{case_id}: result has no matching fixture")
            continue
        expected = str(case.get("expected", "observe"))
        if expected != "observe":
            errors.append(f"{case_id}: expected is already {expected!r}")
            continue
        outcome = str(result.get("outcome", ""))
        if outcome not in ("accepted", "rejected"):
            errors.append(f"{case_id}: outcome {outcome!r} cannot be promoted")

    if errors:
        return 0, errors

    promoted = 0
    for case_id, result in result_by_id.items():
        case = cases_by_id[case_id]
        expected = str(result["outcome"])
        case["expected"] = expected
        case["provenance"] = "vbe-oracle-verified"
        case["evidencePhase"] = evidence_phase_for_result(case, result)
        case["diagnosticMeaning"] = diagnostic_meaning_for_result(case, result, expected)
        promoted += 1

    return promoted, []


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
    mode = str(case.get("mode", "compile"))
    is_compile_dialog = dialog.get("kind") == "vbe_compile_dialog"
    outcome = "rejected" if is_compile_dialog or mode == "run" else "accepted"
    stage = "compile_dialog" if is_compile_dialog else ("runtime_dialog" if mode == "run" else "vbe_dialog_after_compile")
    return {
        "caseId": case.get("id"),
        "outcome": outcome,
        "stage": stage,
        "message": message,
        "hresult": None,
        "dialog": dialog,
    }


def run_case(case: dict[str, Any], timeout: int) -> dict[str, Any]:
    return run_case_attempt(case, timeout, 0)


def case_with_mode(case: dict[str, Any], mode: str) -> dict[str, Any]:
    copy = dict(case)
    copy["mode"] = mode
    return copy


def run_case_attempt(
    case: dict[str, Any],
    timeout: int,
    dialog_hold_seconds: int,
) -> dict[str, Any]:
    mode = str(case.get("mode", "compile"))
    if mode != "compile_then_run":
        return run_case_once(case, timeout, dialog_hold_seconds)

    if not case.get("entryPoint"):
        return {
            "caseId": case.get("id"),
            "outcome": "worker_error",
            "stage": "setup",
            "message": "compile_then_run oracle cases require entryPoint.",
            "hresult": None,
        }

    compile_result = run_case_once(case_with_mode(case, "compile"), timeout, dialog_hold_seconds)
    if compile_result.get("outcome") != "accepted":
        compile_result["probeMode"] = "compile_then_run"
        compile_result["evidencePhase"] = "compile"
        return compile_result

    run_result = run_case_once(case_with_mode(case, "run"), timeout, dialog_hold_seconds)
    run_result["probeMode"] = "compile_then_run"
    run_result["compileResult"] = {
        "outcome": compile_result.get("outcome"),
        "stage": compile_result.get("stage"),
    }
    if run_result.get("outcome") in EVIDENCE_OUTCOMES:
        run_result["evidencePhase"] = "runtime"
    return run_result


def run_case_once(case: dict[str, Any], timeout: int, dialog_hold_seconds: int) -> dict[str, Any]:
    with tempfile.TemporaryDirectory(prefix="xlide-vbe-oracle-") as tmp:
        tmp_dir = Path(tmp)
        case_path = tmp_dir / "case.json"
        pid_path = tmp_dir / "excel.pid"
        stage_path = tmp_dir / "stage.txt"
        dialog_path = tmp_dir / "vbe_dialog.json"
        case_path.write_text(json.dumps(case), encoding="utf-8")

        if dialog_hold_seconds > 0:
            dialog_watch_seconds = max(1, timeout - dialog_hold_seconds - 5)
        else:
            dialog_watch_seconds = max(1, min(8, max(1, timeout - 3)))

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
            "-DialogHoldSeconds",
            str(dialog_hold_seconds),
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
            kill_recorded_excel(pid_path)
            return {
                "caseId": case.get("id"),
                "outcome": "timeout",
                "stage": stage,
                "message": f"Timed out after {timeout} seconds",
                "hresult": None,
            }

        try:
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
        finally:
            kill_recorded_excel(pid_path)


def run_case_with_retries(
    case: dict[str, Any],
    timeout: int,
    timeout_retries: int,
    dialog_hold_seconds: int,
) -> dict[str, Any]:
    attempts = max(1, timeout_retries + 1)
    timeout_results: list[dict[str, Any]] = []
    for attempt in range(1, attempts + 1):
        hold_seconds = dialog_hold_seconds if attempt == 1 else 0
        result = run_case_attempt(case, timeout, hold_seconds)
        result["attempt"] = attempt
        if result.get("outcome") != "timeout":
            if timeout_results:
                result["attempts"] = attempt
                result["previousTimeouts"] = timeout_results
            return result
        timeout_results.append(result)

    last = timeout_results[-1]
    return {
        "caseId": case.get("id"),
        "outcome": "oracle_failure",
        "stage": last.get("stage", "timeout"),
        "message": (
            f"Oracle timed out after {attempts} attempt(s) at {timeout} seconds each. "
            "Treat this as an oracle harness failure and investigate before running "
            "additional oracle cases."
        ),
        "hresult": None,
        "attempts": attempts,
        "timeoutSeconds": timeout,
        "previousTimeouts": timeout_results,
    }


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--cases", type=Path, default=DEFAULT_CASES)
    parser.add_argument("--case", dest="case_ids", action="append", default=[])
    parser.add_argument("--timeout", type=int, default=20)
    parser.add_argument(
        "--timeout-retries",
        type=int,
        default=DEFAULT_TIMEOUT_RETRIES,
        help=(
            "Retry a case this many times when the worker times out. Exhausted "
            "timeouts abort the oracle run as an infrastructure failure."
        ),
    )
    parser.add_argument(
        "--dialog-hold-seconds",
        type=int,
        default=0,
        help=(
            "Developer debugging aid: when a VBE dialog is detected, keep it "
            "visible for this many seconds before dismissing it. Applies only "
            "to the first attempt of a case."
        ),
    )
    parser.add_argument("--json", action="store_true", help="Emit machine-readable JSON")
    parser.add_argument(
        "--strict",
        action="store_true",
        help="Exit non-zero when observed outcomes do not match fixture expectations",
    )
    parser.add_argument(
        "--promote-observed",
        action="store_true",
        help=(
            "After running selected observe-only cases, write accepted/rejected "
            "oracle outcomes back as asserted vbe-oracle-verified expectations. "
            "Requires at least one --case."
        ),
    )
    args = parser.parse_args(argv)

    if args.timeout <= 0:
        print("--timeout must be greater than zero", file=sys.stderr)
        return 2
    if args.timeout_retries < 0:
        print("--timeout-retries cannot be negative", file=sys.stderr)
        return 2
    if args.dialog_hold_seconds < 0:
        print("--dialog-hold-seconds cannot be negative", file=sys.stderr)
        return 2
    if args.promote_observed and not args.case_ids:
        print("--promote-observed requires at least one --case", file=sys.stderr)
        return 2

    if os.name != "nt":
        print("Excel/VBE oracle tests require Windows.", file=sys.stderr)
        return 2

    document = load_cases_document(args.cases)
    cases = document["cases"]
    if args.case_ids:
        wanted = set(args.case_ids)
        cases = [case for case in cases if case.get("id") in wanted]
        missing = wanted.difference(str(case.get("id")) for case in cases)
        if missing:
            print(f"Unknown oracle case(s): {', '.join(sorted(missing))}", file=sys.stderr)
            return 2

    results: list[dict[str, Any]] = []
    failures = 0
    oracle_failures = 0
    for case in cases:
        result = run_case_with_retries(
            case,
            args.timeout,
            args.timeout_retries,
            args.dialog_hold_seconds,
        )
        expected = str(case.get("expected", "observe"))
        result["expected"] = expected
        result["matched"] = expected_matches(expected, str(result.get("outcome", "")))
        result["description"] = case.get("description", "")
        outcome = str(result.get("outcome", ""))
        if is_oracle_infrastructure_failure(outcome):
            oracle_failures += 1
        elif not result["matched"]:
            failures += 1
        results.append(result)
        if is_oracle_infrastructure_failure(outcome):
            break

    if args.promote_observed:
        if oracle_failures:
            print("Cannot promote while oracle infrastructure failures exist.", file=sys.stderr)
            return 1
        if failures:
            print("Cannot promote while expectation failures exist.", file=sys.stderr)
            return 1
        promoted, errors = promote_observed_cases(document, results)
        if errors:
            print("Cannot promote observed oracle result(s):", file=sys.stderr)
            for error in errors:
                print(f"  {error}", file=sys.stderr)
            return 1
        args.cases.write_text(json.dumps(document, indent=2) + "\n", encoding="utf-8")
        print(f"Promoted {promoted} oracle case(s) in {args.cases}", file=sys.stderr)

    if args.json:
        print(
            json.dumps(
                {
                    "results": results,
                    "failureCount": failures,
                    "oracleFailureCount": oracle_failures,
                },
                indent=2,
            )
        )
    else:
        for result in results:
            marker = "ORACLE-FAIL" if is_oracle_infrastructure_failure(str(result.get("outcome", ""))) else ("PASS" if result["matched"] else "FAIL")
            print(
                f"{marker} {result['caseId']}: outcome={result['outcome']} "
                f"expected={result['expected']} stage={result.get('stage', '')}"
            )
            if result.get("attempts"):
                print(f"  attempts={result['attempts']}")
            if result.get("message"):
                print(f"  {result['message']}")
        print(
            f"\n{len(results)} oracle case(s), {failures} expectation failure(s), "
            f"{oracle_failures} oracle infrastructure failure(s)."
        )

    return 1 if oracle_failures or (args.strict and failures) else 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
