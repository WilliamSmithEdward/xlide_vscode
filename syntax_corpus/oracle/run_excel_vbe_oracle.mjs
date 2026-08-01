#!/usr/bin/env node
// Run optional Excel/VBE oracle fixtures.
//
// This coordinator keeps COM automation out of the toolchain by launching a
// small PowerShell worker for each fixture. It is intentionally not part of
// normal CI: Excel/VBE is a developer oracle for empirical compatibility
// checks only.

import { spawnSync } from 'node:child_process';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = path.dirname(fileURLToPath(import.meta.url));
const DEFAULT_CASES = path.join(ROOT, 'vbe_oracle_cases.json');
const WORKER = path.join(ROOT, 'excel_vbe_oracle_worker.ps1');
const EVIDENCE_OUTCOMES = new Set(['accepted', 'rejected']);
const DEFAULT_TIMEOUT_RETRIES = 2;

function excelComRegistered() {
    const completed = spawnSync('powershell.exe', [
        '-NoProfile', '-ExecutionPolicy', 'Bypass', '-Command',
        '$type = [type]::GetTypeFromProgID("Excel.Application"); if ($null -eq $type) { exit 2 }',
    ], { encoding: 'utf8', timeout: 5000 });
    return !completed.error && completed.status === 0;
}

function loadCasesDocument(file) {
    // utf-8-sig: PowerShell writes this file back with a BOM.
    const raw = fs.readFileSync(file, 'utf8').replace(/^﻿/, '');
    const data = JSON.parse(raw);
    if (!Array.isArray(data.cases)) {
        throw new Error(`${file} does not contain a cases array`);
    }
    return data;
}

function expectedMatches(expected, outcome) {
    if (!EVIDENCE_OUTCOMES.has(outcome)) { return false; }
    return expected === undefined || expected === null || expected === ''
        || expected === 'observe' || expected === outcome;
}

function isOracleInfrastructureFailure(outcome) {
    return !EVIDENCE_OUTCOMES.has(outcome);
}

function evidencePhaseForCase(kase) {
    return String(kase.mode ?? 'compile') === 'run' ? 'runtime' : 'compile';
}

function evidencePhaseForResult(kase, result) {
    const resultPhase = String(result.evidencePhase ?? '');
    if (resultPhase === 'compile' || resultPhase === 'runtime') { return resultPhase; }
    const stage = String(result.stage ?? '');
    if (stage === 'run' || stage === 'runtime_dialog') { return 'runtime'; }
    return evidencePhaseForCase(kase);
}

function diagnosticMeaningForResult(kase, result, expected) {
    if (expected === 'observe') { return 'observation'; }
    if (evidencePhaseForResult(kase, result) === 'runtime') {
        return expected === 'rejected' ? 'runtime-error' : 'runtime-valid';
    }
    return expected === 'rejected' ? 'compile-error' : 'compile-valid';
}

function promoteObservedCases(document, results) {
    const resultById = new Map();
    for (const result of results) {
        const id = String(result.caseId ?? '');
        if (id) { resultById.set(id, result); }
    }
    const cases = document.cases;
    if (!Array.isArray(cases)) { return { promoted: 0, errors: ['cases is not an array'] }; }

    const errors = [];
    const casesById = new Map();
    for (const kase of cases) {
        const id = String(kase.id ?? '');
        if (id) { casesById.set(id, kase); }
    }

    for (const [caseId, result] of resultById) {
        const kase = casesById.get(caseId);
        if (kase === undefined) {
            errors.push(`${caseId}: result has no matching fixture`);
            continue;
        }
        const expected = String(kase.expected ?? 'observe');
        if (expected !== 'observe') {
            errors.push(`${caseId}: expected is already '${expected}'`);
            continue;
        }
        const outcome = String(result.outcome ?? '');
        if (outcome !== 'accepted' && outcome !== 'rejected') {
            errors.push(`${caseId}: outcome '${outcome}' cannot be promoted`);
        }
    }
    if (errors.length) { return { promoted: 0, errors }; }

    let promoted = 0;
    for (const [caseId, result] of resultById) {
        const kase = casesById.get(caseId);
        const expected = String(result.outcome);
        kase.expected = expected;
        kase.provenance = 'vbe-oracle-verified';
        kase.evidencePhase = evidencePhaseForResult(kase, result);
        kase.diagnosticMeaning = diagnosticMeaningForResult(kase, result, expected);
        promoted += 1;
    }
    return { promoted, errors: [] };
}

function killRecordedExcel(pidPath) {
    let pid;
    try {
        const raw = fs.readFileSync(pidPath, 'ascii').trim();
        if (!raw) { return; }
        pid = Number.parseInt(raw, 10);
        if (!Number.isInteger(pid)) { return; }
    } catch {
        return;
    }
    spawnSync('taskkill', ['/PID', String(pid), '/T', '/F'], { encoding: 'utf8', timeout: 10_000 });
}

function readDialogResult(kase, dialogPath) {
    let dialog;
    try {
        dialog = JSON.parse(fs.readFileSync(dialogPath, 'utf8'));
    } catch {
        return undefined;
    }
    const message = String(dialog.message || 'VBE showed a compile error dialog.');
    const mode = String(kase.mode ?? 'compile');
    const isCompileDialog = dialog.kind === 'vbe_compile_dialog';
    return {
        caseId: kase.id,
        outcome: isCompileDialog || mode === 'run' ? 'rejected' : 'accepted',
        stage: isCompileDialog
            ? 'compile_dialog'
            : (mode === 'run' ? 'runtime_dialog' : 'vbe_dialog_after_compile'),
        message,
        hresult: null,
        dialog,
    };
}

function runCaseOnce(kase, timeoutSeconds, dialogHoldSeconds) {
    const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'xlide-vbe-oracle-'));
    const casePath = path.join(tmpDir, 'case.json');
    const pidPath = path.join(tmpDir, 'excel.pid');
    const stagePath = path.join(tmpDir, 'stage.txt');
    const dialogPath = path.join(tmpDir, 'vbe_dialog.json');
    fs.writeFileSync(casePath, JSON.stringify(kase), 'utf8');

    // The watcher must outlast the VBE compile-/runtime-dialog latency. That
    // latency is dominated by the FIRST compile in a freshly spawned Excel
    // (each case gets its own process); on a cold VBE it exceeds a fixed 8s
    // cap, which silently misrecorded rejects as "accepted".
    //
    // Two competing constraints set the window:
    //   * reject cases need watch >= the cold dialog latency (~20s here);
    //   * accept cases show NO dialog and wait the FULL window, so
    //     excel_start + watch + cleanup must fit inside --timeout or the
    //     coordinator kills the worker as a (false) timeout.
    // So reserve generous startup/cleanup headroom and cap the window: ~22s
    // reliably catches the dialog without making accept cases wait forever.
    // Raise --timeout for slower machines (adds headroom); for a window longer
    // than the cap, use --dialog-hold-seconds.
    const dialogWatchSeconds = dialogHoldSeconds > 0
        ? Math.max(1, timeoutSeconds - dialogHoldSeconds - 5)
        : Math.max(12, Math.min(22, timeoutSeconds - 22));

    try {
        const completed = spawnSync('powershell.exe', [
            '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', WORKER,
            '-CasePath', casePath,
            '-PidPath', pidPath,
            '-StagePath', stagePath,
            '-DialogPath', dialogPath,
            '-DialogWatchSeconds', String(dialogWatchSeconds),
            '-DialogHoldSeconds', String(dialogHoldSeconds),
        ], { encoding: 'utf8', timeout: timeoutSeconds * 1000, killSignal: 'SIGKILL' });

        if (completed.error && completed.error.code === 'ETIMEDOUT') {
            let stage = 'timeout';
            try { stage = fs.readFileSync(stagePath, 'ascii').trim() || 'timeout'; } catch { /* keep */ }
            killRecordedExcel(pidPath);
            return {
                caseId: kase.id,
                outcome: 'timeout',
                stage,
                message: `Timed out after ${timeoutSeconds} seconds`,
                hresult: null,
            };
        }

        const stdout = (completed.stdout ?? '').trim();
        const stderr = (completed.stderr ?? '').trim();
        if (!stdout) {
            return readDialogResult(kase, dialogPath) ?? {
                caseId: kase.id,
                outcome: 'worker_error',
                stage: 'worker',
                message: stderr || `PowerShell exited ${completed.status}`,
                hresult: null,
            };
        }
        let result;
        try {
            result = JSON.parse(stdout.split(/\r?\n/).at(-1));
        } catch {
            return readDialogResult(kase, dialogPath) ?? {
                caseId: kase.id,
                outcome: 'worker_error',
                stage: 'worker',
                message: stdout,
                stderr,
                hresult: null,
            };
        }
        if (stderr) { result.stderr = stderr; }
        return result;
    } finally {
        killRecordedExcel(pidPath);
        fs.rmSync(tmpDir, { recursive: true, force: true });
    }
}

function runCaseAttempt(kase, timeoutSeconds, dialogHoldSeconds) {
    const mode = String(kase.mode ?? 'compile');
    if (mode !== 'compile_then_run') {
        return runCaseOnce(kase, timeoutSeconds, dialogHoldSeconds);
    }
    if (!kase.entryPoint) {
        return {
            caseId: kase.id,
            outcome: 'worker_error',
            stage: 'setup',
            message: 'compile_then_run oracle cases require entryPoint.',
            hresult: null,
        };
    }

    const compileResult = runCaseOnce({ ...kase, mode: 'compile' }, timeoutSeconds, dialogHoldSeconds);
    if (compileResult.outcome !== 'accepted') {
        compileResult.probeMode = 'compile_then_run';
        compileResult.evidencePhase = 'compile';
        return compileResult;
    }

    const runResult = runCaseOnce({ ...kase, mode: 'run' }, timeoutSeconds, dialogHoldSeconds);
    runResult.probeMode = 'compile_then_run';
    runResult.compileResult = { outcome: compileResult.outcome, stage: compileResult.stage };
    if (EVIDENCE_OUTCOMES.has(String(runResult.outcome))) {
        runResult.evidencePhase = 'runtime';
    }
    return runResult;
}

function runCaseWithRetries(kase, timeoutSeconds, timeoutRetries, dialogHoldSeconds) {
    const attempts = Math.max(1, timeoutRetries + 1);
    const timeoutResults = [];
    for (let attempt = 1; attempt <= attempts; attempt++) {
        const holdSeconds = attempt === 1 ? dialogHoldSeconds : 0;
        const result = runCaseAttempt(kase, timeoutSeconds, holdSeconds);
        result.attempt = attempt;
        if (result.outcome !== 'timeout') {
            if (timeoutResults.length) {
                result.attempts = attempt;
                result.previousTimeouts = timeoutResults;
            }
            return result;
        }
        timeoutResults.push(result);
    }
    const last = timeoutResults.at(-1);
    return {
        caseId: kase.id,
        outcome: 'oracle_failure',
        stage: last.stage ?? 'timeout',
        message:
            `Oracle timed out after ${attempts} attempt(s) at ${timeoutSeconds} seconds each. `
            + 'Treat this as an oracle harness failure and investigate before running '
            + 'additional oracle cases.',
        hresult: null,
        attempts,
        timeoutSeconds,
        previousTimeouts: timeoutResults,
    };
}

const USAGE = `Usage: node run_excel_vbe_oracle.mjs [options]

  --cases <path>              Fixture document (default: vbe_oracle_cases.json)
  --case <id>                 Run only this case; repeatable
  --timeout <seconds>         Per-case worker timeout (default 45). Must fit cold
                              Excel startup plus the dialog-watch window
                              (~min(22, timeout-22)s) plus COM cleanup.
  --timeout-retries <n>       Retry a timed-out case n times (default 2).
                              Exhausted timeouts abort as an infrastructure failure.
  --dialog-hold-seconds <n>   Debugging aid: hold a detected VBE dialog visible
                              this long before dismissing (first attempt only).
  --json                      Emit machine-readable JSON
  --strict                    Exit non-zero when outcomes do not match fixtures
  --promote-observed          Write accepted/rejected outcomes back as asserted
                              vbe-oracle-verified expectations. Requires --case.
`;

function parseArgs(argv) {
    const args = {
        cases: DEFAULT_CASES,
        caseIds: [],
        timeout: 45,
        timeoutRetries: DEFAULT_TIMEOUT_RETRIES,
        dialogHoldSeconds: 0,
        json: false,
        strict: false,
        promoteObserved: false,
    };
    const number = (raw, flag) => {
        const value = Number.parseInt(raw, 10);
        if (!Number.isInteger(value)) { throw new Error(`${flag} expects an integer (got '${raw}')`); }
        return value;
    };
    for (let i = 0; i < argv.length; i++) {
        const flag = argv[i];
        switch (flag) {
            case '--cases': args.cases = argv[++i]; break;
            case '--case': args.caseIds.push(argv[++i]); break;
            case '--timeout': args.timeout = number(argv[++i], flag); break;
            case '--timeout-retries': args.timeoutRetries = number(argv[++i], flag); break;
            case '--dialog-hold-seconds': args.dialogHoldSeconds = number(argv[++i], flag); break;
            case '--json': args.json = true; break;
            case '--strict': args.strict = true; break;
            case '--promote-observed': args.promoteObserved = true; break;
            case '-h': case '--help': args.help = true; break;
            default: throw new Error(`Unknown argument: ${flag}`);
        }
    }
    return args;
}

function main(argv) {
    let args;
    try {
        args = parseArgs(argv);
    } catch (err) {
        console.error(err.message);
        console.error(USAGE);
        return 2;
    }
    if (args.help) { console.log(USAGE); return 0; }

    if (args.timeout <= 0) { console.error('--timeout must be greater than zero'); return 2; }
    if (args.timeoutRetries < 0) { console.error('--timeout-retries cannot be negative'); return 2; }
    if (args.dialogHoldSeconds < 0) { console.error('--dialog-hold-seconds cannot be negative'); return 2; }
    if (args.promoteObserved && !args.caseIds.length) {
        console.error('--promote-observed requires at least one --case');
        return 2;
    }
    if (process.platform !== 'win32') {
        console.error('Excel/VBE oracle tests require Windows.');
        return 2;
    }
    if (!excelComRegistered()) {
        console.error('Excel/VBE oracle tests require Microsoft Excel COM registration.');
        return 2;
    }

    const document = loadCasesDocument(args.cases);
    let cases = document.cases;
    if (args.caseIds.length) {
        const wanted = new Set(args.caseIds);
        cases = cases.filter((kase) => wanted.has(kase.id));
        const found = new Set(cases.map((kase) => String(kase.id)));
        const missing = [...wanted].filter((id) => !found.has(id)).sort();
        if (missing.length) {
            console.error(`Unknown oracle case(s): ${missing.join(', ')}`);
            return 2;
        }
    }

    const results = [];
    let failures = 0;
    let oracleFailures = 0;
    for (const kase of cases) {
        const result = runCaseWithRetries(
            kase, args.timeout, args.timeoutRetries, args.dialogHoldSeconds,
        );
        const expected = String(kase.expected ?? 'observe');
        result.expected = expected;
        result.matched = expectedMatches(expected, String(result.outcome ?? ''));
        result.description = kase.description ?? '';
        const outcome = String(result.outcome ?? '');
        if (isOracleInfrastructureFailure(outcome)) {
            oracleFailures += 1;
        } else if (!result.matched) {
            failures += 1;
        }
        results.push(result);
        if (isOracleInfrastructureFailure(outcome)) { break; }
    }

    if (args.promoteObserved) {
        if (oracleFailures) {
            console.error('Cannot promote while oracle infrastructure failures exist.');
            return 1;
        }
        if (failures) {
            console.error('Cannot promote while expectation failures exist.');
            return 1;
        }
        const { promoted, errors } = promoteObservedCases(document, results);
        if (errors.length) {
            console.error('Cannot promote observed oracle result(s):');
            for (const error of errors) { console.error(`  ${error}`); }
            return 1;
        }
        fs.writeFileSync(args.cases, `${JSON.stringify(document, null, 2)}\n`, 'utf8');
        console.error(`Promoted ${promoted} oracle case(s) in ${args.cases}`);
    }

    if (args.json) {
        console.log(JSON.stringify({
            results,
            failureCount: failures,
            oracleFailureCount: oracleFailures,
        }, null, 2));
    } else {
        for (const result of results) {
            const marker = isOracleInfrastructureFailure(String(result.outcome ?? ''))
                ? 'ORACLE-FAIL'
                : (result.matched ? 'PASS' : 'FAIL');
            console.log(
                `${marker} ${result.caseId}: outcome=${result.outcome} `
                + `expected=${result.expected} stage=${result.stage ?? ''}`,
            );
            if (result.attempts) { console.log(`  attempts=${result.attempts}`); }
            if (result.message) { console.log(`  ${result.message}`); }
        }
        console.log(
            `\n${results.length} oracle case(s), ${failures} expectation failure(s), `
            + `${oracleFailures} oracle infrastructure failure(s).`,
        );
    }

    return oracleFailures || (args.strict && failures) ? 1 : 0;
}

process.exit(main(process.argv.slice(2)));
