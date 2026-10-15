#!/usr/bin/env node
// Run optional Excel/VBE oracle fixtures.
//
// This coordinator keeps COM automation out of the toolchain by handing
// fixtures to a small PowerShell worker. It is intentionally not part of
// normal CI: Excel/VBE is a developer oracle for empirical compatibility
// checks only.

import { spawn, spawnSync } from 'node:child_process';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import readline from 'node:readline';
import { fileURLToPath } from 'node:url';

const ROOT = path.dirname(fileURLToPath(import.meta.url));
const DEFAULT_CASES = path.join(ROOT, 'vbe_oracle_cases.json');
const WORKER = path.join(ROOT, 'excel_vbe_oracle_worker.ps1');
const EVIDENCE_OUTCOMES = new Set(['accepted', 'rejected']);
const DEFAULT_TIMEOUT_RETRIES = 2;
const DEFAULT_CASES_PER_EXCEL = 40;
const RESULT_PREFIX = 'XLIDE-ORACLE-RESULT ';

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

// A rejection also has to come from the phase the fixture claims: a runtime
// fixture whose source does not compile proves nothing about the runtime.
function phaseMatches(kase, result) {
    const declared = String(kase.evidencePhase ?? '');
    if (String(kase.expected) !== 'rejected' || String(result.outcome) !== 'rejected'
        || (declared !== 'compile' && declared !== 'runtime')) {
        return true;
    }
    return evidencePhaseForResult(kase, result) === declared;
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
    // A run that stops at a compile error is compile evidence.
    if (stage === 'compile_dialog') { return 'compile'; }
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

function readRecordedPid(pidPath) {
    try {
        const pid = Number.parseInt(fs.readFileSync(pidPath, 'ascii').trim(), 10);
        return Number.isInteger(pid) ? pid : undefined;
    } catch {
        return undefined;
    }
}

// Only an Excel: a recorded pid may since have been reused by another process.
function killProcess(pid) {
    if (pid === undefined) { return; }
    spawnSync('taskkill', [
        '/PID', String(pid), '/FI', 'IMAGENAME eq EXCEL.EXE', '/T', '/F',
    ], { encoding: 'utf8', timeout: 10_000 });
}

function killRecordedExcel(pidPath) {
    killProcess(readRecordedPid(pidPath));
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

// One PowerShell worker serves the cases in turn and keeps its Excel between
// compile cases, each in a fresh workbook; the worker replaces Excel after
// every run case, every --cases-per-excel cases, and any case that is not
// evidence. The
// coordinator still owns each case's deadline: on a timeout it kills the
// worker and its Excel, and the next case starts a new worker.
class OracleWorker {
    constructor(casesPerExcel) {
        this.casesPerExcel = casesPerExcel;
        this.child = null;
        this.pending = null;
        this.stderr = '';
        // The Excel the worker last reported, kept past the case's temp files
        // so that close() can still end it.
        this.excelPid = undefined;
    }

    start() {
        const child = spawn('powershell.exe', [
            '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', WORKER,
            '-Serve', '-CasesPerExcel', String(this.casesPerExcel),
        ], { stdio: ['pipe', 'pipe', 'pipe'] });
        child.stdout.setEncoding('utf8');
        child.stderr.setEncoding('utf8');
        child.stdin.on('error', () => { /* reported through exit */ });
        child.stderr.on('data', (chunk) => {
            if (this.child === child) { this.stderr += chunk; }
        });
        readline.createInterface({ input: child.stdout }).on('line', (line) => {
            if (this.child === child && line.startsWith(RESULT_PREFIX)) {
                this.settle({ line: line.slice(RESULT_PREFIX.length) });
            }
        });
        const ended = (error, code) => {
            if (this.child !== child) { return; }
            this.child = null;
            // A worker that did not exit cleanly did not quit its Excel.
            if (error || code !== 0) { killProcess(this.excelPid); }
            this.settle({ exited: true, error, code });
        };
        child.on('error', (error) => ended(error, null));
        // 'close' comes after the last stdout line has been read.
        child.on('close', (code) => ended(null, code));
        this.child = child;
    }

    settle(reply) {
        const pending = this.pending;
        if (!pending) { return; }
        this.pending = null;
        // Includes anything written since the last result, so a crash between
        // cases is reported with the next one.
        const stderr = this.stderr.trim();
        this.stderr = '';
        pending({ ...reply, stderr });
    }

    request(message, timeoutMs) {
        if (!this.child) { this.start(); }
        return new Promise((resolve) => {
            const timer = setTimeout(() => this.settle({ timedOut: true }), timeoutMs);
            this.pending = (reply) => {
                clearTimeout(timer);
                resolve(reply);
            };
            // Escaped to ASCII so a path survives the console code page.
            const json = JSON.stringify(message).replace(
                /[\u0080-\uffff]/g,
                (ch) => `\\u${ch.charCodeAt(0).toString(16).padStart(4, '0')}`,
            );
            this.child.stdin.write(`${json}\n`);
        });
    }

    kill() {
        const child = this.child;
        this.child = null;
        if (child) {
            const killed = spawnSync('taskkill', ['/PID', String(child.pid), '/T', '/F'], { encoding: 'utf8', timeout: 10_000 });
            if (killed.status !== 0) { child.kill(); }
        }
    }

    // Ends the worker, which quits its Excel; kills it if it does not exit.
    close() {
        const child = this.child;
        if (!child) { return Promise.resolve(); }
        return new Promise((resolve) => {
            const timer = setTimeout(() => {
                this.kill();
                killProcess(this.excelPid);
                resolve();
            }, 30_000);
            child.once('close', () => { clearTimeout(timer); resolve(); });
            child.stdin.end();
        });
    }
}

async function runCaseOnce(worker, kase, timeoutSeconds, dialogHoldSeconds) {
    const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'xlide-vbe-oracle-'));
    const casePath = path.join(tmpDir, 'case.json');
    const pidPath = path.join(tmpDir, 'excel.pid');
    const stagePath = path.join(tmpDir, 'stage.txt');
    const dialogPath = path.join(tmpDir, 'vbe_dialog.json');
    fs.writeFileSync(casePath, JSON.stringify(kase), 'utf8');

    // The watch window is an upper bound, not a wait. A compile case ends when
    // its error dialog is recorded or the Compile command turns disabled, and a
    // run case ends shortly after Excel.Run returns. Only a case that gives
    // neither signal waits out the window, so excel_start + watch + cleanup
    // must still fit inside --timeout. Raising --timeout adds Excel-startup
    // headroom; the window itself is capped at 22 s, so for a longer one use
    // --dialog-hold-seconds.
    const dialogWatchSeconds = dialogHoldSeconds > 0
        ? Math.max(1, timeoutSeconds - dialogHoldSeconds - 5)
        : Math.max(12, Math.min(22, timeoutSeconds - 22));

    try {
        const reply = await worker.request({
            casePath, pidPath, stagePath, dialogPath, dialogWatchSeconds, dialogHoldSeconds,
        }, timeoutSeconds * 1000);
        worker.excelPid = readRecordedPid(pidPath) ?? worker.excelPid;

        if (reply.timedOut) {
            let stage = 'timeout';
            try { stage = fs.readFileSync(stagePath, 'ascii').trim() || 'timeout'; } catch { /* keep */ }
            worker.kill();
            killRecordedExcel(pidPath);
            return {
                caseId: kase.id,
                outcome: 'timeout',
                stage,
                message: `Timed out after ${timeoutSeconds} seconds`,
                hresult: null,
            };
        }

        if (reply.exited) {
            // This case's Excel, which the worker's exit handler may not have known.
            killProcess(worker.excelPid);
            return readDialogResult(kase, dialogPath) ?? {
                caseId: kase.id,
                outcome: 'worker_error',
                stage: 'worker',
                message: reply.stderr || reply.error?.message || `PowerShell exited ${reply.code}`,
                hresult: null,
            };
        }

        let result;
        try {
            result = JSON.parse(reply.line);
        } catch {
            return readDialogResult(kase, dialogPath) ?? {
                caseId: kase.id,
                outcome: 'worker_error',
                stage: 'worker',
                message: reply.line,
                stderr: reply.stderr,
                hresult: null,
            };
        }
        if (reply.stderr) { result.stderr = reply.stderr; }
        return result;
    } finally {
        fs.rmSync(tmpDir, { recursive: true, force: true });
    }
}

async function runCaseAttempt(worker, kase, timeoutSeconds, dialogHoldSeconds) {
    const mode = String(kase.mode ?? 'compile');
    if (mode !== 'compile_then_run') {
        return runCaseOnce(worker, kase, timeoutSeconds, dialogHoldSeconds);
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

    const compileResult = await runCaseOnce(worker, { ...kase, mode: 'compile' }, timeoutSeconds, dialogHoldSeconds);
    if (compileResult.outcome !== 'accepted') {
        compileResult.probeMode = 'compile_then_run';
        compileResult.evidencePhase = 'compile';
        return compileResult;
    }

    const runResult = await runCaseOnce(worker, { ...kase, mode: 'run' }, timeoutSeconds, dialogHoldSeconds);
    runResult.probeMode = 'compile_then_run';
    runResult.compileResult = { outcome: compileResult.outcome, stage: compileResult.stage };
    if (EVIDENCE_OUTCOMES.has(String(runResult.outcome))) {
        runResult.evidencePhase = evidencePhaseForResult({ ...kase, mode: 'run' }, runResult);
    }
    return runResult;
}

async function runCaseWithRetries(worker, kase, timeoutSeconds, timeoutRetries, dialogHoldSeconds) {
    const maxAttempts = Math.max(1, timeoutRetries + 1);
    const timeoutResults = [];
    for (let attempt = 1; attempt <= maxAttempts; attempt++) {
        const holdSeconds = attempt === 1 ? dialogHoldSeconds : 0;
        const result = await runCaseAttempt(worker, kase, timeoutSeconds, holdSeconds);
        result.attempt = attempt;
        if (result.outcome !== 'timeout') {
            if (timeoutResults.length) {
                result.attempts = attempt;
                result.previousTimeouts = timeoutResults;
            }
            return result;
        }
        timeoutResults.push(result);
        // A timeout still behind a compile-error dialog means the dialog could
        // not be dismissed (a run case's Excel is ended at one); another
        // attempt would meet the same dialog.
        if (result.stage === 'compile_dialog') { break; }
    }
    const attempts = timeoutResults.length;
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
  --cases-per-excel <n>       Start a new Excel after this many cases (default 40).
                              Every run case also gets a new one afterwards.
                              1 gives every case its own Excel.
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
        casesPerExcel: DEFAULT_CASES_PER_EXCEL,
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
            case '--cases-per-excel': args.casesPerExcel = number(argv[++i], flag); break;
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

async function main(argv) {
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
    if (args.casesPerExcel < 1) { console.error('--cases-per-excel must be at least 1'); return 2; }
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
    const worker = new OracleWorker(args.casesPerExcel);
    try {
        for (const kase of cases) {
            const result = await runCaseWithRetries(
                worker, kase, args.timeout, args.timeoutRetries, args.dialogHoldSeconds,
            );
            const expected = String(kase.expected ?? 'observe');
            result.expected = expected;
            result.matched = expectedMatches(expected, String(result.outcome ?? ''))
                && phaseMatches(kase, result);
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
    } finally {
        await worker.close();
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

process.exit(await main(process.argv.slice(2)));
