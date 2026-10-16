#!/usr/bin/env node
// Run optional VBE oracle fixtures in Excel, Word, PowerPoint or Access.
//
// This coordinator keeps COM automation out of the toolchain by handing
// fixtures to a small PowerShell worker. It is intentionally not part of
// normal CI: the Office VBE is a developer oracle for empirical
// compatibility checks only.

import { spawn, spawnSync } from 'node:child_process';
import crypto from 'node:crypto';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import readline from 'node:readline';
import { fileURLToPath } from 'node:url';

const ROOT = path.dirname(fileURLToPath(import.meta.url));
const DEFAULT_CASES = path.join(ROOT, 'vbe_oracle_cases.json');
const WORKER = path.join(ROOT, 'vbe_oracle_worker.ps1');
const EVIDENCE_OUTCOMES = new Set(['accepted', 'rejected']);
const DEFAULT_TIMEOUT_RETRIES = 2;
const DEFAULT_CASES_PER_INSTANCE = 40;
const RESULT_PREFIX = 'XLIDE-ORACLE-RESULT ';

// The Office hosts a fixture can run in, as src/officeHostApps.ts describes
// them (tests/vbeOracleWorker.test.ts keeps the two in step). A fixture names
// its host with `host`; one without runs in Excel.
const OFFICE_HOSTS = {
    excel: { progId: 'Excel.Application', image: 'EXCEL', noun: 'Excel' },
    word: { progId: 'Word.Application', image: 'WINWORD', noun: 'Word' },
    powerpoint: { progId: 'PowerPoint.Application', image: 'POWERPNT', noun: 'PowerPoint' },
    access: { progId: 'Access.Application', image: 'MSACCESS', noun: 'Access' },
};
const DEFAULT_HOST = 'excel';

function caseHost(kase) {
    return String(kase.host ?? DEFAULT_HOST).toLowerCase();
}

function comRegistered(progId) {
    const completed = spawnSync('powershell.exe', [
        '-NoProfile', '-ExecutionPolicy', 'Bypass', '-Command',
        `$type = [type]::GetTypeFromProgID("${progId}"); if ($null -eq $type) { exit 2 }`,
    ], { encoding: 'utf8', timeout: 5000 });
    return !completed.error && completed.status === 0;
}

// Resolve the COM server without starting the host. Use numeric file-version
// parts so registry lookup and worker reports do not depend on display strings.
function registeredHostBuild(progId) {
    const completed = spawnSync('powershell.exe', [
        '-NoProfile', '-ExecutionPolicy', 'Bypass', '-Command',
        `$ErrorActionPreference = 'Stop'
        try {
            $type = [type]::GetTypeFromProgID("${progId}")
            if ($null -eq $type) { exit 2 }
            $path = "CLSID\\" + $type.GUID.ToString('B') + "\\LocalServer32"
            $key = [Microsoft.Win32.Registry]::ClassesRoot.OpenSubKey($path)
            # 32-bit Office may register only in the 32-bit view.
            if ($null -eq $key) {
                $key = [Microsoft.Win32.RegistryKey]::OpenBaseKey('ClassesRoot', [Microsoft.Win32.RegistryView]::Registry32).OpenSubKey($path)
            }
            $command = [string]$key.GetValue('')
            if ($command -notmatch '^\\s*(?:"([^"]+[.]exe)"|(.+?[.]exe))(?:\\s|$)') { exit 2 }
            $exe = $matches[1]
            if (-not $exe) { $exe = $matches[2] }
            $v = [Diagnostics.FileVersionInfo]::GetVersionInfo($exe)
            if ($v.FileMajorPart -le 0) { exit 2 }
            '{0}.{1}.{2}.{3}' -f $v.FileMajorPart, $v.FileMinorPart, $v.FileBuildPart, $v.FilePrivatePart
        } catch { exit 2 }`,
    ], { encoding: 'utf8', timeout: 15_000 });
    const build = completed.stdout?.trim();
    if (!completed.error && completed.status === 0 && /^\d+\.\d+\.\d+\.\d+$/.test(build)) {
        return { build };
    }
    return { build: null, why: completed.error?.message || completed.stderr?.trim() || `PowerShell exited ${completed.status}` };
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

// A checkpoint file holds one JSON line per finished case: a hash of the
// case, this harness and the build of the case's host, and the case's
// result. A case whose id and hash match a line with evidence is not run
// again, so an aborted run can resume where it stopped; editing the case or
// the harness, or a new build of its host, runs it again.
const HARNESS_HASH = crypto.createHash('sha256')
    .update(fs.readFileSync(fileURLToPath(import.meta.url)))
    .update(fs.readFileSync(WORKER))
    .digest('hex');

function caseHash(kase, hostBuild) {
    return crypto.createHash('sha256').update(HARNESS_HASH).update(JSON.stringify([hostBuild, kase])).digest('hex');
}

function checkpointKey(caseId, hash) {
    return `${caseId}\n${hash}`;
}

function loadCheckpoint(file) {
    const done = new Map();
    let raw;
    try {
        raw = fs.readFileSync(file, 'utf8');
    } catch (err) {
        if (err.code === 'ENOENT') { return done; }
        throw err;
    }
    // End a line cut short by a killed run, so the next result starts afresh.
    if (raw && !raw.endsWith('\n')) {
        fs.appendFileSync(file, '\n', 'utf8');
    }
    for (const line of raw.split(/\r?\n/)) {
        if (!line.trim()) { continue; }
        let entry;
        try {
            entry = JSON.parse(line);
        } catch {
            continue; // a line cut short when a run was killed
        }
        const result = entry?.result;
        if (result && EVIDENCE_OUTCOMES.has(String(result.outcome))) {
            done.set(checkpointKey(String(result.caseId), String(entry.caseHash)), result);
        }
    }
    return done;
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

// The worker records "<pid> <image>" for the host instance it started.
function readRecordedProcess(pidPath) {
    try {
        const [rawPid, image] = fs.readFileSync(pidPath, 'ascii').trim().split(/\s+/);
        const pid = Number.parseInt(rawPid, 10);
        return Number.isInteger(pid) && /^[A-Z]+$/.test(image ?? '') ? { pid, image } : undefined;
    } catch {
        return undefined;
    }
}

// Only that host's image: a recorded pid may since have been reused by
// another process.
function killProcess(recorded) {
    if (recorded === undefined) { return; }
    spawnSync('taskkill', [
        '/PID', String(recorded.pid), '/FI', `IMAGENAME eq ${recorded.image}.EXE`, '/T', '/F',
    ], { encoding: 'utf8', timeout: 10_000 });
}

function killRecordedHost(pidPath) {
    killProcess(readRecordedProcess(pidPath));
}

function readDialogResult(kase, dialogPath) {
    let dialog;
    try {
        // The watcher writes the file with a UTF-8 byte order mark.
        dialog = JSON.parse(fs.readFileSync(dialogPath, 'utf8').replace(/^\uFEFF/, ''));
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

// One PowerShell worker serves the cases in turn and keeps its host instance
// between compile cases of one host, each in a fresh file; the worker
// replaces the instance after every run case, every --cases-per-instance
// cases, any case that is not evidence, and when the next case names another
// host. The coordinator still owns each case's deadline: on a timeout it
// kills the worker and its host instance, and the next case starts a new
// worker.
class OracleWorker {
    constructor(casesPerInstance) {
        this.casesPerInstance = casesPerInstance;
        this.child = null;
        this.pending = null;
        this.stderr = '';
        // The host process the worker last reported, kept past the case's
        // temp files so that close() can still end it.
        this.hostProcess = undefined;
        // Each host's registered build, for checkpoint runs.
        this.hostBuilds = null;
    }

    start() {
        const child = spawn('powershell.exe', [
            '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', WORKER,
            '-Serve', '-CasesPerInstance', String(this.casesPerInstance),
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
            // A worker that did not exit cleanly did not quit its host instance.
            if (error || code !== 0) { killProcess(this.hostProcess); }
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

    // Ends the worker, which quits its host instance; kills both if it does not exit.
    close() {
        const child = this.child;
        if (!child) { return Promise.resolve(); }
        return new Promise((resolve) => {
            const timer = setTimeout(() => {
                this.kill();
                killProcess(this.hostProcess);
                resolve();
            }, 30_000);
            child.once('close', () => { clearTimeout(timer); resolve(); });
            child.stdin.end();
        });
    }
}

async function runCaseOnce(worker, kase, timeoutSeconds, dialogHoldSeconds) {
    const result = await runCaseOnceUnchecked(worker, kase, timeoutSeconds, dialogHoldSeconds);
    const host = caseHost(kase);
    const expectedBuild = worker.hostBuilds?.get(host);
    if (expectedBuild && EVIDENCE_OUTCOMES.has(String(result.outcome))
        && result.hostBuild !== expectedBuild) {
        // A result without a build usually comes from a worker that died, and
        // its own stage and message say why; keep them.
        const noun = OFFICE_HOSTS[host].noun;
        return {
            ...result,
            outcome: 'worker_error',
            stage: result.hostBuild ? 'host_build' : result.stage,
            message: result.hostBuild
                ? `Expected ${noun} build ${expectedBuild}; this case ran on ${result.hostBuild}.`
                : `No ${noun} build was reported, so this result cannot be cached. ${result.message ?? ''}`.trim(),
        };
    }
    return result;
}

async function runCaseOnceUnchecked(worker, kase, timeoutSeconds, dialogHoldSeconds) {
    const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'xlide-vbe-oracle-'));
    const casePath = path.join(tmpDir, 'case.json');
    const pidPath = path.join(tmpDir, 'host.pid');
    const stagePath = path.join(tmpDir, 'stage.txt');
    const dialogPath = path.join(tmpDir, 'vbe_dialog.json');
    fs.writeFileSync(casePath, JSON.stringify(kase), 'utf8');

    // The watch window is an upper bound, not a wait. A compile case ends when
    // its error dialog is recorded or the Compile command turns disabled, and a
    // run case ends shortly after the host's Run returns. Only a case that gives
    // neither signal waits out the window, so host start + watch + cleanup
    // must still fit inside --timeout. Raising --timeout adds startup
    // headroom; the window itself is capped at 22 s, so for a longer one use
    // --dialog-hold-seconds.
    const dialogWatchSeconds = dialogHoldSeconds > 0
        ? Math.max(1, timeoutSeconds - dialogHoldSeconds - 5)
        : Math.max(12, Math.min(22, timeoutSeconds - 22));

    try {
        const host = caseHost(kase);
        const reply = await worker.request({
            casePath, pidPath, stagePath, dialogPath, dialogWatchSeconds, dialogHoldSeconds,
            host, hostProgId: OFFICE_HOSTS[host].progId, hostImage: OFFICE_HOSTS[host].image,
        }, timeoutSeconds * 1000);
        worker.hostProcess = readRecordedProcess(pidPath) ?? worker.hostProcess;

        if (reply.timedOut) {
            let stage = 'timeout';
            try { stage = fs.readFileSync(stagePath, 'ascii').trim() || 'timeout'; } catch { /* keep */ }
            worker.kill();
            killRecordedHost(pidPath);
            return {
                caseId: kase.id,
                outcome: 'timeout',
                stage,
                message: `Timed out after ${timeoutSeconds} seconds`,
                hresult: null,
            };
        }

        if (reply.exited) {
            // This case's host instance, which the worker's exit handler may not have known.
            killProcess(worker.hostProcess);
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
        // not be dismissed (a run case's host is ended at one); another
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

const USAGE = `Usage: node run_vbe_oracle.mjs [options]

  A fixture runs in the Office host its "host" names: excel (the default),
  word, powerpoint or access.

  --cases <path>              Fixture document (default: vbe_oracle_cases.json)
  --case <id>                 Run only this case; repeatable
  --timeout <seconds>         Per-case worker timeout (default 45). Must fit a cold
                              host start plus the dialog-watch window
                              (~min(22, timeout-22)s) plus COM cleanup.
  --timeout-retries <n>       Retry a timed-out case n times (default 2).
                              Exhausted timeouts abort as an infrastructure failure.
  --cases-per-instance <n>    Start a new host instance after this many cases
                              (default 40). Every run case, and every Access case,
                              also gets a new one afterwards; 1 gives every case
                              its own. --cases-per-excel is the old name.
  --dialog-hold-seconds <n>   Debugging aid: hold a detected VBE dialog visible
                              this long before dismissing (first attempt only).
  --checkpoint <path>         Append each case's result to this JSON-lines file,
                              and skip cases it already holds an accepted or
                              rejected result for, so a stopped run can resume.
                              Fixture, harness or host build changes rerun cases.
                              Requires a readable registered executable version
                              for each host the cases use; a missing or
                              mismatched build from the worker aborts the run.
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
        casesPerInstance: DEFAULT_CASES_PER_INSTANCE,
        dialogHoldSeconds: 0,
        checkpoint: null,
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
            case '--cases-per-instance':
            case '--cases-per-excel': args.casesPerInstance = number(argv[++i], flag); break;
            case '--dialog-hold-seconds': args.dialogHoldSeconds = number(argv[++i], flag); break;
            case '--checkpoint': args.checkpoint = argv[++i]; break;
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
    if (args.casesPerInstance < 1) { console.error('--cases-per-instance must be at least 1'); return 2; }
    if (args.promoteObserved && !args.caseIds.length) {
        console.error('--promote-observed requires at least one --case');
        return 2;
    }
    if (process.platform !== 'win32') {
        console.error('VBE oracle tests require Windows.');
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

    // Every host the selected cases use must be installed, and for a
    // checkpoint run its build must be readable.
    const unknownHosts = [...new Set(cases.map(caseHost))].filter((host) => !(host in OFFICE_HOSTS));
    if (unknownHosts.length) {
        console.error(`Unknown oracle host(s): ${unknownHosts.join(', ')}. Use ${Object.keys(OFFICE_HOSTS).join(', ')}.`);
        return 2;
    }
    const hosts = [...new Set(cases.map(caseHost))];
    for (const host of hosts) {
        if (!comRegistered(OFFICE_HOSTS[host].progId)) {
            console.error(`VBE oracle cases for ${host} require Microsoft ${OFFICE_HOSTS[host].noun} COM registration.`);
            return 2;
        }
    }
    const hostBuilds = new Map();
    if (args.checkpoint) {
        for (const host of hosts) {
            const registered = registeredHostBuild(OFFICE_HOSTS[host].progId);
            if (!registered.build) {
                console.error(`--checkpoint requires a readable file version for the registered ${OFFICE_HOSTS[host].image}.EXE: ${registered.why}`);
                return 2;
            }
            hostBuilds.set(host, registered.build);
        }
    }

    const results = [];
    let failures = 0;
    let oracleFailures = 0;
    const checkpoint = args.checkpoint ? loadCheckpoint(args.checkpoint) : null;
    const worker = new OracleWorker(args.casesPerInstance);
    worker.hostBuilds = args.checkpoint ? hostBuilds : null;
    try {
        for (const kase of cases) {
            const hash = caseHash(kase, hostBuilds.get(caseHost(kase)) ?? null);
            const saved = checkpoint?.get(checkpointKey(String(kase.id), hash));
            let result;
            if (saved) {
                result = { ...saved, resumed: true };
            } else {
                result = await runCaseWithRetries(
                    worker, kase, args.timeout, args.timeoutRetries, args.dialogHoldSeconds,
                );
                if (args.checkpoint) {
                    fs.appendFileSync(args.checkpoint, `${JSON.stringify({ caseHash: hash, result })}\n`, 'utf8');
                }
            }
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
