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
import { acquireOracleRunLock, describeRefusal, oracleRunLockAddress, oracleRunOwner } from './oracleRunLock.mjs';
import { describeLeftHosts, OwnedHosts, ownedHostsPath } from './oracleOwnedHosts.mjs';

const ROOT = path.dirname(fileURLToPath(import.meta.url));
const DEFAULT_CASES = path.join(ROOT, 'vbe_oracle_cases.json');
const WORKER = path.join(ROOT, 'vbe_oracle_worker.ps1');
const EVIDENCE_OUTCOMES = new Set(['accepted', 'rejected']);
const DEFAULT_TIMEOUT_RETRIES = 2;
const DEFAULT_CASES_PER_INSTANCE = 40;
const RESULT_PREFIX = 'XLIDE-ORACLE-RESULT ';

// The Office hosts a fixture can run in, as src/officeHostApps.ts describes
// them (tests/oracle/vbeOracleWorker.test.ts keeps the two in step). A fixture names
// its host with `host`; one without runs in Excel.
const OFFICE_HOSTS = {
    excel: { progId: 'Excel.Application', image: 'EXCEL', noun: 'Excel' },
    word: { progId: 'Word.Application', image: 'WINWORD', noun: 'Word' },
    powerpoint: { progId: 'PowerPoint.Application', image: 'POWERPNT', noun: 'PowerPoint' },
    access: { progId: 'Access.Application', image: 'MSACCESS', noun: 'Access' },
};
const DEFAULT_HOST = 'excel';

/** The case's host: absent is Excel, and anything else must be one of OFFICE_HOSTS, exactly. */
function caseHost(kase) {
    return kase.host === undefined ? DEFAULT_HOST : kase.host;
}

/**
 * Why a case's host is not one the oracle drives, or undefined. One rule for
 * every case and every check (issue #1440): `"Excel"`, `" excel"`, `null`
 * and `1` are refused, where the coordinator once lowercased some of them.
 */
function unsupportedHost(kase) {
    const host = kase.host;
    if (host === undefined || (typeof host === 'string' && Object.hasOwn(OFFICE_HOSTS, host))) {
        return undefined;
    }
    return `${String(kase.id)} names host ${JSON.stringify(host)}`;
}

// A case may reference type libraries, which the worker adds by GUID and
// version before its modules: `references: [{ guid, major, minor }]`.
function malformedReferences(kase) {
    const references = kase.references;
    if (references === undefined) {
        return undefined;
    }
    const guid = /^\{[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}\}$/;
    const valid = Array.isArray(references) && references.every((one) => one !== null && typeof one === 'object'
        && typeof one.guid === 'string' && guid.test(one.guid)
        && Number.isInteger(one.major) && one.major >= 0 && Number.isInteger(one.minor) && one.minor >= 0);
    return valid ? undefined : `${String(kase.id)} has references ${JSON.stringify(references)}`;
}

// Most rejected cases show a VBE error box, and Windows plays a message-box
// sound for each. On some machines each sound leaves thousands of kernel trace
// registrations that Windows frees only slowly, so a long run can fill
// gigabytes of kernel memory. While a run goes, it silences the four
// message-box sounds for its user (Windows falls back to the Default Beep for
// a silenced dialog sound, so all four are needed) and then puts back exactly
// what was there. The previous sounds are saved first, so a run killed before
// it could put them back leaves them for the next run to restore.
const DIALOG_SOUNDS = ['.Default', 'SystemAsterisk', 'SystemExclamation', 'SystemHand'];

function dialogSoundsPath() {
    let home = '';
    try { home = os.userInfo().homedir; } catch { /* no account record */ }
    return path.join(home || os.homedir(), '.xlide-vbe-oracle-sounds.json');
}

function dialogSoundsScript() {
    return String.raw`
    $file = $env:XLIDE_SOUNDS_FILE
    $owner = [int]$env:XLIDE_SOUNDS_OWNER
    $events = @(${DIALOG_SOUNDS.map((name) => `'${name}'`).join(', ')})
    $schemes = 'AppEvents\Schemes\Apps\.Default'
    function Open-Sound($name, $writable) {
        [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey((Join-Path (Join-Path $schemes $name) '.Current'), $writable)
    }
    # A normal shell cannot read an elevated run's StartTime, but CIM can read
    # its creation time, to the microsecond: both are compared to that.
    # Every process, elevated or not, is listed, so one not listed has ended.
    function Started($id) {
        $listed = Get-Process -Id $id -ErrorAction SilentlyContinue
        if ($null -eq $listed) { return '' }
        try { $time = $listed.StartTime.ToFileTimeUtc() }
        catch {
            # A query that fails is no proof the process has ended.
            try { $process = Get-CimInstance Win32_Process -Filter "ProcessId=$id" -ErrorAction Stop } catch { return 'unreadable' }
            if ($null -eq $process) { return '' }
            try { $time = $process.CreationDate.ToFileTimeUtc() } catch { return 'unreadable' }
        }
        return [string]($time - ($time % 10))
    }
    # Only a sound still silent is put back: one the user set to another
    # sound, or removed, during the run stays as the user left it.
    function Restore($saved) {
        foreach ($sound in @($saved.sounds)) {
            $key = Open-Sound $sound.name $true
            if ($null -eq $key) { continue }
            try {
                if ($key.GetValueNames() -contains '' -and -not [string]$key.GetValue('', $null, 'DoNotExpandEnvironmentNames')) {
                    $key.SetValue('', [string]$sound.value, [Microsoft.Win32.RegistryValueKind]$sound.kind)
                }
            } finally { $key.Close() }
        }
    }
    function Test-SavedSounds($saved) {
        if ($null -eq $saved -or
            ($saved.owner -isnot [int] -and $saved.owner -isnot [long]) -or
            $saved.owner -le 0 -or $saved.owner -gt [int]::MaxValue -or
            $saved.started -isnot [string] -or $saved.started -notmatch '^\d+$') { return $false }
        $entries = @($saved.sounds)
        if ($entries.Count -lt 1 -or $entries.Count -gt $events.Count) { return $false }
        $seen = @()
        foreach ($sound in $entries) {
            if ($null -eq $sound -or $sound.name -isnot [string] -or
                $events -notcontains $sound.name -or $seen -contains $sound.name -or
                $sound.kind -isnot [string] -or $sound.kind -notin @('String', 'ExpandString') -or
                $sound.value -isnot [string] -or [string]::IsNullOrEmpty($sound.value)) { return $false }
            $seen += $sound.name
        }
        return $true
    }
    if (Test-Path -LiteralPath $file) {
        $saved = $null
        try { $saved = Get-Content -Raw -Encoding UTF8 -LiteralPath $file | ConvertFrom-Json } catch { }
        # Validate the entire backup before the first restore write or deletion.
        if (-not (Test-SavedSounds $saved)) {
            # Unreadable: its sounds cannot be put back, so set it aside where
            # the user can find them, and go on.
            $aside = $file + '.unreadable-' + [DateTime]::UtcNow.Ticks
            [IO.File]::Move($file, $aside)
            [Console]::Error.WriteLine("The saved message-box sounds could not be read; they were moved to $aside.")
        } elseif ($env:XLIDE_SOUNDS_ACTION -eq 'restore') {
            if ($saved.owner -eq $owner -and $saved.started -eq (Started $owner)) {
                Restore $saved
                Remove-Item -Force -LiteralPath $file
            }
            exit 0
        } elseif ($saved.started -and (Started $saved.owner) -in @($saved.started, 'unreadable')) {
            # Another run that is still going silenced them and will put them
            # back; one that cannot be checked is left alone too.
            exit 0
        } else {
            # Left by a run that was killed.
            Restore $saved
            Remove-Item -Force -LiteralPath $file
        }
    }
    if ($env:XLIDE_SOUNDS_ACTION -eq 'restore') { exit 0 }
    $sounds = @()
    foreach ($name in $events) {
        $key = Open-Sound $name $false
        if ($null -eq $key) { continue }
        try {
            $value = [string]$key.GetValue('', $null, 'DoNotExpandEnvironmentNames')
            $kind = if ($value) { [string]$key.GetValueKind('') } else { '' }
            if ($value -and $kind -in @('String', 'ExpandString')) { $sounds += [pscustomobject]@{ name = $name; value = $value; kind = $kind } }
        } finally { $key.Close() }
    }
    if (-not $sounds.Count) { exit 0 }
    # Created only if absent, so of two runs starting together one saves and
    # silences, and the other leaves the sounds to it.
    $snapshot = [pscustomobject]@{ owner = $owner; started = (Started $owner); sounds = $sounds }
    if (-not (Test-SavedSounds $snapshot)) { throw 'Could not validate the sound recovery snapshot; sounds were not changed.' }
    $json = $snapshot | ConvertTo-Json -Depth 4
    try { $stream = [IO.File]::Open($file, 'CreateNew', 'Write', 'None') }
    catch [IO.IOException] {
        # ERROR_FILE_EXISTS: another run got there first. Anything else is a real failure.
        if ($_.Exception.HResult -eq -2147024816) { exit 0 }
        throw
    }
    try {
        $bytes = [Text.Encoding]::UTF8.GetBytes($json)
        $stream.Write($bytes, 0, $bytes.Length)
        $stream.Flush($true)
    } finally { $stream.Dispose() }
    foreach ($sound in $sounds) {
        $key = Open-Sound $sound.name $true
        if ($null -eq $key) { continue }
        try { $key.SetValue('', '', [Microsoft.Win32.RegistryValueKind]$sound.kind) } finally { $key.Close() }
    }
`;
}

// Never stops a run: a sound that cannot be changed only costs memory.
function setDialogSounds(action) {
    const completed = spawnSync('powershell.exe', [
        '-NoProfile', '-ExecutionPolicy', 'Bypass', '-Command', `$ErrorActionPreference = 'Stop'
    # One run at a time saves, repairs or puts back the account's sounds, so
    # two runs starting together cannot undo each other's saved copy. Windows
    # releases the mutex if this process ends while holding it. One mutex per
    # account, in every session, open to that account elevated or not.
    [Console]::OutputEncoding = [Text.Encoding]::UTF8
    $user = [Security.Principal.WindowsIdentity]::GetCurrent().User
    $access = New-Object Security.AccessControl.MutexSecurity
    $access.AddAccessRule((New-Object Security.AccessControl.MutexAccessRule($user, 'FullControl', 'Allow')))
    $created = $false
    $mutex = [Threading.Mutex]::new($false, "Global\\xlide-vbe-oracle-sounds-$($user.Value)", [ref]$created, $access)
    $owned = $false
    try { $owned = $mutex.WaitOne(10000) } catch [Threading.AbandonedMutexException] { $owned = $true }
    if (-not $owned) {
        [Console]::Error.WriteLine('another oracle run held the sounds for over ten seconds')
        exit 3
    }
    try {
    ${dialogSoundsScript()}
    } finally { try { $mutex.ReleaseMutex() } catch { } }`,
    ], {
        encoding: 'utf8', timeout: 15_000,
        env: { ...process.env, XLIDE_SOUNDS_FILE: dialogSoundsPath(), XLIDE_SOUNDS_OWNER: String(process.pid), XLIDE_SOUNDS_ACTION: action },
    });
    // A notice from a successful run, such as an unreadable file set aside.
    if (!completed.error && completed.status === 0 && completed.stderr?.trim()) { console.error(completed.stderr.trim()); }
    if (completed.error || completed.status !== 0) {
        const why = completed.error?.message || completed.stderr?.trim() || `PowerShell exited ${completed.status}`;
        console.error(action === 'silence'
            ? `Could not silence the message-box sounds for this run (${why}); rejected cases may play them.`
            : `Could not put back the message-box sounds (${why}); the next oracle run will.`);
    }
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
// One process serves both checks: the marker line, flushed after the null check,
// confirms registration even when the build probe times out. With no marker,
// registration falls back to comRegistered's 5 s, so a lookup that never
// completes takes up to 20 s.
function registeredHostBuild(progId) {
    const completed = spawnSync('powershell.exe', [
        '-NoProfile', '-ExecutionPolicy', 'Bypass', '-Command',
        `$ErrorActionPreference = 'Stop'
        try {
            $type = [type]::GetTypeFromProgID("${progId}")
            if ($null -eq $type) { exit 2 }
            [Console]::Out.WriteLine('XLIDE-ORACLE-REGISTERED')
            [Console]::Out.Flush()
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
    const lines = (completed.stdout ?? '').split(/\r?\n/);
    const trailing = lines.pop();
    const marker = 'XLIDE-ORACLE-REGISTERED';
    const confirmed = lines.includes(marker);
    const registered = confirmed || comRegistered(progId);
    const build = [...lines.filter((line) => line !== marker), trailing].join('\n').trim();
    if (confirmed && !completed.error && completed.status === 0 && /^\d+\.\d+\.\d+\.\d+$/.test(build)) {
        return { registered, build };
    }
    return { registered, build: null, why: completed.error?.message || completed.stderr?.trim() || `PowerShell exited ${completed.status}` };
}

function loadCasesDocument(file) {
    // utf-8-sig: PowerShell writes this file back with a BOM.
    const raw = fs.readFileSync(file, 'utf8').replace(/^﻿/, '');
    const data = JSON.parse(raw);
    if (!Array.isArray(data.cases)) {
        throw new Error(`${file} does not contain a cases array`);
    }
    validateCrashBuilds(data.cases);
    return data;
}

function validateCrashBuilds(cases) {
    for (const kase of cases) {
        if (!Object.hasOwn(kase, 'crashesHostBuilds')) { continue; }
        if (!Array.isArray(kase.crashesHostBuilds)
            || kase.crashesHostBuilds.some((build) => typeof build !== 'string'
                || !/^\d+\.\d+\.\d+\.\d+$/.test(build) || build.trim() !== build)) {
            throw new Error(`${kase.id}: crashesHostBuilds must be an array of exact four-part numeric host build strings`);
        }
    }
}

function hostsNeedingBuilds(cases, checkpoint) {
    return [...new Set(cases.filter((kase) => checkpoint
        || Object.hasOwn(kase, 'crashesHostBuilds')).map(caseHost))];
}

function knownHostCrashResult(kase, hostBuilds) {
    const build = hostBuilds.get(caseHost(kase));
    if (!build || !kase.crashesHostBuilds?.includes(build)) { return undefined; }
    return {
        caseId: kase.id,
        outcome: 'skipped',
        stage: 'known_host_crash',
        hostBuild: build,
        message: `Skipped known host crash on ${caseHost(kase)} build ${build}.`,
        hresult: null,
    };
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
    return outcome !== 'skipped' && !EVIDENCE_OUTCOMES.has(outcome);
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

// Written as the fixture file is kept: two-space JSON in printable ASCII,
// with DEL and every non-ASCII character escaped, so a no-break space or an
// accented name stays visible and a promotion changes only the cases it
// promotes.
function formatCasesDocument(document) {
    return `${JSON.stringify(document, null, 2).replace(
        /[\u007f-\uffff]/g,
        (ch) => `\\u${ch.charCodeAt(0).toString(16).toUpperCase().padStart(4, '0')}`,
    )}\n`;
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

// The worker records "<pid> <image> <start>" for the host instance it started,
// the start time as a UTC FileTime (issue #1439).
function readRecordedProcess(pidPath) {
    try {
        const [rawPid, image, start] = fs.readFileSync(pidPath, 'ascii').trim().split(/\s+/);
        const pid = Number.parseInt(rawPid, 10);
        return Number.isInteger(pid) && /^[A-Z]+$/.test(image ?? '') && /^\d+$/.test(start ?? '')
            ? { pid, image, start } : undefined;
    } catch {
        return undefined;
    }
}

/** A host an earlier worker started is still running and could not be ended. */
class HostCleanupError extends Error {
    constructor(left) {
        super(describeLeftHosts(left));
        this.left = left;
    }
}

// The temporary folder of each case under way, so that an interrupt can still
// read which host it started and remove the folder (issue #1439).
const ACTIVE_CASE_DIRS = new Set();

/** Records the hosts the cases under way started, and removes their folders. */
function abandonActiveCases(owned) {
    for (const dir of [...ACTIVE_CASE_DIRS]) {
        const recorded = readRecordedProcess(path.join(dir, 'host.pid'));
        if (recorded) { owned.add(recorded); }
        try { fs.rmSync(dir, { recursive: true, force: true }); } catch { /* reported by the next run's temp */ }
        ACTIVE_CASE_DIRS.delete(dir);
    }
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
    constructor(casesPerInstance, owned = new OwnedHosts()) {
        this.casesPerInstance = casesPerInstance;
        // Every host a case started, until each is known to be gone (issue #1439).
        this.owned = owned;
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
        // A new worker owns no host yet: any host an earlier one left is ended first.
        const left = this.owned.sweep();
        if (left.length) { throw new HostCleanupError(left); }
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
            if (error || code !== 0) { this.owned.sweep(); }
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
                this.owned.sweep();
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
    // Every host but the one this worker still drives must be gone before the
    // next case: one that cannot be ended stops the run (issue #1439).
    const left = worker.owned.sweep(worker.child ? worker.hostProcess : undefined);
    if (left.length) {
        return { caseId: kase.id, outcome: 'worker_error', stage: 'host_cleanup', message: describeLeftHosts(left), hresult: null };
    }
    const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'xlide-vbe-oracle-'));
    ACTIVE_CASE_DIRS.add(tmpDir);
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
        let reply;
        try {
            reply = await worker.request({
                casePath, pidPath, stagePath, dialogPath, dialogWatchSeconds, dialogHoldSeconds,
                host, hostProgId: OFFICE_HOSTS[host].progId, hostImage: OFFICE_HOSTS[host].image,
            }, timeoutSeconds * 1000);
        } catch (error) {
            if (!(error instanceof HostCleanupError)) { throw error; }
            return { caseId: kase.id, outcome: 'worker_error', stage: 'host_cleanup', message: error.message, hresult: null };
        }
        const recorded = readRecordedProcess(pidPath);
        if (recorded) {
            worker.owned.add(recorded);
            worker.hostProcess = recorded;
        }

        if (reply.timedOut) {
            let stage = 'timeout';
            try { stage = fs.readFileSync(stagePath, 'ascii').trim() || 'timeout'; } catch { /* keep */ }
            worker.kill();
            worker.owned.sweep();
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
            worker.owned.sweep();
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
        ACTIVE_CASE_DIRS.delete(tmpDir);
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

  One run at a time per Windows user: while another run holds the lock, a
  run starts nothing and exits with code 3, naming that run.
`;

/** Why a checkpoint cannot be kept at `file`, or undefined: it is a directory, or its directory is missing or read-only. */
function unwritableCheckpoint(file) {
    try {
        if (fs.statSync(file).isDirectory()) { return 'is a directory'; }
        fs.accessSync(file, fs.constants.W_OK);
        return undefined;
    } catch (err) {
        if (err.code !== 'ENOENT') { return `cannot be written (${err.code ?? err.message})`; }
    }
    const dir = path.dirname(path.resolve(file));
    try {
        if (!fs.statSync(dir).isDirectory()) { return `${dir} is not a directory`; }
        fs.accessSync(dir, fs.constants.W_OK);
        return undefined;
    } catch (err) {
        return err.code === 'ENOENT' ? `its directory ${dir} does not exist` : `its directory cannot be written (${err.code ?? err.message})`;
    }
}

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
    let i = 0;
    // An option's value is the next argument, which must exist and must not
    // be another option: `--checkpoint --strict` names no file (issue #1440).
    // So a value cannot begin with `--`; write `./--name` for such a path.
    const value = (flag) => {
        const raw = argv[i + 1];
        if (raw === undefined || raw === '' || raw.startsWith('--')) {
            throw new Error(raw === undefined || raw === '' ? `${flag} expects a value` : `${flag} expects a value, not the option '${raw}'`);
        }
        i += 1;
        return raw;
    };
    const number = (flag) => {
        const raw = value(flag);
        if (!/^-?\d+$/.test(raw)) { throw new Error(`${flag} expects an integer (got '${raw}')`); }
        return Number.parseInt(raw, 10);
    };
    for (; i < argv.length; i++) {
        const flag = argv[i];
        switch (flag) {
            case '--cases': args.cases = value(flag); break;
            case '--case': args.caseIds.push(value(flag)); break;
            case '--timeout': args.timeout = number(flag); break;
            case '--timeout-retries': args.timeoutRetries = number(flag); break;
            case '--cases-per-instance':
            case '--cases-per-excel': args.casesPerInstance = number(flag); break;
            case '--dialog-hold-seconds': args.dialogHoldSeconds = number(flag); break;
            case '--checkpoint': args.checkpoint = value(flag); break;
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

    let document;
    try {
        document = loadCasesDocument(args.cases);
    } catch (err) {
        console.error(err.message);
        return 2;
    }
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
    const unknownHosts = cases.map(unsupportedHost).filter((why) => why !== undefined);
    if (unknownHosts.length) {
        console.error(`Unknown oracle host: ${unknownHosts.slice(0, 5).join('; ')}${unknownHosts.length > 5 ? `; and ${unknownHosts.length - 5} more` : ''}. Use ${Object.keys(OFFICE_HOSTS).join(', ')}, or leave host out for excel.`);
        return 2;
    }
    const badReferences = cases.map(malformedReferences).filter((why) => why !== undefined);
    if (badReferences.length) {
        console.error(`Malformed oracle references: ${badReferences.slice(0, 5).join('; ')}${badReferences.length > 5 ? `; and ${badReferences.length - 5} more` : ''}. Each is { "guid": "{...}", "major": n, "minor": n }.`);
        return 2;
    }
    // A checkpoint is first written after a case has run, so a path that
    // cannot take it is refused now, before Office starts (issue #1440).
    if (args.checkpoint) {
        const why = unwritableCheckpoint(args.checkpoint);
        if (why) {
            console.error(`--checkpoint ${args.checkpoint}: ${why}`);
            return 2;
        }
    }
    const hosts = [...new Set(cases.map(caseHost))];
    const buildHosts = hostsNeedingBuilds(cases, args.checkpoint);
    const registrations = new Map();
    // Every host's probe runs before any build failure is reported.
    for (const host of hosts) {
        const progId = OFFICE_HOSTS[host].progId;
        const registered = buildHosts.includes(host)
            ? registeredHostBuild(progId) : { registered: comRegistered(progId) };
        registrations.set(host, registered);
        if (!registered.registered) {
            console.error(`VBE oracle cases for ${host} require Microsoft ${OFFICE_HOSTS[host].noun} COM registration.`);
            return 2;
        }
    }
    const hostBuilds = new Map();
    // A later host's missing registration must win over any build failure.
    for (const host of buildHosts) {
        const registered = registrations.get(host);
        if (!registered.build && args.checkpoint) {
            console.error(`--checkpoint requires a readable file version for the registered ${OFFICE_HOSTS[host].image}.EXE: ${registered.why}`);
            return 2;
        }
        if (registered.build) { hostBuilds.set(host, registered.build); }
    }

    // One run at a time per user, taken before any Office automation or sound
    // change (issue #1438). Ending this process, however it ends, frees it.
    const lock = await acquireOracleRunLock(oracleRunLockAddress(), oracleRunOwner(argv));
    if (!lock.held) {
        console.error(describeRefusal(lock));
        return 3;
    }

    // Hosts an earlier run started and was killed before ending are this run's
    // to end, before it starts any of its own (issue #1439).
    const owned = new OwnedHosts({ file: ownedHostsPath() });
    const inherited = owned.load();
    if (inherited.length) {
        const left = owned.sweep();
        if (left.length) {
            console.error(describeLeftHosts(left));
            return 2;
        }
        console.error(`Ended ${inherited.map((r) => `${r.image}.EXE pid ${r.pid}`).join(', ')}, left running by an earlier oracle run.`);
    }

    const results = [];
    let skippedCount = 0;
    let failures = 0;
    let oracleFailures = 0;
    const checkpoint = args.checkpoint ? loadCheckpoint(args.checkpoint) : null;
    const worker = new OracleWorker(args.casesPerInstance, owned);
    worker.hostBuilds = args.checkpoint ? hostBuilds : null;
    // Ctrl-C, closing the console and a normal end all put the sounds back,
    // unless Windows ends the process first; the next run then does.
    let restored = false;
    const restoreSounds = () => {
        if (restored) { return; }
        restored = true;
        setDialogSounds('restore');
    };
    // One cleanup for every way out but a normal end (issue #1439): the worker,
    // the hosts the run started, the cases' temporary folders and the sounds.
    // It is synchronous, so it finishes inside Windows' console-close deadline
    // for a run's few hosts, and running it twice does nothing more.
    let cleanedUp = false;
    const cleanUp = () => {
        if (cleanedUp) { return; }
        cleanedUp = true;
        worker.kill();
        abandonActiveCases(owned);
        const left = owned.sweep();
        if (left.length) { console.error(describeLeftHosts(left)); }
        restoreSounds();
    };
    const interrupt = (code) => () => { cleanUp(); process.exit(code); };
    const signals = [['SIGINT', interrupt(130)], ['SIGBREAK', interrupt(149)], ['SIGTERM', interrupt(143)], ['SIGHUP', interrupt(129)]];
    for (const [signal, handler] of signals) { process.once(signal, handler); }
    process.once('exit', cleanUp);
    setDialogSounds('silence');
    try {
        for (const kase of cases) {
            const hash = caseHash(kase, hostBuilds.get(caseHost(kase)) ?? null);
            const saved = checkpoint?.get(checkpointKey(String(kase.id), hash));
            let result = knownHostCrashResult(kase, hostBuilds);
            if (result) {
                skippedCount += 1;
            } else if (saved) {
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
            } else if (outcome !== 'skipped' && !result.matched) {
                failures += 1;
            }
            results.push(result);
            if (isOracleInfrastructureFailure(outcome)) { break; }
        }
    } finally {
        try { await worker.close(); }
        finally {
            // The worker quit its host; any it could not is ended now, or named.
            abandonActiveCases(owned);
            const left = owned.sweep();
            if (left.length) {
                console.error(describeLeftHosts(left));
                oracleFailures += 1;
            }
            cleanedUp = true;
            restoreSounds();
            for (const [signal, handler] of signals) { process.removeListener(signal, handler); }
            process.removeListener('exit', cleanUp);
        }
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
        fs.writeFileSync(args.cases, formatCasesDocument(document), 'utf8');
        console.error(`Promoted ${promoted} oracle case(s) in ${args.cases}`);
    }

    if (args.json) {
        console.log(JSON.stringify({
            results,
            failureCount: failures,
            oracleFailureCount: oracleFailures,
            skippedCount,
        }, null, 2));
    } else {
        for (const result of results) {
            const marker = isOracleInfrastructureFailure(String(result.outcome ?? ''))
                ? 'ORACLE-FAIL'
                : (result.outcome === 'skipped' ? 'SKIP' : (result.matched ? 'PASS' : 'FAIL'));
            console.log(
                `${marker} ${result.caseId}: outcome=${result.outcome} `
                + `expected=${result.expected} stage=${result.stage ?? ''}`,
            );
            if (result.attempts) { console.log(`  attempts=${result.attempts}`); }
            if (result.message) { console.log(`  ${result.message}`); }
        }
        console.log(
            `\n${results.length} oracle case(s), ${failures} expectation failure(s), `
            + `${oracleFailures} oracle infrastructure failure(s), ${skippedCount} skipped case(s).`,
        );
    }

    return oracleFailures || (args.strict && failures) ? 1 : 0;
}

process.exit(await main(process.argv.slice(2)));
