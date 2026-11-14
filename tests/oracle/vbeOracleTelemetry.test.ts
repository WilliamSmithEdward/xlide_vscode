import { spawnSync } from 'node:child_process';
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { EventEmitter } from 'node:events';
import { runInNewContext } from 'node:vm';
import { describe, expect, it } from 'vitest';

// A served run turns Office's older client telemetry off for the user and
// puts the setting back when the worker ends (vbe_oracle_worker.ps1).
const worker = readFileSync(join(process.cwd(), 'syntax_corpus/oracle/vbe_oracle_worker.ps1'), 'utf8').replaceAll('\r\n', '\n');
const blockStart = worker.indexOf('$TelemetryKey = ');
const blockEnd = worker.indexOf('\n}\n', worker.indexOf('function Disable-OracleTelemetry')) + 3;
const block = worker.slice(blockStart, blockEnd);
const serve = worker.slice(worker.indexOf('\nif ($Serve) {'), worker.indexOf('\nelse {', worker.indexOf('\nif ($Serve) {')));

describe('VBE oracle worker Office telemetry', () => {
    it('uses the user\'s own Office setting, not a policy', () => {
        expect(block).toContain("$TelemetryKey = 'Software\\Microsoft\\Office\\Common\\ClientTelemetry'");
        expect(block).not.toMatch(/Policies|LocalMachine/);
        expect(block.match(/\[Microsoft\.Win32\.Registry\]::CurrentUser/g)).toHaveLength(2);
        // A key it created is never deleted: that could take another program's value with it.
        expect(block).not.toContain('DeleteSubKey');
    });

    it('turns telemetry off before the first case and puts it back after the host is stopped', () => {
        const disable = serve.indexOf('try { Disable-OracleTelemetry } catch {');
        const loop = serve.indexOf('while ($null -ne ($line = [Console]::In.ReadLine()))');
        const stop = serve.indexOf('try { Stop-OracleHost }');
        const restore = serve.indexOf('try { Restore-OracleTelemetry } catch {');
        expect(disable).toBeGreaterThan(0);
        expect(loop).toBeGreaterThan(disable);
        expect(stop).toBeGreaterThan(loop);
        // In a finally of its own, so a host that would not exit still lets it run.
        expect(serve.slice(stop, restore)).toMatch(/^try \{ Stop-OracleHost \}\n\s*finally \{\n/);
        // A failure is reported and never stops the run. What turning it off
        // wrote is kept and written at the end, which the coordinator prints
        // (run_vbe_oracle.mjs close(), below); now it would only reach the
        // first case's result.
        const capture = serve.indexOf('[Console]::SetError($telemetryNotices)');
        expect(capture).toBeGreaterThan(0);
        expect(capture).toBeLessThan(disable);
        expect(serve.indexOf('finally { [Console]::SetError($stderr) }')).toBeGreaterThan(disable);
        expect(serve.indexOf('finally { [Console]::SetError($stderr) }')).toBeLessThan(loop);
        const written = serve.indexOf('if ($notices) { [Console]::Error.WriteLine($notices) }');
        expect(written).toBeGreaterThan(stop);
        expect(written).toBeLessThan(restore);
        expect(serve).toContain('[Console]::Error.WriteLine("Could not turn off Office telemetry for this run: $($_.Exception.Message)")');
        expect(serve).toContain('[Console]::Error.WriteLine("Could not put back the Office telemetry setting: $($_.Exception.Message) The next oracle run tries again.")');
    });

    it('saves what was there before changing it, and puts back only its own value', () => {
        const disable = block.slice(block.indexOf('function Disable-OracleTelemetry'));
        expect(disable.indexOf('Restore-OracleTelemetry')).toBeLessThan(disable.indexOf('$stream.Write('));
        expect(disable.indexOf('$stream.Flush($true)')).toBeGreaterThan(0);
        expect(disable.indexOf('$stream.Flush($true)')).toBeLessThan(disable.indexOf("$key.SetValue('DisableTelemetry', 1, 'DWord')"));
        const restore = block.slice(block.indexOf('function Restore-OracleTelemetry'), block.indexOf('function Disable-OracleTelemetry'));
        expect(restore).toContain("$key.GetValueKind('DisableTelemetry') -eq 'DWord' -and $key.GetValue('DisableTelemetry') -eq 1");
        expect(restore.indexOf('Remove-Item -Force -LiteralPath $TelemetryBackup')).toBeGreaterThan(restore.indexOf('$key.Close()'));
    });
});

// The coordinator shows what the worker wrote after its last case, which no
// case result carries: the real close(), with a fake worker process.
describe('VBE oracle coordinator at the end of a run', () => {
    const source = readFileSync(join(process.cwd(), 'syntax_corpus/oracle/run_vbe_oracle.mjs'), 'utf8');
    const start = source.indexOf('class OracleWorker {');
    const classSource = source.slice(start, source.indexOf('\n}', start) + 2);

    it.each([['a notice', 'Could not put back the Office telemetry setting: denied. The next oracle run tries again.\r\n', true],
        ['nothing', '', false]] as const)('prints %s the worker left on stderr, once', async (_name, text, printed) => {
        const errors: string[] = [];
        const OracleWorker = runInNewContext(`${classSource}\nOracleWorker`, {
            console: { error: (line: string) => errors.push(line) }, setTimeout, clearTimeout,
        });
        const worker = new OracleWorker(40, { sweep: () => [] });
        const child = Object.assign(new EventEmitter(), { stdin: { end: () => setImmediate(() => child.emit('close', 0)) } });
        worker.child = child;
        worker.stderr = text;
        await worker.close();
        expect(errors).toEqual(printed ? [text.trim()] : []);
        expect(worker.stderr).toBe('');
    });
});

type Value = { kind: string; value: string | number };
type Step = 'disable' | 'restore' | 'user-sets-2' | 'fail-writes';
type Scenario = { name: string; key?: Record<string, Value>; backup?: string; noBackupFolder?: boolean; noProfile?: boolean; steps: Step[] };
type Snapshot = { step: Step; key: Record<string, Value> | null; backup: string | null; error: string | null };
type Outcome = { name: string; snapshots: Snapshot[]; aside: number; stderr: string };

// The functions themselves, on Windows, with the registry replaced by an
// in-memory fake and the backup file in a temporary folder. One start.
describe.skipIf(process.platform !== 'win32')('executed VBE oracle worker telemetry setting', () => {
    const big = '9007199254740993'; // over 2^53: a double would lose it
    const scenarios: Scenario[] = [
        { name: 'no key', steps: ['disable', 'restore'] },
        { name: 'key without the value', key: { Other: { kind: 'String', value: 'x' } }, steps: ['disable', 'restore'] },
        { name: 'DWord 0', key: { DisableTelemetry: { kind: 'DWord', value: 0 } }, steps: ['disable', 'restore'] },
        { name: 'already DWord 1', key: { DisableTelemetry: { kind: 'DWord', value: 1 } }, steps: ['disable', 'restore'] },
        { name: 'String', key: { DisableTelemetry: { kind: 'String', value: 'abc' } }, steps: ['disable', 'restore'] },
        { name: 'ExpandString', key: { DisableTelemetry: { kind: 'ExpandString', value: '%SystemRoot%\\x' } }, steps: ['disable', 'restore'] },
        { name: 'large QWord', key: { DisableTelemetry: { kind: 'QWord', value: big } }, steps: ['disable', 'restore'] },
        { name: 'Binary', key: { DisableTelemetry: { kind: 'Binary', value: 'bytes' } }, steps: ['disable', 'restore'] },
        // The user's value differs from the saved one, so a restore that ignored it would show.
        { name: 'changed by the user during the run', key: { DisableTelemetry: { kind: 'DWord', value: 0 } }, steps: ['disable', 'user-sets-2', 'restore'] },
        { name: 'a killed run left it set', key: { DisableTelemetry: { kind: 'DWord', value: 1 } },
            backup: '{"present":true,"kind":"DWord","value":0}', steps: ['disable', 'restore'] },
        { name: 'an unreadable backup', key: { DisableTelemetry: { kind: 'DWord', value: 1 } }, backup: 'not json', steps: ['disable'] },
        { name: 'a backup whose value does not fit its kind', key: { DisableTelemetry: { kind: 'DWord', value: 1 } },
            backup: '{"present":true,"kind":"DWord","value":"abc"}', steps: ['disable'] },
        { name: 'the backup cannot be written', key: { DisableTelemetry: { kind: 'DWord', value: 0 } }, noBackupFolder: true, steps: ['disable'] },
        { name: 'the registry refuses the restore', key: { DisableTelemetry: { kind: 'DWord', value: 0 } }, steps: ['disable', 'fail-writes', 'restore'] },
        { name: 'no user profile folder', key: { DisableTelemetry: { kind: 'DWord', value: 0 } }, noProfile: true, steps: ['disable', 'restore'] },
    ];
    // Values are converted by kind, as RegistryKey.SetValue does, and string
    // values are expanded unless DoNotExpandEnvironmentNames is asked for.
    const fakeRegistry = String.raw`
function Convert-FakeValue($v, $k) {
    switch ([string]$k) {
        'DWord' { return [Convert]::ToInt32($v) }
        'QWord' { return [Convert]::ToInt64($v) }
        'String' { if ($v -isnot [string]) { throw "not a string" }; return $v }
        'ExpandString' { if ($v -isnot [string]) { throw "not a string" }; return $v }
    }
    return $v
}
function New-FakeKey($path, $writable) {
    $key = [pscustomobject]@{ Path = $path; Writable = [bool]$writable }
    $key | Add-Member ScriptMethod GetValueNames { return ,@($global:store[$this.Path].Keys) }
    $key | Add-Member ScriptMethod GetValueKind { param($n); return $global:store[$this.Path][$n].kind }
    $key | Add-Member ScriptMethod GetValue { param($n, $d, $o)
        $v = $global:store[$this.Path][$n]
        if ($null -eq $v) { return $d }
        if ($v.kind -eq 'ExpandString' -and [string]$o -ne 'DoNotExpandEnvironmentNames') { return [Environment]::ExpandEnvironmentVariables($v.value) }
        return $v.value }
    $key | Add-Member ScriptMethod SetValue { param($n, $v, $k)
        if (-not $this.Writable) { throw "key not opened for writing" }
        if ($global:failWrites) { throw "access denied" }
        $global:store[$this.Path][$n] = @{ kind = [string]$k; value = (Convert-FakeValue $v $k) } }
    $key | Add-Member ScriptMethod DeleteValue { param($n, $t)
        if (-not $this.Writable) { throw "key not opened for writing" }
        if ($global:failWrites) { throw "access denied" }
        $global:store[$this.Path].Remove($n) }
    $key | Add-Member ScriptMethod Close { }
    return $key
}
$global:FakeHKCU = [pscustomobject]@{}
$global:FakeHKCU | Add-Member ScriptMethod OpenSubKey { param($p, $w); if ($global:store.ContainsKey($p)) { return New-FakeKey $p $w }; return $null }
$global:FakeHKCU | Add-Member ScriptMethod CreateSubKey { param($p); if (-not $global:store.ContainsKey($p)) { $global:store[$p] = @{} }; return New-FakeKey $p $true }
`;

    it('saves, sets and puts back the setting in every case', () => {
        const fake = block.replaceAll('[Microsoft.Win32.Registry]::CurrentUser', '$global:FakeHKCU')
            .replace(/^\$TelemetryProfile = .*$/m, '')
            .replace(/^\$TelemetryBackup = .*$/m, '$TelemetryBackup = $env:XLIDE_TEST_BACKUP');
        // Fail closed if extraction ever leaves a real registry call or the real backup path behind.
        expect(fake).not.toMatch(/Win32\.Registry\]::|UserProfile/);
        const root = mkdtempSync(join(tmpdir(), 'xlide-telemetry-test-'));
        try {
            const manifest = join(root, 'scenarios.json');
            writeFileSync(manifest, JSON.stringify(scenarios.map((scenario, index) => ({ ...scenario, dir: join(root, String(index)) }))));
            const script = String.raw`
$ErrorActionPreference = 'Stop'
` + fakeRegistry + fake + String.raw`
$path = 'Software\Microsoft\Office\Common\ClientTelemetry'
function Get-Snapshot($step, $failure) {
    $key = $null
    if ($global:store.ContainsKey($path)) {
        $key = [ordered]@{}
        foreach ($name in $global:store[$path].Keys) { $key[$name] = [pscustomobject]@{ kind = $global:store[$path][$name].kind; value = [string]$global:store[$path][$name].value } }
    }
    [pscustomobject]@{ step = $step; key = $key; error = $failure
        backup = if ($TelemetryBackup -and (Test-Path -LiteralPath $TelemetryBackup)) { [IO.File]::ReadAllText($TelemetryBackup) } else { $null } }
}
$results = foreach ($scenario in (Get-Content -Raw -LiteralPath $env:XLIDE_TEST_SCENARIOS | ConvertFrom-Json)) {
    if (-not $scenario.noBackupFolder) { New-Item -ItemType Directory -Path $scenario.dir | Out-Null }
    $env:XLIDE_TEST_BACKUP = Join-Path $scenario.dir 'telemetry.json'
    $TelemetryBackup = if ($scenario.noProfile) { $null } else { $env:XLIDE_TEST_BACKUP }
    $global:store = @{}
    $global:failWrites = $false
    if ($null -ne $scenario.key) {
        $global:store[$path] = @{}
        foreach ($p in $scenario.key.PSObject.Properties) {
            $value = if ($p.Value.kind -eq 'QWord') { [long]$p.Value.value } else { $p.Value.value }
            $global:store[$path][$p.Name] = @{ kind = $p.Value.kind; value = $value }
        }
    }
    if ($null -ne $scenario.backup) { [IO.File]::WriteAllText($TelemetryBackup, $scenario.backup) }
    $errors = New-Object IO.StringWriter
    $stderr = [Console]::Error
    [Console]::SetError($errors)
    $snapshots = @()
    try {
        foreach ($step in $scenario.steps) {
            $failure = $null
            try {
                switch ($step) {
                    'disable' { Disable-OracleTelemetry }
                    'restore' { Restore-OracleTelemetry }
                    'user-sets-2' { $global:store[$path]['DisableTelemetry'] = @{ kind = 'DWord'; value = 2 } }
                    'fail-writes' { $global:failWrites = $true }
                }
            } catch { $failure = $_.Exception.Message }
            $snapshots += Get-Snapshot $step $failure
        }
    } finally { [Console]::SetError($stderr) }
    [pscustomobject]@{
        name = $scenario.name
        snapshots = $snapshots
        aside = if (Test-Path -LiteralPath $scenario.dir) { @(Get-ChildItem -LiteralPath $scenario.dir -Filter 'telemetry.json.unreadable-*').Count } else { 0 }
        stderr = $errors.ToString().Trim()
    }
}
ConvertTo-Json -InputObject @($results) -Depth 8 -Compress
`;
            const run = spawnSync('powershell.exe', ['-NoProfile', '-NonInteractive', '-EncodedCommand', Buffer.from(script, 'utf16le').toString('base64')], {
                encoding: 'utf8', timeout: 30_000, env: { ...process.env, XLIDE_TEST_SCENARIOS: manifest },
            });
            expect(run.error).toBeUndefined();
            expect(run.status, run.stderr).toBe(0);
            const outcomes: Outcome[] = JSON.parse(run.stdout.trim());
            expect(outcomes.map((outcome) => outcome.name)).toEqual(scenarios.map((scenario) => scenario.name));
            const by = Object.fromEntries(outcomes.map((outcome) => [outcome.name, outcome]));
            const at = (name: string, index: number) => by[name].snapshots[index];
            const dword = (value: number) => ({ kind: 'DWord', value: String(value) });
            const backup = (name: string, index: number) => JSON.parse(at(name, index).backup!);

            // Set to DWORD 1 with what was there saved, then put back exactly.
            const roundTrips: [string, Record<string, Value> | null, object][] = [
                ['no key', {}, { present: false }],
                ['key without the value', { Other: { kind: 'String', value: 'x' } }, { present: false }],
                ['DWord 0', { DisableTelemetry: dword(0) }, { present: true, kind: 'DWord', value: 0 }],
                ['already DWord 1', { DisableTelemetry: dword(1) }, { present: true, kind: 'DWord', value: 1 }],
                ['String', { DisableTelemetry: { kind: 'String', value: 'abc' } }, { present: true, kind: 'String', value: 'abc' }],
                ['ExpandString', { DisableTelemetry: { kind: 'ExpandString', value: '%SystemRoot%\\x' } },
                    { present: true, kind: 'ExpandString', value: '%SystemRoot%\\x' }],
                ['large QWord', { DisableTelemetry: { kind: 'QWord', value: big } }, { present: true, kind: 'QWord', value: Number(big) }],
            ];
            for (const [name, after, saved] of roundTrips) {
                expect(at(name, 0).error, name).toBeNull();
                expect(at(name, 0).key?.DisableTelemetry, name).toEqual(dword(1));
                expect(backup(name, 0), name).toEqual(saved);
                expect(at(name, 1).error, name).toBeNull();
                expect(at(name, 1).key, name).toEqual(after);
                expect(at(name, 1).backup, name).toBeNull();
                expect(by[name].stderr, name).toBe('');
            }
            // The large QWord is saved digit for digit, not through a double.
            expect(at('large QWord', 0).backup).toContain(`"value":${big}`);
            // A key the run created stays, empty: deleting a key could take another program's value with it.
            expect(at('no key', 1).key).toEqual({});
            // A kind it could not put back exactly is never touched, and nothing is saved.
            expect(at('Binary', 0).key).toEqual({ DisableTelemetry: { kind: 'Binary', value: 'bytes' } });
            expect(at('Binary', 0).backup).toBeNull();
            expect(by.Binary.stderr).toBe('Office telemetry was left on for this run: DisableTelemetry holds a Binary value it could not put back exactly.');
            // The user's own value during the run stays.
            expect(at('changed by the user during the run', 2).key).toEqual({ DisableTelemetry: dword(2) });
            expect(at('changed by the user during the run', 2).backup).toBeNull();
            // A killed run's backup is put back first, then saved again for this run.
            expect(at('a killed run left it set', 0).key).toEqual({ DisableTelemetry: dword(1) });
            expect(backup('a killed run left it set', 0)).toEqual({ present: true, kind: 'DWord', value: 0 });
            expect(at('a killed run left it set', 1).key).toEqual({ DisableTelemetry: dword(0) });
            // A backup that cannot be read, or holds a value its kind cannot, is set aside with a notice.
            for (const name of ['an unreadable backup', 'a backup whose value does not fit its kind']) {
                expect(at(name, 0).error, name).toBeNull();
                expect(by[name].aside, name).toBe(1);
                expect(by[name].stderr, name).toMatch(/^The saved Office telemetry setting could not be read; it was moved to .*telemetry\.json\.unreadable-\d+\. DisableTelemetry may stay set to 1\.$/);
                expect(at(name, 0).key, name).toEqual({ DisableTelemetry: dword(1) });
                // The value found, 1, is what this run saves: the notice says so.
                expect(backup(name, 0), name).toEqual({ present: true, kind: 'DWord', value: 1 });
            }
            // No backup, no change: the value is set only once it is saved.
            expect(at('the backup cannot be written', 0).error).not.toBeNull();
            expect(at('the backup cannot be written', 0).key).toEqual({ DisableTelemetry: dword(0) });
            expect(at('the backup cannot be written', 0).backup).toBeNull();
            // A refused restore keeps the backup for the next run.
            expect(at('the registry refuses the restore', 2).error).toContain('access denied');
            expect(at('the registry refuses the restore', 2).key).toEqual({ DisableTelemetry: dword(1) });
            expect(backup('the registry refuses the restore', 2)).toEqual({ present: true, kind: 'DWord', value: 0 });
            // No profile folder: nothing is saved, so nothing is changed.
            expect(at('no user profile folder', 0).error).toContain('there is no user profile folder');
            expect(at('no user profile folder', 0).key).toEqual({ DisableTelemetry: dword(0) });
            expect(at('no user profile folder', 1).error).toBeNull();
        } finally { rmSync(root, { recursive: true, force: true }); }
    }, 40_000);
});
