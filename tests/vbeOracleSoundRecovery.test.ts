import { spawnSync } from 'node:child_process';
import { mkdirSync, mkdtempSync, readFileSync, readdirSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { runInNewContext } from 'node:vm';
import { beforeAll, describe, expect, it } from 'vitest';

// Execute the actual generated PowerShell with every registry/process effect
// replaced by an in-memory fake. Only recovery files in this test's temp dir
// are real. Windows CI executes these; other platforms retain lifecycle tests.
const source = readFileSync(join(process.cwd(), 'syntax_corpus/oracle/run_vbe_oracle.mjs'), 'utf8');
const begin = source.indexOf('function dialogSoundsScript()');
const end = source.indexOf('// Never stops a run:', begin);
let body: string = runInNewContext(`const DIALOG_SOUNDS = ['.Default', 'SystemAsterisk', 'SystemExclamation', 'SystemHand'];\n${source.slice(begin, end)}\ndialogSoundsScript()`);
const openBegin = body.indexOf('    function Open-Sound(');
const openEnd = body.indexOf('    # A normal shell', openBegin);
body = body.slice(0, openBegin) + String.raw`
    function Open-Sound($name, $writable) {
        $key = [pscustomobject]@{ Name = [string]$name }
        $key | Add-Member ScriptMethod GetValueNames { return ,@('') }
        $key | Add-Member ScriptMethod GetValue { param($name, $default, $options); return $script:current }
        $key | Add-Member ScriptMethod SetValue { param($name, $value, $kind); $script:writes.Add([pscustomobject]@{ name=$this.Name; value=$value; kind=[string]$kind }) }
        $key | Add-Member ScriptMethod Close { }
        return $key
    }
` + body.slice(openEnd);
const startedBegin = body.indexOf('    function Started(');
const startedEnd = body.indexOf('    # Only a sound', startedBegin);
body = body.slice(0, startedBegin) + "    function Started($id) { return '1234567890' }\n" + body.slice(startedEnd);
body = body.replaceAll('exit 0', 'return');

type Probe = { saved: unknown; current: string };
type RecoveryResult = { writes: unknown[]; exists: boolean; quarantined: string[]; original: string };
type ProbeResult = RecoveryResult & { index: number; error: string | null; outputCount: number };

function runProbes(probes: Probe[]): ProbeResult[] {
    // Fail closed if extraction ever leaves a real registry/process call behind.
    expect(body).not.toMatch(/::CurrentUser|Get-CimInstance|Get-Process/);
    const root = mkdtempSync(join(tmpdir(), 'xlide-sound-recovery-test-'));
    try {
        const files = probes.map(({ saved }, index) => {
            const dir = join(root, String(index));
            mkdirSync(dir);
            const file = join(dir, 'recovery.json');
            const original = JSON.stringify(saved);
            writeFileSync(file, original);
            return { dir, file, original };
        });
        const manifest = join(root, 'probes.json');
        writeFileSync(manifest, JSON.stringify(probes.map(({ current }, index) => ({
            index, file: files[index].file, current,
        }))));
        const script = join(root, 'probe.ps1');
        writeFileSync(script, String.raw`
$ErrorActionPreference = 'Stop'
$probes = Get-Content -Raw -Encoding UTF8 -LiteralPath $env:XLIDE_TEST_PROBES | ConvertFrom-Json
$results = foreach ($probe in $probes) {
    # A child scope prevents helper definitions and recovery locals surviving
    # into the next probe. Only the two explicitly reset fake fields are shared.
    & {
        param($probe)
        $env:XLIDE_SOUNDS_FILE = $probe.file
        $script:current = $probe.current
        $script:writes = [Collections.Generic.List[object]]::new()
        $failure = $null
        $emitted = @()
        try {
            $emitted = @(& {
` + body + String.raw`
            })
        } catch { $failure = $_.Exception.ToString() }
        [pscustomobject]@{
            index = $probe.index
            writes = @($script:writes.ToArray())
            exists = (Test-Path -LiteralPath $probe.file)
            error = $failure
            outputCount = $emitted.Count
        }
    } $probe
}
ConvertTo-Json -InputObject @($results) -Depth 6 -Compress
`);
        const result = spawnSync('powershell.exe', ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', script], {
            encoding: 'utf8', timeout: 15_000,
            env: { ...process.env, XLIDE_TEST_PROBES: manifest,
                XLIDE_SOUNDS_OWNER: '4321', XLIDE_SOUNDS_ACTION: 'restore' },
        });
        expect(result.error).toBeUndefined();
        expect(result.status, result.stderr).toBe(0);
        const outputs: Omit<ProbeResult, 'quarantined' | 'original'>[] = JSON.parse(result.stdout.trim());
        // Missing, repeated or reordered results must never silently pass.
        expect(outputs.map(output => output.index)).toEqual(probes.map((_, index) => index));
        return outputs.map((output, index) => {
            const { dir, original } = files[index];
            const quarantined = readdirSync(dir).filter(name => name.startsWith('recovery.json.unreadable-'));
            return { ...output, original, quarantined: quarantined.map(name => readFileSync(join(dir, name), 'utf8')) };
        });
    } finally { rmSync(root, { recursive: true, force: true }); }
}

const valid = () => ({ owner: 4321, started: '1234567890', sounds: [{ name: '.Default', value: 'original.wav', kind: 'String' }] });

describe.skipIf(process.platform !== 'win32')('executed PowerShell sound recovery validation', () => {
    const probes: Probe[] = [];
    let results: ProbeResult[];
    // Registration retains an individual Vitest result for every probe while
    // one PowerShell startup serves the complete collection.
    function recoveryTest(name: string, saved: unknown, current: string, verify: (result: RecoveryResult) => void) {
        const index = probes.push({ saved, current }) - 1;
        it(name, () => {
            const result = results[index];
            expect(result.error).toBeNull();
            expect(result.outputCount).toBe(0);
            verify(result);
        });
    }
    beforeAll(() => { results = runProbes(probes); }, 20_000);

    for (const kind of ['String', 'ExpandString']) {
        const saved = valid(); saved.sounds[0].kind = kind; saved.sounds[0].value = String.raw`%SystemRoot%\Media\original.wav`;
        recoveryTest('restores a valid raw ' + kind + ' sound and removes the consumed backup', saved, '', result => {
            expect(result.writes).toEqual([{ name: '.Default', value: saved.sounds[0].value, kind }]);
            expect(result.exists).toBe(false);
            expect(result.quarantined).toEqual([]);
        });
    }
    const literalPath = valid();
    literalPath.sounds[0].value = String.raw`C:\sounds\'; $(throw 'must stay data'); #.wav`;
    recoveryTest('restores shell metacharacters as literal sound-path data', literalPath, '', result => {
        expect(result.writes).toEqual([{ name: '.Default', value: literalPath.sounds[0].value, kind: 'String' }]);
        expect(result.exists).toBe(false);
        expect(result.quarantined).toEqual([]);
    });
    recoveryTest('preserves a sound changed by the user during the run', valid(), 'user-changed.wav', result => {
        expect(result.writes).toEqual([]);
    });
    const malformed = [
        ['missing value', (s: any) => { delete s.sounds[0].value; }],
        ['null value', (s: any) => { s.sounds[0].value = null; }],
        ['empty value', (s: any) => { s.sounds[0].value = ''; }],
        ['object value', (s: any) => { s.sounds[0].value = { payload: 'not a path' }; }],
        ['array value', (s: any) => { s.sounds[0].value = ['a.wav']; }],
        ['number value', (s: any) => { s.sounds[0].value = 42; }],
        ['boolean value', (s: any) => { s.sounds[0].value = true; }],
        ['path traversal', (s: any) => { s.sounds[0].name = '..\\..\\Unrelated'; }],
        ['duplicate name', (s: any) => { s.sounds.push({ ...s.sounds[0], name: '.default' }); }],
        ['invalid kind', (s: any) => { s.sounds[0].kind = 'bad'; }],
        ['non-string kind', (s: any) => { s.sounds[0].kind = 'DWord'; }],
        ['missing owner', (s: any) => { delete s.owner; }],
        ['string owner', (s: any) => { s.owner = '4321'; }],
        ['invalid owner', (s: any) => { s.owner = -1; }],
        ['missing start', (s: any) => { delete s.started; }],
        ['invalid start', (s: any) => { s.started = 'unreadable'; }],
        ['empty entries', (s: any) => { s.sounds = []; }],
    ] as const;

    for (const [name, mutate] of malformed) {
        const saved = valid(); mutate(saved);
        recoveryTest('quarantines ' + name + ' before any restore write, preserving all evidence', saved, '', result => {
            expect(result.writes).toEqual([]);
            expect(result.exists).toBe(false);
            expect(result.quarantined).toEqual([result.original]);
        });
    }
    const invalidSecond: any = valid();
    invalidSecond.sounds.push({ name: 'SystemHand', kind: 'String' });
    recoveryTest('validates all entries before restoring even a valid first entry', invalidSecond, '', result => {
        expect(result.writes).toEqual([]);
        expect(result.exists).toBe(false);
        expect(result.quarantined).toEqual([result.original]);
    });
});
