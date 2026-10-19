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

type Probe = { name: string; saved: unknown; current?: string };
type Outcome = { writes: unknown[]; exists: boolean; quarantined: string[]; original: string; failure: string | null; output: number };

// Every case runs in one Windows PowerShell, since starting one takes seconds
// on a busy CI runner. Each case has its own folder, saved file, current value
// and list of writes, and the script keeps no other state between them.
function probeAll(probes: Probe[]): Map<string, Outcome> {
    // Never silently fall back to a real registry or process implementation.
    expect(body).not.toMatch(/::CurrentUser|Get-CimInstance|Get-Process/);
    const root = mkdtempSync(join(tmpdir(), 'xlide-sound-recovery-test-'));
    try {
        const listed = probes.map((probe, index) => {
            const dir = join(root, String(index));
            mkdirSync(dir);
            const file = join(dir, 'recovery.json');
            writeFileSync(file, JSON.stringify(probe.saved));
            return { name: probe.name, dir, file, current: probe.current ?? '' };
        });
        const list = join(root, 'cases.json');
        writeFileSync(list, JSON.stringify(listed));
        const script = join(root, 'probe.ps1');
        writeFileSync(script, String.raw`
$ErrorActionPreference = 'Stop'
function Invoke-Probe {
` + body + String.raw`
}
# Windows PowerShell's ConvertFrom-Json passes an array on as one object;
# from a variable, foreach takes its items one by one.
$cases = Get-Content -Raw -Encoding UTF8 -LiteralPath $env:XLIDE_TEST_CASES | ConvertFrom-Json
$results = foreach ($case in $cases) {
    $env:XLIDE_SOUNDS_FILE = $case.file
    $script:current = $case.current
    $script:writes = [Collections.Generic.List[object]]::new()
    $failure = $null
    $output = @()
    try { $output = @(Invoke-Probe) } catch { $failure = $_.Exception.Message }
    [pscustomobject]@{ name = $case.name; writes = @($script:writes.ToArray()); exists = (Test-Path -LiteralPath $case.file); failure = $failure; output = $output.Count }
}
ConvertTo-Json -InputObject @($results) -Depth 6 -Compress
`);
        const result = spawnSync('powershell.exe', ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', script], {
            encoding: 'utf8', timeout: 15_000,
            env: { ...process.env, XLIDE_TEST_CASES: list, XLIDE_SOUNDS_OWNER: '4321', XLIDE_SOUNDS_ACTION: 'restore' },
        });
        expect(result.error).toBeUndefined();
        expect(result.status, result.stderr).toBe(0);
        const outputs: { name: string; writes: unknown[]; exists: boolean; failure: string | null; output: number }[] = JSON.parse(result.stdout.trim());
        expect(outputs.map(output => output.name)).toEqual(listed.map(item => item.name));
        return new Map(outputs.map((output, index) => {
            const { dir } = listed[index];
            const quarantined = readdirSync(dir).filter(name => name.startsWith('recovery.json.unreadable-'));
            return [output.name, { writes: output.writes, exists: output.exists, failure: output.failure, output: output.output, original: JSON.stringify(probes[index].saved),
                quarantined: quarantined.map(name => readFileSync(join(dir, name), 'utf8')) }];
        }));
    } finally { rmSync(root, { recursive: true, force: true }); }
}

const valid = () => ({ owner: 4321, started: '1234567890', sounds: [{ name: '.Default', value: 'original.wav', kind: 'String' }] });

const kinds = ['String', 'ExpandString'] as const;
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
const probes: Probe[] = [
    ...kinds.map((kind) => {
        const saved = valid(); saved.sounds[0].kind = kind; saved.sounds[0].value = '%SystemRoot%\\Media\\original.wav';
        return { name: `valid ${kind}`, saved };
    }),
    { name: 'changed by the user', saved: valid(), current: 'user-changed.wav' },
    ...malformed.map(([name, mutate]) => { const saved = valid(); mutate(saved); return { name, saved }; }),
    { name: 'invalid second entry', saved: (() => { const saved: any = valid(); saved.sounds.push({ name: 'SystemHand', kind: 'String' }); return saved; })() },
];

describe.skipIf(process.platform !== 'win32')('executed PowerShell sound recovery validation', () => {
    let results: Map<string, Outcome>;
    beforeAll(() => { results = probeAll(probes); }, 20_000);
    // A case whose script failed, or wrote to its output, fails its own test only.
    const outcome = (name: string) => {
        const result = results.get(name)!;
        expect(result.failure).toBeNull();
        expect(result.output).toBe(0);
        return result;
    };

    it.each(kinds)('restores a valid raw %s sound and removes the consumed backup', kind => {
        const result = outcome(`valid ${kind}`);
        expect(result.writes).toEqual([{ name: '.Default', value: '%SystemRoot%\\Media\\original.wav', kind }]);
        expect(result.exists).toBe(false);
        expect(result.quarantined).toEqual([]);
    });
    it('preserves a sound changed by the user during the run', () => {
        expect(outcome('changed by the user').writes).toEqual([]);
    });
    it.each(malformed)('quarantines %s before any restore write, preserving all evidence', (name) => {
        const result = outcome(name);
        expect(result.writes).toEqual([]);
        expect(result.exists).toBe(false);
        expect(result.quarantined).toEqual([result.original]);
    });
    it('validates all entries before restoring even a valid first entry', () => {
        expect(outcome('invalid second entry').writes).toEqual([]);
    });
});
