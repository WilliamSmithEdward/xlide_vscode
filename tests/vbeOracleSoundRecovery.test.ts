import { spawnSync } from 'node:child_process';
import { mkdtempSync, readFileSync, readdirSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { runInNewContext } from 'node:vm';
import { describe, expect, it } from 'vitest';

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

function probe(saved: unknown, current = '') {
    // Never silently fall back to a real registry or process implementation.
    expect(body).not.toMatch(/::CurrentUser|Get-CimInstance|Get-Process/);
    const dir = mkdtempSync(join(tmpdir(), 'xlide-sound-recovery-test-'));
    const file = join(dir, 'recovery.json');
    const original = JSON.stringify(saved);
    writeFileSync(file, original);
    const script = join(dir, 'probe.ps1');
    writeFileSync(script, String.raw`
$ErrorActionPreference = 'Stop'
$file = $env:XLIDE_TEST_FILE
$script:current = $env:XLIDE_TEST_CURRENT
$script:writes = [Collections.Generic.List[object]]::new()
function Invoke-Probe {
` + body + String.raw`
}
Invoke-Probe
[pscustomobject]@{writes=@($script:writes.ToArray());exists=(Test-Path -LiteralPath $file)} | ConvertTo-Json -Depth 6 -Compress
`);
    try {
        const result = spawnSync('powershell.exe', ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', script], {
            encoding: 'utf8', timeout: 10_000,
            env: { ...process.env, XLIDE_TEST_FILE: file, XLIDE_TEST_CURRENT: current,
                XLIDE_SOUNDS_FILE: file, XLIDE_SOUNDS_OWNER: '4321', XLIDE_SOUNDS_ACTION: 'restore' },
        });
        expect(result.error).toBeUndefined();
        expect(result.status, result.stderr).toBe(0);
        const output = JSON.parse(result.stdout.trim());
        const quarantined = readdirSync(dir).filter(name => name.startsWith('recovery.json.unreadable-'));
        return { ...output, quarantined: quarantined.map(name => readFileSync(join(dir, name), 'utf8')), original };
    } finally { rmSync(dir, { recursive: true, force: true }); }
}

const valid = () => ({ owner: 4321, started: '1234567890', sounds: [{ name: '.Default', value: 'original.wav', kind: 'String' }] });

describe.skipIf(process.platform !== 'win32')('executed PowerShell sound recovery validation', () => {
    it.each(['String', 'ExpandString'])('restores a valid raw %s sound and removes the consumed backup', kind => {
        const saved = valid(); saved.sounds[0].kind = kind; saved.sounds[0].value = '%SystemRoot%\\Media\\original.wav';
        const result = probe(saved);
        expect(result.writes).toEqual([{ name: '.Default', value: saved.sounds[0].value, kind }]);
        expect(result.exists).toBe(false);
        expect(result.quarantined).toEqual([]);
    });
    it('preserves a sound changed by the user during the run', () => {
        expect(probe(valid(), 'user-changed.wav').writes).toEqual([]);
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
    it.each(malformed)('quarantines %s before any restore write, preserving all evidence', (_name, mutate) => {
        const saved = valid(); mutate(saved);
        const result = probe(saved);
        expect(result.writes).toEqual([]);
        expect(result.exists).toBe(false);
        expect(result.quarantined).toEqual([result.original]);
    });
    it('validates all entries before restoring even a valid first entry', () => {
        const saved: any = valid(); saved.sounds.push({ name: 'SystemHand', kind: 'String' });
        expect(probe(saved).writes).toEqual([]);
    });
});
