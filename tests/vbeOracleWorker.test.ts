import { spawnSync } from 'child_process';
import * as fs from 'fs';
import * as path from 'path';
import { describe, expect, it } from 'vitest';
import { OFFICE_HOST_APPS } from '../src/officeHostApps';

const ORACLE = path.join(__dirname, '..', 'syntax_corpus', 'oracle');
const readOracle = (name: string): string => fs.readFileSync(path.join(ORACLE, name), 'utf8');

/** The coordinator's host table, read from its source: `host: { progId, image }`. */
function coordinatorHosts(): Record<string, { progId: string; image: string }> {
    const coordinator = readOracle('run_vbe_oracle.mjs');
    const table = /const OFFICE_HOSTS = \{([\s\S]*?)\n\};/.exec(coordinator)?.[1] ?? '';
    const hosts: Record<string, { progId: string; image: string }> = {};
    for (const match of table.matchAll(/(\w+): \{ progId: '([^']+)', image: '([^']+)'/g)) {
        hosts[match[1]] = { progId: match[2], image: match[3] };
    }
    return hosts;
}

describe('VBE oracle worker', () => {
    it('runs the host visible so the VBE error dialog is detectable', () => {
        // The Win32 dialog watcher only sees VISIBLE top-level windows. When the
        // host is hidden, the VBE Compile-error / Run-time-error modal is not
        // surfaced as a detectable visible window, so every rejection is
        // silently misrecorded as "accepted". The worker therefore keeps every
        // host visible (and shows the VBE main window so the Compile command
        // and its error dialog appear).
        const worker = readOracle('vbe_oracle_worker.ps1');
        const settings = worker.slice(
            worker.indexOf('function Set-OracleHostSettings'),
            worker.indexOf('\nfunction ', worker.indexOf('function Set-OracleHostSettings')),
        );

        for (const host of Object.keys(OFFICE_HOST_APPS)) {
            const arm = new RegExp(`"${host}" \\{([\\s\\S]*?)\\n        \\}`).exec(settings)?.[1] ?? '';
            expect(arm, host).toMatch(/\$App\.Visible = (\$true|-1)/);
        }
        expect(worker).not.toMatch(/\.Visible = (\$false|0)\b/);
        expect(worker).toContain('$app.VBE.MainWindow.Visible = $true');
    });

    it('checks that the host is alive before accepting without a dialog', () => {
        // A crashed host shows no dialog either, so "no dialog" alone recorded
        // a crash as accepted. Every accept that is not backed by a dialog must
        // first confirm that the worker's host instance is still running.
        const worker = readOracle('vbe_oracle_worker.ps1');
        const lines = worker.split(/\r?\n/).map((line) => line.trim());
        const accepts = lines
            .map((line, index) => ({ line, index }))
            .filter(({ line }) => line === '$result.outcome = "accepted"');
        const dialogAccepts = accepts.filter(({ index }) => lines[index + 1] === '$result.stage = "vbe_dialog_after_compile"');
        const plainAccepts = accepts.filter((accept) => !dialogAccepts.includes(accept));

        expect(dialogAccepts).toHaveLength(1);
        expect(plainAccepts.length).toBeGreaterThan(0);
        for (const { index } of plainAccepts) {
            expect(lines[index - 1]).toBe('Assert-HostAlive $app $appPid');
        }

        const assertStart = worker.indexOf('function Assert-HostAlive');
        const assertBody = worker.slice(assertStart, worker.indexOf('\nfunction ', assertStart));
        expect(assertBody).toContain('Get-Process -Id $HostPid');
        expect(assertBody).toContain('$null = $App.Name');
        expect(assertBody.match(/throw "/g)).toHaveLength(2);
    });

    it('ends the host once a run case has recorded its dialog, whatever the dialog', () => {
        // The dialog is a run case's evidence. In Excel 16.0 (build 20430),
        // after a run-time error in a Function declared As String, Excel
        // crashes once the dialog is dismissed and Run returns only when
        // Windows Error Reporting is done, 6 to 10 s later; after a compile
        // error Run does not return at all.
        const worker = readOracle('vbe_oracle_worker.ps1');
        const watchStart = worker.indexOf('private static void Watch(');
        const watch = worker.slice(watchStart, worker.indexOf('private static void EndProcess', watchStart));
        const recorded = watch.indexOf('WriteDialogResult(resultPath, dialog);');
        const ended = watch.indexOf('EndProcess(hostProcessId);');
        expect(recorded).toBeGreaterThan(0);
        expect(ended).toBeGreaterThan(recorded);
        expect(watch.slice(recorded, ended)).toMatch(/^WriteDialogResult\(resultPath, dialog\);\s*if \(endHostAfterDialog\) \{\s*$/);

        // Only a run case ends its host this way; a compile case keeps it.
        const starts = [...worker.matchAll(/\[XlideVbeDialogWatcher\]::Start\(([^)]*)\)/g)].map((match) => match[1]);
        expect(starts).toEqual([
            '$appPid, $DialogPath, $StagePath, $DialogWatchSeconds, $DialogHoldSeconds',
            '$appPid, $DialogPath, $StagePath, $DialogWatchSeconds, $DialogHoldSeconds, $true',
        ]);
        const runStart = worker.indexOf('Set-Stage "run"');
        expect(worker.indexOf(starts[1], runStart)).toBeGreaterThan(runStart);
    });

    it('marks served results the way the coordinator reads them', () => {
        // In -Serve mode the coordinator takes only stdout lines that start
        // with this prefix as results, so the two files must agree on it.
        const worker = readOracle('vbe_oracle_worker.ps1');
        const coordinator = readOracle('run_vbe_oracle.mjs');

        expect(coordinator).toContain("const RESULT_PREFIX = 'XLIDE-ORACLE-RESULT ';");
        expect(coordinator).toContain("'-Serve'");
        expect(worker).toContain('[Console]::Out.WriteLine("XLIDE-ORACLE-RESULT " + (Write-ResultJson $output[-1]))');
    });

    it('drives every Office host XLIDE supports, as src/officeHostApps.ts names them', () => {
        // The oracle is a COM driver outside the extension's build, so it keeps
        // its own copy of the host table; this keeps the copy honest.
        const hosts = coordinatorHosts();
        const expected = Object.fromEntries(Object.entries(OFFICE_HOST_APPS)
            .map(([host, info]) => [host, { progId: info.progId, image: info.processName }]));
        expect(hosts).toEqual(expected);

        // Each host has its own arm wherever the worker does something
        // host-specific.
        const worker = readOracle('vbe_oracle_worker.ps1');
        for (const fn of ['Set-OracleHost', 'Set-OracleHostSettings', 'New-OracleContainer', 'Close-OracleContainer', 'Get-RunMacroNames']) {
            const start = worker.indexOf(`function ${fn}`);
            expect(start, fn).toBeGreaterThan(0);
            const body = worker.slice(start, worker.indexOf('\nfunction ', start + 1));
            for (const host of Object.keys(OFFICE_HOST_APPS)) {
                expect(body, `${fn} handles ${host}`).toContain(`"${host}"`);
            }
        }
    });

    it('names only a known host in every fixture, and covers each one', () => {
        const corpus: { cases: Array<{ id: string; host?: string }> } = JSON.parse(readOracle('vbe_oracle_cases.json'));
        const known = new Set(Object.keys(OFFICE_HOST_APPS));
        const unknown = corpus.cases.filter((kase) => kase.host !== undefined && !known.has(kase.host)).map((kase) => kase.id);
        expect(unknown).toEqual([]);
        for (const host of known) {
            expect(corpus.cases.some((kase) => (kase.host ?? 'excel') === host), host).toBe(true);
        }
    });

    it('counts a Run error that comes after VBA finished as a return, and nothing else', () => {
        // A Variant Function that VBA leaves early without an error (as at
        // `Arr 1`, issue609_13/17) makes the host's Run return E_INVALIDARG,
        // and a returned Date beyond .NET's range fails to convert. Both are
        // returns; the run catch must test for them before anything else, so
        // the dialog and alive checks after the loop still decide the case.
        const worker = readOracle('vbe_oracle_worker.ps1');
        const runStart = worker.indexOf('Set-Stage "run"');
        const run = worker.slice(runStart, worker.indexOf('Resolve-DialogAfterStop', runStart));
        expect(run).toMatch(/catch \{\s*if \(Test-RunReturnError \$_\.Exception\) \{\s*\$lastError = \$null\s*break\s*\}/);
        expect(run.match(/\[void\]\$app\.Run\(/g)).toHaveLength(1);

        const messages = /\$ReturnErrorMessages = @\(([^)]*)\)/.exec(worker)?.[1] ?? '';
        expect([...messages.matchAll(/"([^"]+)"/g)].map((match) => match[1])).toEqual([
            'Not a legal OleAut date.',
            'Value does not fall within the expected range.',
        ]);
    });

    // Starting Windows PowerShell can take over five seconds on a busy CI runner.
    it.skipIf(process.platform !== 'win32')('tells a return error from any other Run error in PowerShell', () => {
        // A Windows checkout has CRLF lines; made LF here, the cut below ends
        // with the function's own line break, and PowerShell reading commands
        // from standard input needs the blank line that then follows it to
        // run the definition.
        const worker = readOracle('vbe_oracle_worker.ps1').replaceAll('\r\n', '\n');
        const start = worker.indexOf('$ReturnErrorMessages = @(');
        const fn = worker.indexOf('function Test-RunReturnError');
        const end = fn + (/\n\}\n/.exec(worker.slice(fn))?.index ?? -fn) + 3;
        const probes = [
            ['wrapped date', true, 'New-Wrapped ([ArgumentException]::new("Not a legal OleAut date."))'],
            ['wrapped E_INVALIDARG', true, 'New-Wrapped ([ArgumentException]::new("Value does not fall within the expected range."))'],
            ['bare E_INVALIDARG', true, '[ArgumentException]::new("Value does not fall within the expected range.")'],
            ['bare date', true, '[ArgumentException]::new("Not a legal OleAut date.")'],
            ['other message', false, 'New-Wrapped ([ArgumentException]::new("Value does not fall within the expected range"))'],
            ['other case', false, '[ArgumentException]::new("value does not fall within the expected range.")'],
            // A subclass with the very message: only the type check refuses it.
            ['derived type', false, '[System.Text.DecoderFallbackException]::new("Value does not fall within the expected range.")'],
            ['macro not found', false, 'New-Wrapped ([Runtime.InteropServices.COMException]::new("Cannot run the macro", -2146827284))'],
            ['remote call failed', false, '[Runtime.InteropServices.COMException]::new("The remote procedure call failed.", -2147023170)'],
            ['wrapped nothing', false, 'New-Wrapped $null'],
        ] as const;
        const script = [
            worker.slice(start, end),
            'function New-Wrapped($Inner) { [System.Management.Automation.MethodInvocationException]::new("Exception calling Run", $Inner) }',
            ...probes.map(([name, , exception]) => `"${name}=" + (Test-RunReturnError (${exception}))`),
        ].join('\n');
        const run = spawnSync('powershell.exe', ['-NoProfile', '-NonInteractive', '-Command', '-'], { input: script, encoding: 'utf8', timeout: 12_000 });
        expect(run.error).toBeUndefined();
        expect(run.stderr).toBe('');
        const lines = run.stdout.trim().split(/\r?\n/);
        expect(lines).toEqual(probes.map(([name, expected]) => `${name}=${expected ? 'True' : 'False'}`));
    }, 15_000);

    it('keeps the old coordinator name working', () => {
        expect(readOracle('run_excel_vbe_oracle.mjs')).toContain("await import('./run_vbe_oracle.mjs');");
    });
});
