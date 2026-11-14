import { spawnSync } from 'node:child_process';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import { OFFICE_HOST_APPS } from '../../src/officeHostApps';

const ORACLE = path.join(__dirname, '..', '..', 'syntax_corpus', 'oracle');
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
        // The one window hidden is Word's VBE while a project reference is
        // added (issue #1565); the compile shows it again.
        const wordHide = 'if ($script:hostName -eq "word") {\n                $app.VBE.MainWindow.Visible = $false';
        const unix = worker.replaceAll('\r\n', '\n');
        expect(unix.split(wordHide)).toHaveLength(2);
        expect(unix.replace(wordHide, '')).not.toMatch(/\.Visible = (\$false|0)\b/);
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
        const recorded = watch.indexOf('WriteDialogResult(recording, dialog);');
        const ended = watch.indexOf('EndProcess(hostProcessId);');
        expect(recorded).toBeGreaterThan(0);
        expect(ended).toBeGreaterThan(recorded);
        expect(watch.slice(recorded, ended)).toMatch(/^WriteDialogResult\(recording, dialog\);\s*if \(endHostAfterDialog\) \{\s*$/);

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

    it('keeps the old coordinator name working', () => {
        expect(readOracle('run_excel_vbe_oracle.mjs')).toContain("await import('./run_vbe_oracle.mjs');");
    });

    // Issue #1419: what a designer or an attribute must declare is read back
    // after setup, and a read-back that fails is a setup error, never evidence.
    const body = (worker: string, name: string): string => {
        const start = worker.indexOf(`function ${name}(`);
        expect(start, name).toBeGreaterThanOrEqual(0);
        return worker.slice(start, worker.indexOf('\nfunction ', start));
    };

    it('reads back every designer control by name and measured type', () => {
        const worker = readOracle('vbe_oracle_worker.ps1');
        const controls = body(worker, 'Add-DesignerControls');
        expect(controls).toContain('$designer.Controls.Add(');
        expect(controls).toContain('$designer.Controls.Item(');
        expect(controls).toContain('[Microsoft.VisualBasic.Information]::TypeName($found)');
        expect(controls.match(/throw "/g)?.length).toBeGreaterThanOrEqual(4);
        // Only ProgIDs whose read-back type was measured are taken.
        const table = /\$DesignerControlInterfaces = @\{([\s\S]*?)\n\}/.exec(worker)?.[1] ?? '';
        expect([...table.matchAll(/"(Forms\.\w+\.\d+)" = "(\w+)"/g)].map((m) => [m[1], m[2]])).toEqual([
            ['Forms.CommandButton.1', 'ICommandButton'],
            ['Forms.ComboBox.1', 'IMdcCombo'],
        ]);
        expect(body(worker, 'Add-OracleComponent')).toContain('Add-DesignerControls $component $ModuleSpec');
    });

    it('renames a document module to its codeName and reads the name back', () => {
        const worker = readOracle('vbe_oracle_worker.ps1');
        const rename = body(worker, 'Set-DocumentCodeName');
        expect(rename).toContain('$Component.Name = [string]$ModuleSpec.codeName');
        expect(rename).toMatch(/if \(\[string\]\$Component\.Name -cne \[string\]\$ModuleSpec\.codeName\) \{\s*throw "/);
        // Before the case's code goes in, so the code compiles under the new name.
        const component = body(worker, 'Add-OracleComponent');
        expect(component.indexOf('Set-DocumentCodeName $component $ModuleSpec'))
            .toBeLessThan(component.indexOf('$component.CodeModule.AddFromString($source)'));
        expect(component.indexOf('Set-DocumentCodeName $component $ModuleSpec')).toBeGreaterThanOrEqual(0);
    });

    it('imports module attributes and reads every attribute line back', () => {
        const worker = readOracle('vbe_oracle_worker.ps1');
        const component = body(worker, 'Add-OracleComponent');
        expect(component).toMatch(/\$moduleAttributes = \$source -match/);
        expect(component).toContain('Assert-ImportedAttributes $imported $source $extension');
        const readBack = body(worker, 'Assert-ImportedAttributes');
        expect(readBack).toContain('$Component.Export($readBack)');
        expect(readBack).toContain('throw "The imported module');
        expect(readBack).toContain('Remove-Item -LiteralPath $readBack');
    });

    it('records the selection only after a compile dialog, and never fails a case on it', () => {
        const worker = readOracle('vbe_oracle_worker.ps1');
        const selection = body(worker, 'Get-SelectionAfterDialog');
        expect(selection).toContain('$pane.GetSelection([ref]$startLine, [ref]$startColumn, [ref]$endLine, [ref]$endColumn)');
        expect(selection).toMatch(/catch \{\s*return \$null\s*\}/);
        const lines = worker.split(/\r?\n/).map((line) => line.trim());
        const at = lines.indexOf('$selection = Get-SelectionAfterDialog $app');
        expect(lines.filter((line) => line.includes('Get-SelectionAfterDialog $app'))).toHaveLength(1);
        expect(lines[at - 1]).toBe('Set-DialogOutcome $dialog $true');
    });
});


it.skipIf(process.platform !== 'win32')('distinguishes invocation failures from normal returns in the worker run path', () => {
    const worker = readOracle('vbe_oracle_worker.ps1').replaceAll('\r\n', '\n');
    const functions = ['Test-MacroNotFound', 'Assert-HostAlive', 'Resolve-DialogAfterStop', 'Set-DialogOutcome', 'Get-RunMacroNames']
        .map(name => {
            const start = worker.indexOf('function ' + name + '(');
            expect(start, name).toBeGreaterThanOrEqual(0);
            return worker.slice(start, worker.indexOf('\n}', start) + 2);
        }).join('\n');
    const runStart = worker.indexOf('        elseif ($mode -eq "run" -and $case.entryPoint) {');
    const verdict = worker.indexOf('        if ($result.outcome -eq "setup_error")', runStart);
    const cleanup = worker.indexOf('\n    finally {', verdict);
    expect(runStart).toBeGreaterThan(0);
    expect(verdict).toBeGreaterThan(runStart);
    expect(cleanup).toBeGreaterThan(verdict);
    // Run the production branch, acceptance fallback and outer catch. Only the
    // COM host and dialog watcher are fakes; no Office application is started.
    const flow = worker.slice(runStart, cleanup).trimStart().replace(/^elseif/, 'if');
    const script = String.raw`
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public class OracleInvocationProbe {
    public string Mode;
    public bool MacroExecuted;
    public string Name { get { return "Probe"; } }
    public void Run(string macro) {
        if (Mode == "argument") throw new ArgumentException("Value does not fall within the expected range.");
        if (Mode == "rpc") throw new COMException("The remote procedure call failed.", -2147023170);
        if (Mode == "missing") throw new COMException("Cannot run the macro", -2146827284);
        MacroExecuted = true;
        if (Mode == "date") throw new ArgumentException("Not a legal OleAut date.");
    }
}
public static class XlideVbeDialogWatcher {
    public static bool DialogSeen = false;
    public static void Start(int pid, object result, object stage, object seconds, object hold, bool end) {}
    public static void Stop() {}
}
"@
function Set-Stage($value) { $script:stage = $value }
function Wait-ForRunDialogResult { return $script:dialog }
function Read-DialogResult { return $script:dialog }
` + functions + String.raw`
$results = foreach ($hostName in @('excel', 'word', 'powerpoint', 'access')) {
    foreach ($scenario in @('success', 'date', 'argument', 'rpc', 'missing', 'dialog')) {
        $script:hostName = $hostName
        $script:hostNoun = $hostName
        $script:dialog = if ($scenario -eq 'dialog') { [pscustomobject]@{ kind='vbe_runtime_dialog'; message='Run-time error' } } else { $null }
        $app = [OracleInvocationProbe]::new()
        $app.Mode = if ($scenario -eq 'dialog') { 'argument' } else { $scenario }
        $appPid = $PID
        $container = [pscustomobject]@{ Name='Scratch' }
        $component = [pscustomobject]@{ Name='Probe' }
        $vbProject = [pscustomobject]@{ Name='Oracle' }
        $case = [pscustomobject]@{ entryPoint='Main' }
        $mode = 'run'
        $result = @{ outcome='setup_error'; stage='setup'; message=''; hresult=$null }
        try {
` + flow + String.raw`
        [pscustomobject]@{ host=$hostName; scenario=$scenario; outcome=$result.outcome; stage=$result.stage; macroExecuted=$app.MacroExecuted }
    }
}
ConvertTo-Json -InputObject @($results) -Compress
`;
    // EncodedCommand handles CRLF and multiline definitions without relying on
    // Windows PowerShell's interactive stdin parsing or interpolating values.
    const run = spawnSync('powershell.exe', ['-NoProfile', '-NonInteractive', '-EncodedCommand', Buffer.from(script, 'utf16le').toString('base64')], {
        encoding: 'utf8', timeout: 30_000,
    });
    expect(run.error).toBeUndefined();
    expect(run.status, run.stderr).toBe(0);
    const results = JSON.parse(run.stdout.trim());
    const expected = Object.keys(OFFICE_HOST_APPS).flatMap(host => [
        { host, scenario: 'success', outcome: 'accepted', stage: 'run', macroExecuted: true },
        { host, scenario: 'date', outcome: 'accepted', stage: 'run', macroExecuted: true },
        ...['argument', 'rpc', 'missing'].map(scenario => ({ host, scenario, outcome: 'worker_error', stage: 'run', macroExecuted: false })),
        { host, scenario: 'dialog', outcome: 'rejected', stage: 'runtime_dialog', macroExecuted: false },
    ]);
    expect(results).toEqual(expected);
}, 40_000);

describe('the macro Security Notice (issue #1529)', () => {
    const worker = readOracle('vbe_oracle_worker.ps1').replaceAll('\r\n', '\n');

    it('declines the notice while a project reference is added, and records no verdict', () => {
        const stage = worker.indexOf('Set-Stage "add_project_reference"');
        const next = worker.indexOf('Set-Stage', stage + 1);
        const adding = worker.slice(stage, next);
        expect(adding).toContain('[XlideSecurityNoticeWatcher]::Start($appPid, $DialogWatchSeconds)');
        expect(adding).toContain('$script:securityNotice = [XlideSecurityNoticeWatcher]::Stop()');
        expect(worker).toContain('$result.stage = "security_notice"');
        // Never pressed: the watcher only posts WM_CLOSE, and no button is clicked.
        const watcher = worker.slice(worker.indexOf('public static class XlideSecurityNoticeWatcher'), worker.indexOf('public static class XlideVbeDialogWatcher'));
        expect(watcher).toContain('PostMessage(hWnd, WM_CLOSE, IntPtr.Zero, IntPtr.Zero)');
        expect(watcher).not.toMatch(/BM_CLICK|Invoke\(|SendKeys|Enable Macros/);
        // No Trust Center setting, Trusted Location or AutomationSecurity is written for it.
        expect(watcher).not.toMatch(/Registry|AutomationSecurity|Trusted/);
    });

    it('adds the reference in the ways that raise no notice, after checking the provider (issue #1565)', () => {
        const stage = worker.indexOf('Set-Stage "add_project_reference"');
        const adding = worker.slice(stage, worker.indexOf('Set-Stage', stage + 1));
        // Word's VBE is hidden before the watcher starts.
        expect(adding.indexOf('$app.VBE.MainWindow.Visible = $false')).toBeGreaterThan(-1);
        expect(adding.indexOf('$app.VBE.MainWindow.Visible = $false')).toBeLessThan(adding.indexOf('[XlideSecurityNoticeWatcher]::Start('));
        // The provider container is kept before a declined notice ends the case.
        expect(adding.indexOf('$provider = $madeProvider.Container')).toBeLessThan(adding.indexOf('throw "Security notice declined."'));

        const making = worker.slice(worker.indexOf('function New-ProjectReference('), worker.indexOf('function New-AccessProjectReference('));
        // The provider is checked before anything is made, in every host.
        expect(making.indexOf('Assert-InertProvider $Case.projectReference')).toBeLessThan(making.indexOf('New-AccessProjectReference'));
        // The path is kept as soon as it is chosen, for the clean-up.
        expect(making.indexOf('$script:providerPath = $path')).toBeLessThan(making.indexOf('New-OracleContainer'));
        // Only PowerPoint loads the add-in, and before the reference names it.
        expect(making.split('AddIns.Add(')).toHaveLength(2);
        expect(making.indexOf('if ($script:hostName -eq "powerpoint") {\n            $addIn = $App.AddIns.Add($path)')).toBeGreaterThan(-1);
        expect(making.indexOf('AddIns.Add(')).toBeLessThan(making.indexOf('References.AddFromFile($path)'));
        expect(making).not.toMatch(/Registry|AutomationSecurity\s*=|Trusted|\.Registered|AutoLoad/);
        const access = worker.slice(worker.indexOf('function New-AccessProjectReference('), worker.indexOf('function Invoke-OracleCase'));
        expect(access).toContain('$script:providerPath = $path');

        // A PowerPoint instance that still lists an add-in is not kept, and
        // the provider's file is deleted after the host may have quit.
        expect(worker).toContain('try { $closed = $script:app.AddIns.Count -eq 0 } catch { $closed = $false }');
        const cleanup = worker.slice(worker.indexOf('$evidence = $result.outcome -eq'));
        expect(cleanup.indexOf('Stop-OracleHost')).toBeLessThan(cleanup.indexOf('Remove-Item -LiteralPath $candidate'));
        expect(cleanup).toContain('foreach ($candidate in @($script:providerPath,');
    });

    it.skipIf(process.platform !== 'win32')('refuses a provider whose code could run when it loads (issue #1565)', () => {
        const start = worker.indexOf('function Assert-InertProvider(');
        const check = worker.slice(start, worker.indexOf('\nfunction ', start + 1));
        const cases: Array<[string, unknown, boolean]> = [
            ['plain standard and class modules', [{ name: 'Service', type: 'standard', source: 'Public Function Value()\r\n    Value = 1\r\nEnd Function\r\n' }, { name: 'Thing', type: 'class', source: '' }], true],
            ['a name that only contains one', [{ name: 'Service', type: 'standard', source: 'Public Sub MyAuto_Open()\r\nEnd Sub\r\nPublic Sub AutoOpener()\r\nEnd Sub\r\n' }], true],
            ['Auto_Open', [{ name: 'Service', type: 'standard', source: 'Sub Auto_Open()\r\nEnd Sub\r\n' }], false],
            ['AUTO_CLOSE in any case', [{ name: 'Service', type: 'standard', source: 'Private Sub AUTO_CLOSE()\r\nEnd Sub\r\n' }], false],
            ['Word AutoExec', [{ name: 'Service', type: 'standard', source: 'Public Sub AutoExec()\r\nEnd Sub\r\n' }], false],
            ['a module named AutoExec', [{ name: 'AutoExec', type: 'standard', source: 'Sub Main()\r\nEnd Sub\r\n' }], false],
            ['AutoNew in a class', [{ name: 'Thing', type: 'class', source: 'Public Sub AutoNew()\r\nEnd Sub\r\n' }], false],
            ['AutoOpen named in a comment', [{ name: 'Service', type: 'standard', source: "' AutoOpen is not here\r\n" }], false],
            ['a document module', [{ name: 'ThisWorkbook', type: 'document', source: '' }], false],
        ];
        const script = String.raw`
$ErrorActionPreference = 'Stop'
. ([scriptblock]::Create($env:XLIDE_INERT_CHECK))
$cases = $env:XLIDE_INERT_CASES | ConvertFrom-Json
$seen = foreach ($modules in $cases) {
    try { Assert-InertProvider ([pscustomobject]@{ modules = @($modules) }); 'ok' } catch { 'refused: ' + $_.Exception.Message }
}
@($seen) | ConvertTo-Json -Compress
`;
        const run = spawnSync('powershell.exe', ['-NoProfile', '-NonInteractive', '-EncodedCommand', Buffer.from(script, 'utf16le').toString('base64')], {
            encoding: 'utf8',
            timeout: 30_000,
            env: { ...process.env, XLIDE_INERT_CHECK: check, XLIDE_INERT_CASES: JSON.stringify(cases.map(([, modules]) => modules)) },
        });
        expect(run.error).toBeUndefined();
        expect(run.status, run.stderr).toBe(0);
        const seen: string[] = JSON.parse(run.stdout.trim());
        cases.forEach(([label, , inert], i) => {
            if (inert) {
                expect(seen[i], label).toBe('ok');
            } else {
                expect(seen[i], label).toMatch(/^refused: The referenced project's module /);
            }
        });
    }, 40_000);

    it.skipIf(process.platform !== 'win32')('closes only a notice of the watched process, and only one so titled', () => {
        const start = worker.indexOf('public static class XlideSecurityNoticeWatcher');
        const source = 'using System;\nusing System.Runtime.InteropServices;\nusing System.Text;\nusing System.Threading;\n'
            + worker.slice(start, worker.indexOf('public static class XlideVbeDialogWatcher'));
        // Stand-in windows of this PowerShell, placed off screen. Process 4 is
        // System, which owns no window: watching it must close nothing.
        const script = String.raw`
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -TypeDefinition $env:XLIDE_NOTICE_WATCHER
function New-StandIn($title) {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = $title
    $form.ShowInTaskbar = $false
    $form.StartPosition = 'Manual'
    $form.Location = New-Object System.Drawing.Point(-32000, -32000)
    $form.Show()
    return $form
}
function Watch-For($processId, $seconds) {
    [XlideSecurityNoticeWatcher]::Start([uint32]$processId, 10)
    $until = (Get-Date).AddSeconds($seconds)
    while ((Get-Date) -lt $until) { [System.Windows.Forms.Application]::DoEvents(); Start-Sleep -Milliseconds 20 }
    $started = Get-Date
    $title = [XlideSecurityNoticeWatcher]::Stop()
    return @{ title = $title; stopMs = ((Get-Date) - $started).TotalMilliseconds }
}
$notice = New-StandIn 'Microsoft PowerPoint Security Notice'
$other = New-StandIn 'Microsoft PowerPoint'
$elsewhere = Watch-For 4 1
$noticeAfterElsewhere = $notice.Visible
$own = Watch-For $PID 2
$none = Watch-For $PID 1
$otherAfterOwn = -not $other.IsDisposed
$other.Close()
[pscustomobject]@{
    elsewhereTitle = $elsewhere.title
    noticeAfterElsewhere = $noticeAfterElsewhere
    ownTitle = $own.title
    noticeAfterOwn = $notice.Visible
    otherAfterOwn = $otherAfterOwn
    noneTitle = $none.title
    noneStopMs = $none.stopMs
} | ConvertTo-Json -Compress
`;
        const run = spawnSync('powershell.exe', ['-NoProfile', '-NonInteractive', '-STA', '-EncodedCommand', Buffer.from(script, 'utf16le').toString('base64')], {
            encoding: 'utf8', timeout: 30_000, env: { ...process.env, XLIDE_NOTICE_WATCHER: source },
        });
        expect(run.error).toBeUndefined();
        expect(run.status, run.stderr).toBe(0);
        const seen = JSON.parse(run.stdout.trim());
        expect(seen.elsewhereTitle).toBeNull();
        expect(seen.noticeAfterElsewhere).toBe(true);
        expect(seen.ownTitle).toBe('Microsoft PowerPoint Security Notice');
        expect(seen.noticeAfterOwn).toBe(false);
        expect(seen.otherAfterOwn).toBe(true);
        // A notice that never comes is not waited for.
        expect(seen.noneTitle).toBeNull();
        expect(seen.noneStopMs).toBeLessThan(1000);
    }, 40_000);
});

// Waits that end on a signal instead of a poll, and work skipped when its
// answer cannot change what happens. None of it may change a verdict.
describe('the worker waits on signals, not polls', () => {
    const worker = readOracle('vbe_oracle_worker.ps1').replaceAll('\r\n', '\n');
    const coordinator = readOracle('run_vbe_oracle.mjs').replaceAll('\r\n', '\n');
    const fn = (name: string): string => {
        const start = worker.search(new RegExp(`\\nfunction ${name}[ ({]`)) + 1;
        expect(start, name).toBeGreaterThan(0);
        return worker.slice(start, worker.indexOf('\nfunction ', start));
    };
    const watcherSource = 'using System;\nusing System.Collections.Generic;\nusing System.Diagnostics;\nusing System.IO;\n'
        + 'using System.Runtime.InteropServices;\nusing System.Text;\nusing System.Threading;\n'
        + worker.slice(worker.indexOf('public static class XlideVbeDialogWatcher'), worker.indexOf('\n"@'));

    it('scans for a dialog every 15 ms with the same rule, and hands its result over in memory', () => {
        const watch = watcherSource.slice(watcherSource.indexOf('private static void Watch('), watcherSource.indexOf('private static void EndProcess('));
        expect(watch).toContain('stopEvent.WaitOne(15);');
        expect(watch).not.toContain('WaitOne(100)');
        // The result is handed over only once the dialog is dismissed and its file written.
        expect(watch.indexOf('DismissDialog(dialog.Handle)')).toBeLessThan(watch.indexOf('WriteDialogResult(recording, dialog)'));
        const write = watcherSource.slice(watcherSource.indexOf('private static void WriteDialogResult('), watcherSource.indexOf('private static string JsonString(')).trimEnd();
        expect(write).toMatch(/File\.WriteAllText\(recording\.Path, json\.ToString\(\), Encoding\.UTF8\);\n\s+recording\.Json = json\.ToString\(\);\n\s+recording\.Written\.Set\(\);\n\s+\}$/);
        // Every start forgets the last case's dialog, and its watcher writes
        // only to the record made for it: one that outlives Stop() cannot
        // touch a later case's.
        const start = watcherSource.slice(watcherSource.lastIndexOf('public static void Start('), watcherSource.indexOf('public static void Stop('));
        for (const reset of ['current = null;', 'dialogSeen = false;', 'Recording recording = new Recording(resultPath);', 'current = recording;', 'Watch(hostProcessId, recording,']) {
            expect(start).toContain(reset);
        }
        expect(watcherSource).not.toMatch(/\bcurrent\.(Json|Written)\b/);
        // The worker waits for the result instead of reading the file again and again.
        expect(fn('Wait-ForSeenDialogResult')).not.toContain('Start-Sleep');
        expect(fn('Read-DialogResult')).not.toMatch(/Get-Content|Test-Path/);
    });

    it('stops the host with the same sequence, waiting on its handle first', () => {
        const stop = fn('Stop-OracleHost');
        const order = [
            '$script:app.Quit(0)', '$script:app.Quit(2)', '$script:app.Quit()',
            'Remove-ComObjectReference $script:app', '[GC]::Collect()', '[GC]::WaitForPendingFinalizers()',
            "Where-Object { $_.ProcessName -eq $script:hostImageOfApp }", 'Stop-Process -Force',
            '$deadline = [DateTime]::UtcNow.AddSeconds(10)', '$script:appProcess.WaitForExit(10000)',
            'while ([DateTime]::UtcNow -lt $deadline -and $null -ne (Get-Process -Id $script:appPid',
            'did not exit.', '$script:appProcess.Dispose()', '$script:appPid = 0',
        ].map((text) => {
            expect(stop, text).toContain(text);
            return stop.indexOf(text);
        });
        expect(order).toEqual([...order].sort((a, b) => a - b));
        // The handle is taken once the host's pid is known.
        const start = fn('Start-OracleHost');
        expect(start.indexOf('$script:appPid = $processId')).toBeLessThan(start.indexOf('[Diagnostics.Process]::GetProcessById($processId)'));
    });

    it('skips work whose answer cannot change what happens', () => {
        // Only Word and PowerPoint use the snapshot of running instances.
        const start = fn('Start-OracleHost');
        expect(start).toMatch(/if \(\$script:hostName -eq "word" -or \$script:hostName -eq "powerpoint"\) \{\n\s+\$before = Get-HostProcessIds\n/);
        expect(start.split('Get-HostProcessIds')).toHaveLength(3);
        // What is still open is asked only of an instance that would be kept.
        const cleanup = worker.slice(worker.indexOf('$evidence = $result.outcome -eq'), worker.indexOf('# A loaded add-in holds its file'));
        expect(cleanup).toContain('$keep = $evidence -and $mode -eq "compile" -and $script:hostName -ne "access" -and $script:appCases -lt $CasesPerInstance');
        expect(cleanup).toContain('if ($keep -and $closed -and $null -ne $script:app) {');
        expect(cleanup).toContain('if ($keep -and $closed -and $script:providerPath -and $script:hostName -eq "powerpoint") {');
        expect(cleanup).toContain('if (-not $closed -or -not $keep) {\n            Stop-OracleHost');
        expect(cleanup.indexOf('$keep = $evidence')).toBeLessThan(cleanup.indexOf('Get-OpenContainerCount'));
        // The stage and the host record are written by .NET, not Set-Content.
        expect(fn('Set-Stage')).toContain('[IO.File]::WriteAllText($StagePath, $Name, [Text.Encoding]::ASCII)');
        expect(fn('Write-HostRecord')).toContain('[IO.File]::WriteAllText($PidPath,');
        // A served case travels in the request; a request without one is no verdict.
        expect(coordinator).not.toContain('case.json');
        expect(coordinator).toContain('case: kase, pidPath, stagePath, dialogPath,');
        expect(worker).toContain('if ($null -eq $request.case) {\n                throw "The request carries no case."\n            }\n            $output = @(Invoke-OracleCase $request.case)');
    });

    describe.skipIf(process.platform !== 'win32')('on Windows', () => {
        // One PowerShell for every check here. The dialog is a real message
        // box of this PowerShell, titled and worded as the VBE's are; the host
        // is a ping.exe this PowerShell starts.
        let seen: Record<string, unknown>;
        const source = 'Caf\u00e9 \u65e5\u672c \ud83d\ude00 "quoted" back\\slash\r\nnext line';
        const kase = { id: 'request_round_trip', source, modules: [{ name: 'M\u00f6dule', type: 'standard', source }] };
        let tmp = '';

        beforeAll(() => {
            tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'xlide-oracle-waits-'));
            // A case file as single-case mode reads it, and the same case as
            // a served request line, escaped as the coordinator escapes it.
            const casePath = path.join(tmp, 'case.json');
            fs.writeFileSync(casePath, JSON.stringify(kase), 'utf8');
            const line = JSON.stringify({ case: kase }).replace(/[\u0080-\uffff]/g, (ch) => `\\u${ch.charCodeAt(0).toString(16).padStart(4, '0')}`);
            const watcherPath = path.join(tmp, 'watcher.cs');
            fs.writeFileSync(watcherPath, watcherSource, 'utf8');
            const script = String.raw`
$ErrorActionPreference = 'Stop'
Add-Type -TypeDefinition ((Get-Content -LiteralPath $env:XLIDE_WATCHER_PATH -Raw -Encoding UTF8) + @'

public static class XlideTestBox {
    [System.Runtime.InteropServices.DllImport("user32.dll", CharSet = System.Runtime.InteropServices.CharSet.Unicode)]
    private static extern int MessageBoxW(System.IntPtr owner, string text, string caption, uint type);
    public static System.Threading.Thread Show(string text, string caption) {
        System.Threading.Thread thread = new System.Threading.Thread(delegate() { MessageBoxW(System.IntPtr.Zero, text, caption, 0); });
        thread.IsBackground = true;
        thread.Start();
        return thread;
    }
}
'@)
` + ['Write-ResultJson', 'Remove-ComObjectReference', 'Read-DialogResult', 'Wait-ForSeenDialogResult', 'Stop-OracleHost'].map(fn).join('\n') + String.raw`
$out = [ordered]@{}
$request = [Console]::In.ReadLine() | ConvertFrom-Json
$fromFile = Get-Content -LiteralPath $env:XLIDE_CASE_PATH -Raw -Encoding UTF8 | ConvertFrom-Json
$out.requestSource = $request.case.source
$out.requestModule = $request.case.modules[0].name
$out.fileSource = $fromFile.source

$dir = $env:XLIDE_TMP
$DialogPath = Join-Path $dir 'vbe_dialog.json'
$StagePath = Join-Path $dir 'stage.txt'
$DialogHoldSeconds = 0
[XlideVbeDialogWatcher]::Start([uint32]$PID, $DialogPath, $StagePath, 10, 0)
$out.beforeDialog = Read-DialogResult
$box = [XlideTestBox]::Show("Compile error:` + '`n`n' + String.raw`Expected: end of statement", 'Microsoft Visual Basic for Applications')
$out.dialog = Wait-ForSeenDialogResult
$out.seen = [XlideVbeDialogWatcher]::DialogSeen
$out.boxClosed = $box.Join(5000)
[XlideVbeDialogWatcher]::Stop()
$out.file = Get-Content -LiteralPath $DialogPath -Raw -Encoding UTF8 | ConvertFrom-Json
$out.stage = [IO.File]::ReadAllText($StagePath)
# Another case's path is told nothing, and is not kept waiting.
$DialogPath = Join-Path $dir 'other.json'
$timer = [Diagnostics.Stopwatch]::StartNew()
$out.otherPath = try { Wait-ForSeenDialogResult; 'returned' } catch { $_.Exception.Message }
$out.otherPathMs = $timer.Elapsed.TotalMilliseconds
# A new start forgets the last dialog.
$DialogPath = Join-Path $dir 'vbe_dialog.json'
[XlideVbeDialogWatcher]::Start([uint32]$PID, $DialogPath, $StagePath, 1, 0)
$out.afterRestart = Read-DialogResult
$out.seenAfterRestart = [XlideVbeDialogWatcher]::DialogSeen
[XlideVbeDialogWatcher]::Stop()

# Stop-OracleHost on a stand-in host: with the handle, and without it.
$script:app = $null
$script:hostNoun = 'Ping'
$script:hostImageOfApp = 'PING'
foreach ($held in @($true, $false)) {
    $child = Start-Process -FilePath ping.exe -ArgumentList '-n', '60', '127.0.0.1' -WindowStyle Hidden -PassThru
    $script:appPid = $child.Id
    $script:appProcess = $null
    if ($held) {
        $script:appProcess = [Diagnostics.Process]::GetProcessById($child.Id)
        $null = $script:appProcess.Handle
    }
    Stop-OracleHost
    $out["stop_$held"] = [ordered]@{
        gone = $null -eq (Get-Process -Id $child.Id -ErrorAction SilentlyContinue)
        exited = $child.HasExited
        appPid = $script:appPid
        handleReleased = $null -eq $script:appProcess
    }
}
Write-ResultJson ([pscustomobject]$out)
`;
            const run = spawnSync('powershell.exe', ['-NoProfile', '-NonInteractive', '-EncodedCommand', Buffer.from(script, 'utf16le').toString('base64')], {
                encoding: 'utf8',
                timeout: 60_000,
                input: `${line}\n`,
                env: { ...process.env, XLIDE_WATCHER_PATH: watcherPath, XLIDE_CASE_PATH: casePath, XLIDE_TMP: tmp },
            });
            expect(run.error).toBeUndefined();
            expect(run.status, run.stderr).toBe(0);
            seen = JSON.parse(run.stdout.trim().split(/\r?\n/).pop() ?? '');
        }, 70_000);

        afterAll(() => { if (tmp) { fs.rmSync(tmp, { recursive: true, force: true }); } });

        it('reads a served case from its request exactly as from a case file', () => {
            expect(seen.requestSource).toBe(source);
            expect(seen.fileSource).toBe(source);
            expect(seen.requestModule).toBe('M\u00f6dule');
        });

        it('records a dialog in memory exactly as in its file, for its own case only', () => {
            expect(seen.beforeDialog).toBeNull();
            expect(seen.dialog).toMatchObject({
                kind: 'vbe_compile_dialog',
                title: 'Microsoft Visual Basic for Applications',
                className: '#32770',
            });
            expect((seen.dialog as { message: string }).message).toMatch(/^Compile error:\s+Expected: end of statement$/);
            expect(seen.dialog).toEqual(seen.file);
            expect(seen.seen).toBe(true);
            expect(seen.boxClosed).toBe(true);
            expect(seen.stage).toBe('compile_dialog');
            expect(seen.otherPath).toBe('A VBE dialog appeared but was not dismissed and recorded.');
            expect(seen.otherPathMs as number).toBeLessThan(1000);
            expect(seen.afterRestart).toBeNull();
            expect(seen.seenAfterRestart).toBe(false);
        });

        it('stops a host it holds a handle to, and one it does not, the same way', () => {
            for (const held of ['True', 'False']) {
                expect(seen[`stop_${held}`], held).toEqual({ gone: true, exited: true, appPid: 0, handleReleased: true });
            }
        });
    });
});
