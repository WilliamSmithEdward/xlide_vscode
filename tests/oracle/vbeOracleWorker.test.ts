import { spawnSync } from 'node:child_process';
import * as fs from 'fs';
import * as path from 'path';
import { describe, expect, it } from 'vitest';
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
        encoding: 'utf8', timeout: 15_000,
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
}, 20_000);

describe('the macro Security Notice', () => {
    it('is accepted for as long as the worker drives a host, whatever the case is doing', () => {
        const worker = readOracle('vbe_oracle_worker.ps1');
        const coordinator = readOracle('run_vbe_oracle.mjs');
        expect(worker).toContain(". (Join-Path $PSScriptRoot 'office_security_notice.ps1')");
        const start = worker.slice(worker.indexOf('function Start-OracleHost'), worker.indexOf('\nfunction ', worker.indexOf('function Start-OracleHost')));
        const stop = worker.slice(worker.indexOf('function Stop-OracleHost'), worker.indexOf('\nfunction ', worker.indexOf('function Stop-OracleHost')));
        expect(start).toContain('[XlideSecurityNotice]::Start([uint32]$processId)');
        expect(start.indexOf('[XlideSecurityNotice]::Start')).toBeGreaterThan(start.indexOf('$script:appPid = $processId'));
        // A notice still being answered is recorded before the case reads its notices.
        const settle = worker.indexOf('[void][XlideSecurityNotice]::Settle(5000)');
        expect(settle).toBeGreaterThan(0);
        expect(worker.indexOf('[XlideSecurityNotice]::Accepted | Select-Object -Skip $noticesBefore')).toBeGreaterThan(settle);
        expect(stop.trimStart().split('\n')[1].trim()).toBe('[XlideSecurityNotice]::Stop()');
        // A checkpoint made before a change to the accepter is not reused.
        expect(coordinator).toContain('.update(fs.readFileSync(SECURITY_NOTICE))');
        expect(readOracle('office_security_notice.ps1')).not.toMatch(/^\s*param\s*\(/m);
    });

    it('is accepted by the Excel formats check and the live Office suite too', () => {
        const root = path.join(__dirname, '..', '..');
        const formats = fs.readFileSync(path.join(root, 'scripts', 'verify_excel_formats_worker.ps1'), 'utf8');
        expect(formats).toContain("office_security_notice.ps1')");
        expect(formats).toContain('[XlideSecurityNotice]::Start([uint32]$ownPid)');
        const setup = fs.readFileSync(path.join(root, 'tests', 'office', 'officeGlobalSetup.ts'), 'utf8');
        expect(setup).toContain('[XlideSecurityNotice]::Poll(0, $Scratch)');
        // The suite's dialog watcher closes what it finds, which would disable
        // the file's code, so it leaves a notice to the accepter.
        const harness = fs.readFileSync(path.join(root, 'tests', 'office', 'officeHarness.ts'), 'utf8');
        expect(harness).toContain('title.EndsWith(" Security Notice")) { return true; }');
    });
});

describe.skipIf(process.platform !== 'win32')('the macro Security Notice, answered in a real window', () => {
    // One PowerShell shows stand-in notices from a thread of its own and
    // answers them with the accepter the harnesses load. No Office is needed.
    let results: Record<string, { clicked: string; lines: string[]; ms: number; open: boolean; shortDiffers?: boolean }>;

    it('runs every scenario', () => {
        const helper = path.join(ORACLE, 'office_security_notice.ps1');
        const script = String.raw`
$ErrorActionPreference = 'Stop'
. '${helper.replace(/'/g, "''")}'
Add-Type -ReferencedAssemblies System.Windows.Forms, System.Drawing -TypeDefinition @"
using System;
using System.Threading;
using System.Windows.Forms;
public static class StandInNotice {
    public static volatile string Clicked = "";
    private static Form form;
    private static Thread thread;
    public static void Show(string title, string path, string[] buttons) { Show(title, path, buttons, ""); }
    public static void Show(string title, string path, string[] buttons, string disabled) {
        Clicked = "";
        ManualResetEvent shown = new ManualResetEvent(false);
        thread = new Thread(delegate() {
            form = new Form();
            form.Text = title;
            form.Width = 520;
            TextBox file = new TextBox();
            file.ReadOnly = true;
            file.Text = path;
            file.Width = 480;
            form.Controls.Add(file);
            for (int i = 0; i < buttons.Length; i++) {
                Button button = new Button();
                button.Text = buttons[i];
                button.Top = 40;
                button.Left = 10 + i * 120;
                button.Width = 110;
                button.Enabled = buttons[i] != disabled;
                string name = buttons[i];
                button.Click += delegate { Clicked = name; form.Close(); };
                form.Controls.Add(button);
            }
            form.Shown += delegate { shown.Set(); };
            Application.Run(form);
        });
        thread.SetApartmentState(ApartmentState.STA);
        thread.IsBackground = true;
        thread.Start();
        shown.WaitOne(10000);
        Thread.Sleep(200);
    }
    public static bool Open { get { return thread != null && thread.IsAlive; } }
    public static void SetPath(string path) {
        form.Invoke(new MethodInvoker(delegate { form.Controls[0].Text = path; }));
    }
    public static void EnableAll() {
        form.Invoke(new MethodInvoker(delegate { foreach (Control control in form.Controls) { control.Enabled = true; } }));
    }
    public static bool WaitClosed(int ms) { return thread.Join(ms); }
    public static void Close() { if (Open) { form.Invoke(new MethodInvoker(form.Close)); thread.Join(5000); } }
}
"@
$results = [ordered]@{}
function Done($name, $lines, $ms) {
    $open = [StandInNotice]::Open
    [StandInNotice]::Close()
    $results[$name] = [ordered]@{ clicked = [StandInNotice]::Clicked; lines = @($lines); ms = [int]$ms; open = $open }
}

# Shown while a call waits: the background accepter presses Enable Macros.
# Timed from the notice being up, not from the stand-in's own start.
[XlideSecurityNotice]::Start([uint32]$PID)
[StandInNotice]::Show('Microsoft PowerPoint Security Notice', 'C:\Scratch\Lib.ppam', @('Disable Macros', 'Enable Macros'))
$watch = [Diagnostics.Stopwatch]::StartNew()
[void][StandInNotice]::WaitClosed(10000)
[XlideSecurityNotice]::Stop()
Done 'shown' ([XlideSecurityNotice]::Accepted) $watch.ElapsedMilliseconds

# None shown: settling and stopping the accepter wait for nothing.
[XlideSecurityNotice]::Start([uint32]$PID)
Start-Sleep -Milliseconds 300
$watch = [Diagnostics.Stopwatch]::StartNew()
$settled = [XlideSecurityNotice]::Settle(5000)
[XlideSecurityNotice]::Stop()
$none = @([XlideSecurityNotice]::Accepted) + @([XlideSecurityNotice]::Accept([uint32]$PID))
if (-not $settled) { $none += 'did not settle' }
Done 'none' $none $watch.ElapsedMilliseconds

# The same notice twice, one after the other: each is answered and reported.
$twice = @()
foreach ($i in 1, 2) {
    [StandInNotice]::Show('Microsoft Excel Security Notice', 'C:\Scratch\Same.xlsm', @('Enable Macros', 'Disable Macros'))
    $twice += [XlideSecurityNotice]::Poll(0, 'C:\Scratch')
    $twice += [XlideSecurityNotice]::Poll(0, 'C:\Scratch')
    [void][StandInNotice]::WaitClosed(5000)
}
Done 'twice' $twice 0

# A folder whose name only starts with the given one is another folder.
[StandInNotice]::Show('Microsoft Excel Security Notice', 'C:\Scratch-other\Book.xlsm', @('Enable Macros', 'Disable Macros'))
Done 'sibling' ([XlideSecurityNotice]::Accept(0, 'C:\Scratch')) 0

# A folder given in its short (8.3) form still owns the notice that names its long one.
$long = Join-Path $env:TEMP ('xlide notice folder ' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $long | Out-Null
$short = (New-Object -ComObject Scripting.FileSystemObject).GetFolder($long).ShortPath
[StandInNotice]::Show('Microsoft Excel Security Notice', (Join-Path $long 'Book.xlsm'), @('Enable Macros', 'Disable Macros'))
Done 'short form' ([XlideSecurityNotice]::Accept(0, $short)) 0
$results['short form'].shortDiffers = $short -ne $long

# A notice naming a partly short path, as under a temp folder such as
# C:\Users\RUNNER~1\AppData\Local\Temp, owns a folder given in another mix.
$inner = Join-Path $long ('xlide notice inner ' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $inner | Out-Null
$files = New-Object -ComObject Scripting.FileSystemObject
$shownFolder = Join-Path $files.GetFolder($long).ShortPath (Split-Path -Leaf $inner)
$givenFolder = Join-Path $long (Split-Path -Leaf $files.GetFolder($inner).ShortPath)
[StandInNotice]::Show('Microsoft Excel Security Notice', (Join-Path $shownFolder 'Book.xlsm'), @('Enable Macros', 'Disable Macros'))
Done 'mixed form' ([XlideSecurityNotice]::Accept(0, $givenFolder)) 0
$results['mixed form'].shortDiffers = $shownFolder -ne $inner
Remove-Item -LiteralPath $long -Recurse -Force

# A notice whose button cannot be pressed yet is reported once, and pressed
# once it can be.
[StandInNotice]::Show('Microsoft Excel Security Notice', 'C:\Scratch\Late.xlsm', @('Enable Macros', 'Disable Macros'), 'Enable Macros')
$late = @([XlideSecurityNotice]::Poll(0, 'C:\Scratch')) + @([XlideSecurityNotice]::Poll(0, 'C:\Scratch'))
[StandInNotice]::EnableAll()
$late += [XlideSecurityNotice]::Poll(0, 'C:\Scratch')
[void][StandInNotice]::WaitClosed(5000)
Done 'late' $late 0

# A notice read before its path is there is looked at again once it is.
[StandInNotice]::Show('Microsoft Excel Security Notice', '', @('Enable Macros', 'Disable Macros'))
$slow = @([XlideSecurityNotice]::Poll(0, 'C:\Scratch')) + @([XlideSecurityNotice]::Poll(0, 'C:\Scratch'))
[StandInNotice]::SetPath('C:\Scratch\Slow.xlsm')
$slow += [XlideSecurityNotice]::Poll(0, 'C:\Scratch')
[void][StandInNotice]::WaitClosed(5000)
Done 'path later' $slow 0

# A button that cannot be pressed is not reported as pressed.
[StandInNotice]::Show('Microsoft Word Security Notice', 'C:\Scratch\Lib.dotm', @('Enable Macros', 'Disable Macros'), 'Enable Macros')
Done 'disabled' ([XlideSecurityNotice]::Accept([uint32]$PID)) 0

# Access asks to Open.
[StandInNotice]::Show('Microsoft Access Security Notice', 'C:\Scratch\Lib.accdb', @('Open', 'Cancel'))
Done 'access' ([XlideSecurityNotice]::Accept([uint32]$PID)) 0

# A notice about a file outside the given folder is left alone.
[StandInNotice]::Show('Microsoft Excel Security Notice', 'C:\Elsewhere\Book.xlsm', @('Enable Macros', 'Disable Macros'))
Done 'elsewhere' ([XlideSecurityNotice]::Accept(0, 'C:\Scratch')) 0

# The same notice about a file inside it is answered, in whichever process shows it.
[StandInNotice]::Show('Microsoft Excel Security Notice', 'C:\Scratch\Book.xlsm', @('Enable Macros', 'Disable Macros'))
Done 'inside' ([XlideSecurityNotice]::Accept(0, 'C:\Scratch')) 0

# Another process's notice is not this one's.
[StandInNotice]::Show('Microsoft Word Security Notice', 'C:\Scratch\Lib.dotm', @('Enable Macros', 'Disable Macros'))
Done 'other process' ([XlideSecurityNotice]::Accept([uint32]($PID + 1))) 0

# A window that is not a Security Notice is never pressed.
[StandInNotice]::Show('Microsoft Visual Basic for Applications', 'C:\Scratch\Lib.ppam', @('Open', 'Enable Macros'))
Done 'other dialog' ([XlideSecurityNotice]::Accept([uint32]$PID)) 0

# A notice with no button it knows is reported and left for the timeout.
[StandInNotice]::Show('Microsoft PowerPoint Security Notice', 'C:\Scratch\Lib.ppam', @('Trust', 'Distrust'))
Done 'unknown buttons' ([XlideSecurityNotice]::Accept([uint32]$PID)) 0

ConvertTo-Json -InputObject $results -Depth 4 -Compress
`;
        const run = spawnSync('powershell.exe', ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-STA', '-EncodedCommand', Buffer.from(script, 'utf16le').toString('base64')], {
            encoding: 'utf8', timeout: 60_000,
        });
        expect(run.error).toBeUndefined();
        expect(run.status, run.stderr).toBe(0);
        results = JSON.parse(run.stdout.trim());
    }, 70_000);

    it('presses Enable Macros when a notice comes while a call waits, and says what it pressed and for which file', () => {
        const shown = results.shown;
        expect(shown.clicked).toBe('Enable Macros');
        expect(shown.open).toBe(false);
        expect(shown.lines).toHaveLength(1);
        expect(shown.lines[0]).toMatch(/^Microsoft PowerPoint Security Notice \| pressed Enable Macros \| .*C:\\Scratch\\Lib\.ppam/);
        expect(shown.ms).toBeLessThan(5000);
    });

    it('waits for nothing when no notice comes', () => {
        expect(results.none.lines).toEqual([]);
        expect(results.none.ms).toBeLessThan(1000);
    });

    it("presses Open, the button Access's notice has", () => {
        expect(results.access.clicked).toBe('Open');
        expect(results.access.lines[0]).toMatch(/^Microsoft Access Security Notice \| pressed Open \| /);
    });

    it('answers only notices about files in the folder it is given', () => {
        expect(results.elsewhere.clicked).toBe('');
        expect(results.elsewhere.open).toBe(true);
        expect(results.elsewhere.lines).toEqual([expect.stringMatching(/^Microsoft Excel Security Notice \| left alone, not about C:\\Scratch \| /)]);
        expect(results.inside.clicked).toBe('Enable Macros');
        expect(results.inside.open).toBe(false);
    });

    it('touches neither another process nor any other dialog', () => {
        for (const name of ['other process', 'other dialog']) {
            expect(results[name].clicked, name).toBe('');
            expect(results[name].open, name).toBe(true);
            expect(results[name].lines, name).toEqual([]);
        }
    });

    it('answers and reports a second notice with the same words as the first', () => {
        expect(results.twice.lines).toHaveLength(2);
        for (const line of results.twice.lines) {
            expect(line).toMatch(/^Microsoft Excel Security Notice \| pressed Enable Macros \| .*C:\\Scratch\\Same\.xlsm/);
        }
    });

    it('takes a folder whose name only starts with the given one for another folder', () => {
        expect(results.sibling.clicked).toBe('');
        expect(results.sibling.lines).toEqual([expect.stringMatching(/ \| left alone, not about C:\\Scratch \| /)]);
    });

    it('matches the folder in its short or long form', () => {
        expect(results['short form'].clicked).toBe('Enable Macros');
        expect(results['short form'].lines).toEqual([expect.stringMatching(/ \| pressed Enable Macros \| /)]);
    });

    it('tries a notice again until it can press the button, reporting the first try once', () => {
        expect(results.late.clicked).toBe('Enable Macros');
        expect(results.late.lines).toEqual([
            expect.stringMatching(/^Microsoft Excel Security Notice \| could not press Enable Macros \| /),
            expect.stringMatching(/^Microsoft Excel Security Notice \| pressed Enable Macros \| /),
        ]);
    });

    it('looks again at a notice it left alone, once its path names the folder', () => {
        expect(results['path later'].clicked).toBe('Enable Macros');
        expect(results['path later'].lines).toEqual([
            expect.stringMatching(/^Microsoft Excel Security Notice \| left alone, not about C:\\Scratch \| /),
            expect.stringMatching(/^Microsoft Excel Security Notice \| pressed Enable Macros \| .*C:\\Scratch\\Slow\.xlsm/),
        ]);
    });

    it('matches a path whose folders are partly short against a folder given in another mix', () => {
        // Where the disk keeps no 8.3 names, the mix is the long path itself.
        expect(results['mixed form'].clicked).toBe('Enable Macros');
        expect(results['mixed form'].lines).toEqual([expect.stringMatching(/ \| pressed Enable Macros \| /)]);
    });

    it('says when it could not press the button', () => {
        expect(results.disabled.clicked).toBe('');
        expect(results.disabled.open).toBe(true);
        expect(results.disabled.lines).toEqual([expect.stringMatching(/^Microsoft Word Security Notice \| could not press Enable Macros \| /)]);
    });

    it('reports a notice whose buttons it does not know, and leaves it', () => {
        const unknown = results['unknown buttons'];
        expect(unknown.clicked).toBe('');
        expect(unknown.open).toBe(true);
        expect(unknown.lines).toEqual([expect.stringMatching(/^Microsoft PowerPoint Security Notice \| no known button \| /)]);
    });
});
