// The macro Security Notice, answered for the whole suite by the accepter
// officeGlobalSetup.ts starts: a load that raises one returns once the
// accepter has pressed Enable Macros, and a load that raises none waits for
// nothing. Each application gets both, in instances the check starts itself.
//
// What raises the notice, measured in Microsoft 365 (16.0 build 20527) from
// the temp folder: PowerPoint, for every add-in a project references, as the
// VBE oracle's PowerPoint references do; Excel and Word, for a file opened
// while the VBE is shown with AutomationSecurity 2 (by UI), and Word also for
// a template a project references while the VBE is shown, even with
// AutomationSecurity 1. Access raised none in any of those ways: an Access
// an automation client starts is at AutomationSecurity 2, and under it opens
// an untrusted database with its code disabled behind the message bar, with
// no notice (see openInAccess in projectReferences.office.ts).
//
// Access gets the loads the oracle makes, a database opened and another
// referenced, the second with the VBE shown and AutomationSecurity 2, and
// must raise no notice and wait for none.
//
// A notice shows only while the application's Trust Center macro setting is
// the default, "Disable with notification" (VBAWarnings 2): under another
// setting the check that needs one says so and is skipped.
//
// Word, PowerPoint and Access run only when not running already, so no
// instance of theirs is ever touched; Excel always runs in an instance of its
// own.

import * as fs from 'fs';
import { beforeAll, describe, expect, it } from 'vitest';
import type { OfficeHostApp } from '../../src/officeHostApps';
import { NOTICE_LOG, fields, hostInstalled, hostIsFree, psFile, q, scratchCopy, sleep } from './officeHarness';

type Load = { opened: string; seconds: number };

/** Per application: start it, load `quiet` the way that raises no notice, then `noticed` the way that does. */
const SCRIPTS: Record<OfficeHostApp, (quiet: string, noticed: string) => string[]> = {
    excel: (quiet, noticed) => [
        '$app = [Activator]::CreateInstance([type]::GetTypeFromProgID("Excel.Application"))',
        '$app.Visible = $true',
        '$app.DisplayAlerts = $false',
        '$app.AutomationSecurity = 1',
        ...timed('quiet', `$app.Workbooks.Open(${q(quiet)}).Name`),
        '$app.VBE.MainWindow.Visible = $true',
        '$app.AutomationSecurity = 2',
        ...timed('noticed', `$app.Workbooks.Open(${q(noticed)}).Name`),
        'foreach ($f in @($app.Workbooks)) { $f.Close($false) }',
        '$app.Quit()',
    ],
    word: (quiet, noticed) => [
        '$app = [Activator]::CreateInstance([type]::GetTypeFromProgID("Word.Application"))',
        '$app.Visible = $true',
        '$app.DisplayAlerts = 0',
        '$app.AutomationSecurity = 1',
        ...timed('quiet', `$app.Documents.Open(${q(quiet)}).Name`),
        '$app.VBE.MainWindow.Visible = $true',
        '$app.AutomationSecurity = 2',
        ...timed('noticed', `$app.Documents.Open(${q(noticed)}).Name`),
        'foreach ($f in @($app.Documents)) { $f.Close(0) }',
        '$app.Quit(0)',
    ],
    powerpoint: (quiet, noticed) => [
        '$app = [Activator]::CreateInstance([type]::GetTypeFromProgID("PowerPoint.Application"))',
        '$app.Visible = -1',
        '$app.AutomationSecurity = 1',
        // Presentations.Open(FileName, ReadOnly:=msoFalse, Untitled:=msoFalse, WithWindow:=msoTrue)
        ...timed('quiet', `$app.Presentations.Open(${q(quiet)}, 0, 0, -1).Name`),
        // A reference to an add-in, as Tools > References > Browse adds it,
        // from a project named apart from the add-in's VBAProject.
        '$consumer = $app.Presentations.Add(-1)',
        '$consumer.VBProject.Name = "NoticeConsumer"',
        ...timed('noticed', `$consumer.VBProject.References.AddFromFile(${q(noticed)}).Name`),
        'foreach ($f in @($app.Presentations)) { $f.Saved = -1; $f.Close() }',
        '$app.Quit()',
    ],
    // Measured to raise none: `noticed` is a second database, referenced.
    access: (quiet, noticed) => [
        '$app = [Activator]::CreateInstance([type]::GetTypeFromProgID("Access.Application"))',
        '$app.Visible = $true',
        '$app.AutomationSecurity = 1',
        ...timed('quiet', `$app.OpenCurrentDatabase(${q(quiet)}); $app.CurrentProject.Name`),
        '$app.VBE.MainWindow.Visible = $true',
        '$app.AutomationSecurity = 2',
        // Named apart, as the oracle names its projects, so the reference cannot clash.
        '$app.VBE.ActiveVBProject.Name = "NoticeConsumer"',
        ...timed('noticed', `$app.VBE.ActiveVBProject.References.AddFromFile(${q(noticed)}).Name`),
        'try { $app.CloseCurrentDatabase() } catch { }',
        '$app.Quit(2)',
    ],
};

/** The fixture each application loads both ways, and the copy's file name. */
const FIXTURES: Record<OfficeHostApp, { quiet: string; noticed: string }> = {
    excel: { quiet: 'ProbeLibraryFixture.xlsm', noticed: 'ProbeLibraryFixture.xlsm' },
    word: { quiet: 'WordFixture.docm', noticed: 'WordFixture.docm' },
    powerpoint: { quiet: 'PowerPointFixture.pptm', noticed: 'PowerPointFixture.ppam' },
    access: { quiet: 'AccessFixture.accdb', noticed: 'AccessTwelveModulesFixture.accdb' },
};

/** The Trust Center macro setting each application is under: the machine's policy, then the user's policy, then the user's own choice. */
const SETTING_KEYS: Record<OfficeHostApp, string> = { excel: 'Excel', word: 'Word', powerpoint: 'PowerPoint', access: 'Access' };
const settingLines = (host: OfficeHostApp): string[] => [
    '$warnings = $null',
    `foreach ($key in @("HKLM:\\Software\\Policies\\Microsoft\\Office\\16.0\\${SETTING_KEYS[host]}\\Security", "HKCU:\\Software\\Policies\\Microsoft\\Office\\16.0\\${SETTING_KEYS[host]}\\Security", "HKCU:\\Software\\Microsoft\\Office\\16.0\\${SETTING_KEYS[host]}\\Security")) {`,
    '    if ($null -eq $warnings) { $warnings = (Get-ItemProperty -LiteralPath $key -Name VBAWarnings -ErrorAction SilentlyContinue).VBAWarnings }',
    '}',
    '"vbaWarnings=" + $warnings',
];

/** Lines that run `expression` and print `name=<result>` and how long it took. */
function timed(name: string, expression: string): string[] {
    return [
        '$watch = [Diagnostics.Stopwatch]::StartNew()',
        `try { "${name}=" + $(${expression}) } catch { "${name}=failed: " + ($_.Exception.Message -replace '\\s+', ' ') }`,
        `"${name}Seconds=" + $watch.Elapsed.TotalSeconds.ToString([Globalization.CultureInfo]::InvariantCulture)`,
    ];
}

/**
 * The accepter's lines that name `file`, once it has had a moment to write
 * them: by its path as given, its short (8.3) path, or its own folder and
 * name, which a long form of a partly short path also ends with.
 */
async function noticesAbout(file: string, short: string | undefined): Promise<string[]> {
    const parts = file.split('\\');
    const tail = '\\' + parts.slice(-2).join('\\');
    const forms = [file, short, tail].filter((form): form is string => !!form).map((form) => form.toLowerCase());
    const read = () => fs.existsSync(NOTICE_LOG)
        ? fs.readFileSync(NOTICE_LOG, 'utf8').replace(/^\uFEFF/, '').split(/\r?\n/)
            .filter((line) => forms.some((form) => line.toLowerCase().includes(form)))
        : [];
    for (let i = 0; i < 20 && read().length === 0; i++) { await sleep(100); }
    return read();
}

for (const host of ['excel', 'word', 'powerpoint', 'access'] as const) {
    describe(`${host}: the macro Security Notice`, () => {
        let ran = false;
        let quiet = '';
        let noticed = '';
        let loads: Record<'quiet' | 'noticed', Load>;
        let err: string[] = [];
        let vbaWarnings = '';
        let shortPaths: Record<'quiet' | 'noticed', string | undefined> = { quiet: undefined, noticed: undefined };

        beforeAll(async () => {
            ran = host === 'excel' ? await hostInstalled(host) : await hostIsFree(host);
            if (!ran) { return; }
            // Named apart, since Excel cannot hold two workbooks of one name.
            const extension = (fixture: string) => fixture.slice(fixture.lastIndexOf('.'));
            quiet = scratchCopy(FIXTURES[host].quiet, `notice-${host}-quiet`, `Quiet${extension(FIXTURES[host].quiet)}`);
            noticed = scratchCopy(FIXTURES[host].noticed, `notice-${host}-noticed`, `Noticed${extension(FIXTURES[host].noticed)}`);
            const result = await psFile([
                '$ErrorActionPreference = "Continue"',
                ...settingLines(host),
                '$files = New-Object -ComObject Scripting.FileSystemObject',
                `"quietShort=" + $files.GetFile(${q(quiet)}).ShortPath`,
                `"noticedShort=" + $files.GetFile(${q(noticed)}).ShortPath`,
                ...SCRIPTS[host](quiet, noticed),
                '$app = $null',
                '[GC]::Collect()',
            ].join('\r\n'), `notice-${host}.ps1`, 180_000);
            const seen = fields(result.out);
            err = result.err;
            vbaWarnings = seen.vbaWarnings ?? '';
            shortPaths = { quiet: seen.quietShort, noticed: seen.noticedShort };
            loads = {
                quiet: { opened: seen.quiet, seconds: Number(seen.quietSeconds) },
                noticed: { opened: seen.noticed, seconds: Number(seen.noticedSeconds) },
            };
        }, 200_000);

        it('loads a file that raises no notice, and answers nothing', async (context) => {
            if (!ran) { context.skip(); }
            expect(loads.quiet.opened, err.join('\n')).not.toMatch(/^failed|^$/);
            expect(loads.quiet.seconds).toBeLessThan(15);
            expect(await noticesAbout(quiet, shortPaths.quiet)).toEqual([]);
        });

        it.skipIf(host === 'access')('loads a file that raises one, once the accepter has pressed Enable Macros', async (context) => {
            if (!ran) { context.skip(); }
            if (vbaWarnings !== '' && vbaWarnings !== '2') {
                // Not the default "Disable with notification": no notice to answer here.
                context.skip(`${host}'s macro setting is VBAWarnings ${vbaWarnings}, under which it shows no notice`);
            }
            expect(loads.noticed.opened, err.join('\n')).not.toMatch(/^failed|^$/);
            const lines = await noticesAbout(noticed, shortPaths.noticed);
            // A notice is reported again when it is answered on a later try.
            const pressed = lines.filter((line) => / \| pressed Enable Macros \| /.test(line));
            expect(pressed, lines.join('\n')).toHaveLength(1);
            expect(pressed[0]).toMatch(/^Microsoft (Excel|Word|PowerPoint) Security Notice \| /);
            expect(loads.noticed.seconds).toBeLessThan(15);
        });

        it.runIf(host === 'access')('references another database, and raises no notice to answer', async (context) => {
            if (!ran) { context.skip(); }
            expect(loads.noticed.opened, err.join('\n')).not.toMatch(/^failed|^$/);
            expect(loads.noticed.seconds).toBeLessThan(15);
            expect(await noticesAbout(noticed, shortPaths.noticed)).toEqual([]);
        });
    });
}
