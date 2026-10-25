// The formatter's casing of conditional-compilation names, checked by each
// application itself: a project whose `#If` and `#Const` lines are written in
// the wrong case is written into a file, the file is opened, and what the VBE
// shows for each module must be what formatting the same source gives, apart
// from indentation, which the VBE leaves as it was written.
//
// Excel and Access run in instances the check starts itself. Word and
// PowerPoint only when neither is running, so no instance of theirs is ever
// touched.

import * as fs from 'fs';
import * as path from 'path';
import { fileURLToPath } from 'url';
import { describe, expect, it } from 'vitest';
import { conditionalConstantDeclarations, conditionalNameCase } from '../../src/analyzer/format/conditionalNameCase';
import { formatVbaModule } from '../../src/analyzer/format/formatModule';
import { openMacroContainer } from '../../src/vba/macroContainer';
import { compress, decompress } from '../../src/vba/ovba';
import { dirRecord, readDirRecords } from '../../src/vba/vbaProject';
import { writeModule } from '../../src/vba/projectService';
import type { OfficeHostApp } from '../../src/officeHostApps';
import { hostInstalled, hostIsFree, psFile, q, runningCount, SCRATCH, sleep } from './officeHarness';

const TEMPLATES = fileURLToPath(new URL('../../assets/templates/', import.meta.url));

const PROJECT_ARGUMENTS = 'ProjArg = 1 : PCase = 1';

/** In project order: a name declared in two modules takes the later one's spelling. */
const MODULES: { name: string; source: string[] }[] = [
    {
        name: 'DirectiveA',
        source: [
            '#Const SharedFlag = 1',
            '#If vba7 And Not mac Then',
            '#ElseIf wIN64 Or win32 Or WIN16 Or vba6 Then',
            '#End If',
            '#If sharedflag Or projarg Then',
            '#End If',
            '#If False Then',
            '#if lateflag then',
            '#end if',
            '#End If',
            '#Const lateflag = 1',
            '#if twinbasic then',
            '#end if',
            'Sub UseThem()',
            '#If LATEFLAG Then',
            '    Debug.Print codeuse, win64',
            '#End If',
            'End Sub',
        ],
    },
    {
        name: 'DirectiveB',
        source: [
            '#Const SHAREDFLAG = 1',
            '#Const CodeUse = 1',
            '#const opa=1',
            '#const opb = "abc"',
            '#if opa=1 and opb="abc" then',
            '#elseif not opa or opa<>2 xor opa>=1 then',
            '#elseif opa mod 2 = 1 imp true then',
            '#elseif -opa then',
            '#else',
            '#end if',
            "#If vba7 _",
            "    And sharedflag Then ' trailing comment vba7",
            '#End If',
        ],
    },
    {
        name: 'DirectiveC',
        source: [
            '#If pcase Or PROJARG Then',
            '#End If',
            '#Const PCASE = 2',
            '# _',
            'Const ContName = 1',
            '#Const RemNote = 1 Rem note',
            '#Const ColonEnd = 1:',
            'Sub AfterStatement()',
            '    x = 1: #Const AfterColon = 1',
            'End Sub',
            '#If contname Or remnote Or colonend Or aftercolon Then',
            '#End If',
        ],
    },
];

/**
 * Sets the project's Conditional Compilation Arguments, which XLIDE reads but
 * never writes: the two PROJECTCONSTANTS records of the dir stream (MS-OVBA
 * 2.3.4.2.1.12).
 */
function setProjectArguments(file: string, text: string): void {
    const container = openMacroContainer(fs.readFileSync(file));
    const cfb = container.vbaCfb();
    let storage: string | undefined = 'VBA';
    let raw: Buffer;
    try {
        raw = cfb.getStreamInStorage('VBA', 'dir');
    } catch {
        raw = cfb.getStream('dir');
        storage = undefined;
    }
    const dir = decompress(raw);
    const records = readDirRecords(dir);
    const ansi = records.find((record) => record.id === 0x000c);
    const unicode = records.find((record) => record.id === 0x003c);
    if (!ansi || !unicode || unicode.start !== ansi.end) {
        throw new Error(`${file}: no PROJECTCONSTANTS records to replace`);
    }
    cfb.setStreamIn(storage, 'dir', compress(Buffer.concat([
        dir.subarray(0, ansi.start),
        dirRecord(0x000c, Buffer.from(text, 'latin1')),
        dirRecord(0x003c, Buffer.from(text, 'utf16le')),
        dir.subarray(unicode.end),
    ])));
    fs.writeFileSync(file, container.toFileBytes(cfb));
}

/** Each module formatted, its lines without their indentation. */
function formatted(): Map<string, string[]> {
    const sources = MODULES.map((module) => module.source.join('\r\n') + '\r\n');
    const identifierCase = conditionalNameCase(sources.map(conditionalConstantDeclarations), PROJECT_ARGUMENTS);
    return new Map(MODULES.map((module, index) => {
        const result = formatVbaModule(sources[index], { tabSize: 4, insertSpaces: true, identifierCase });
        expect(result.refusal).toBeUndefined();
        return [module.name, unindented(result.text as string)];
    }));
}

function unindented(text: string): string[] {
    return text.split(/\r\n|\r|\n/).map((line) => line.trimStart()).filter((line) => line !== '');
}

interface Host {
    host: OfficeHostApp;
    template: string;
    available: () => Promise<boolean>;
    /** PowerShell that opens `$file` with macros off and sets `$app`, `$doc` and `$project`. */
    open: string[];
    /** Closes the file, then quits; each runs only once `$app` exists, and a failure in one does not stop the other. */
    close: string;
    quit: string;
}

const HOSTS: Host[] = [
    {
        host: 'excel',
        template: 'blank.xlsm',
        available: () => hostInstalled('excel'),
        open: [
            '$app = [Activator]::CreateInstance([type]::GetTypeFromProgID("Excel.Application"))',
            '$app.DisplayAlerts = $false',
            // msoAutomationSecurityForceDisable
            '$app.AutomationSecurity = 3',
            '$doc = $app.Workbooks.Open($file, 0, $true)',
            '$project = $doc.VBProject',
        ],
        close: '$doc.Close($false)',
        quit: '$app.Quit()',
    },
    {
        host: 'word',
        template: 'blank.docm',
        available: () => hostIsFree('word'),
        open: [
            '$app = [Activator]::CreateInstance([type]::GetTypeFromProgID("Word.Application"))',
            '$app.DisplayAlerts = 0',
            '$app.AutomationSecurity = 3',
            '$doc = $app.Documents.Open($file, $false, $true)',
            '$project = $doc.VBProject',
        ],
        close: '$doc.Close(0)',
        quit: '$app.Quit(0)',
    },
    {
        host: 'powerpoint',
        template: 'blank.pptm',
        available: () => hostIsFree('powerpoint'),
        open: [
            '$app = [Activator]::CreateInstance([type]::GetTypeFromProgID("PowerPoint.Application"))',
            '$app.AutomationSecurity = 3',
            // Presentations.Open(FileName, ReadOnly:=msoTrue, Untitled:=msoFalse, WithWindow:=msoFalse)
            '$doc = $app.Presentations.Open($file, -1, 0, 0)',
            '$project = $doc.VBProject',
        ],
        close: '$doc.Close()',
        quit: '$app.Quit()',
    },
    {
        host: 'access',
        template: 'blank.accdb',
        available: () => hostInstalled('access'),
        open: [
            '$app = [Activator]::CreateInstance([type]::GetTypeFromProgID("Access.Application"))',
            '$app.AutomationSecurity = 3',
            '$app.OpenCurrentDatabase($file)',
            '$project = $app.VBE.ActiveVBProject',
        ],
        close: '$app.CloseCurrentDatabase()',
        quit: '$app.Quit()',
    },
];

/** Waits for the host's process count to fall back to `count`. */
async function settlesAt(host: OfficeHostApp, count: number, timeoutMs = 30_000): Promise<boolean> {
    const until = Date.now() + timeoutMs;
    while (Date.now() < until) {
        if (await runningCount(host) <= count) {
            return true;
        }
        await sleep(100);
    }
    return false;
}

for (const host of HOSTS) {
    describe.runIf(await host.available())(`conditional-compilation names, opened in ${host.host}`, () => {
        it('show in the VBE as the formatter cases them', async () => {
            const dir = path.join(SCRATCH, `directive-casing-${host.host}`);
            fs.rmSync(dir, { recursive: true, force: true });
            fs.mkdirSync(dir, { recursive: true });
            const file = path.join(dir, `Directives${path.extname(host.template)}`);
            fs.copyFileSync(path.join(TEMPLATES, host.template), file);
            for (const module of MODULES) {
                writeModule(file, module.name, module.source.join('\r\n') + '\r\n', 'standard');
            }
            setProjectArguments(file, PROJECT_ARGUMENTS);

            const before = await runningCount(host.host);
            const result = await psFile([
                '$ErrorActionPreference = "Stop"',
                `$file = ${q(file)}`,
                '$app = $null',
                '$doc = $null',
                'try {',
                ...host.open.map((line) => `  ${line}`),
                `  foreach ($name in @(${MODULES.map((module) => q(module.name)).join(', ')})) {`,
                '    $code = $project.VBComponents.Item($name).CodeModule',
                '    $text = if ($code.CountOfLines -gt 0) { $code.Lines(1, $code.CountOfLines) } else { "" }',
                '    "module=" + (@{ name = $name; text = $text } | ConvertTo-Json -Compress)',
                '  }',
                '} finally {',
                '  if ($app) {',
                `    try { ${host.close} } catch { }`,
                `    try { ${host.quit} } catch { }`,
                '  }',
                '  $doc = $null',
                '  $project = $null',
                '  $app = $null',
                '  [GC]::Collect()',
                '  [GC]::WaitForPendingFinalizers()',
                '}',
            ].join('\r\n'), `directive-casing-${host.host}.ps1`);
            // Other checks skip a host that is busy, so one this check failed
            // to quit would make them skip without failing.
            expect(await settlesAt(host.host, before), `${host.host} did not quit`).toBe(true);
            const shown = new Map(result.out
                .filter((line) => line.startsWith('module='))
                .map((line) => JSON.parse(line.slice('module='.length)) as { name: string; text: string })
                .map((module) => [module.name, unindented(module.text)
                    .filter((line) => !/^Attribute VB_|^Option Compare Database$/.test(line))]));
            expect([...shown.keys()], result.err.join('\n')).toEqual(MODULES.map((module) => module.name));

            const expected = formatted();
            for (const module of MODULES) {
                expect(shown.get(module.name), module.name).toEqual(expected.get(module.name));
                // The check has something to check: the VBE changed what was written.
                expect(shown.get(module.name), module.name).not.toEqual(unindented(module.source.join('\n')));
            }
        });
    });
}
