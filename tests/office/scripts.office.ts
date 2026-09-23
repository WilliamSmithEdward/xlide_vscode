// Every PowerShell script XLIDE can generate, parsed by PowerShell itself.
//
// Unit tests read these scripts as text, and text that looks right can still
// be a script PowerShell reads differently: the lines used to be joined with
// "; ", and an `else` after that separator parsed cleanly - as a command
// called else - and stopped the script where it ran. Only a real Word caught
// it. Needs Windows PowerShell and nothing from Office.

import * as fs from 'fs';
import * as path from 'path';
import { fileURLToPath } from 'url';
import { beforeAll, describe, expect, it } from 'vitest';
import {
    buildAccessMacroLaunchScript,
    buildAccessShowDesignScript,
    buildExcelLaunchScript,
    buildHostOpenScript,
    buildPowerPointMacroLaunchScript,
    buildWordMacroLaunchScript,
} from '../../src/officeHostLauncher';
import { buildCloseFileScript, buildRefreshReadOnlyScript } from '../../src/officeWriteCoordinator';
import { officeComProbePowerShellScript } from '../../src/officeComAvailability';
import { buildLockHoldersScript } from '../../src/fileLockHolders';
import { buildOwnedReadOnlyTestHostScript } from '../../src/vbaTestOfficeHost';
import { setExtensionAssetRoot } from '../../src/extensionAssets';
import type { OfficeHostApp } from '../../src/officeHostApps';
import type { OfficeViewPlace } from '../../src/officeViewPlace';
import { ps, psFile, q, SCRATCH } from './officeHarness';

const HOSTS: readonly OfficeHostApp[] = ['excel', 'word', 'powerpoint', 'access'];
const FILES: Record<OfficeHostApp, string> = {
    excel: "C:\\w\\Bob's Book.xlsm",
    word: 'C:\\w\\Report.docm',
    powerpoint: 'C:\\w\\Deck.pptm',
    access: 'C:\\w\\Orders.accdb',
};
const PLACES: Record<OfficeHostApp, OfficeViewPlace> = {
    excel: { sheet: "Q3 'Actuals'", selection: '$B$2:$C$4', activeCell: '$C$3', scrollRow: 40, scrollColumn: 2, otherActive: 'C:\\w\\Other.xlsx' },
    word: { start: 120, end: 131, scrolled: 42 },
    powerpoint: { slide: 7 },
    access: {},
};

/** Every script, by a name that says which builder and options made it. */
function everyScript(): Array<[string, string]> {
    const scripts: Array<[string, string]> = [];
    for (const host of HOSTS) {
        for (const readOnly of [false, true]) {
            for (const attachToRunning of [false, true]) {
                for (const background of [false, true]) {
                    scripts.push([`open-${host}-ro${readOnly}-attach${attachToRunning}-bg${background}`,
                        buildHostOpenScript({ host, filePath: FILES[host], attachToRunning, readOnly, background })]);
                }
            }
            scripts.push([`open-${host}-ro${readOnly}-place`,
                buildHostOpenScript({ host, filePath: FILES[host], attachToRunning: true, readOnly, background: true, place: PLACES[host] })]);
        }
        for (const force of [false, true]) {
            for (const onlyReadOnlyCopy of [false, true]) {
                scripts.push([`close-${host}-force${force}-ro${onlyReadOnlyCopy}`, buildCloseFileScript(FILES[host], { force, onlyReadOnlyCopy })]);
            }
        }
        const refresh = buildRefreshReadOnlyScript(FILES[host]);
        if (refresh) {
            scripts.push([`refresh-${host}`, refresh]);
        }
        scripts.push([`com-probe-${host}`, officeComProbePowerShellScript(host)]);
        scripts.push([`test-host-${host}`, buildOwnedReadOnlyTestHostScript(FILES[host], [], { hostApp: host })]);
    }
    scripts.push(['f5-excel', buildExcelLaunchScript({ filePath: FILES.excel, attachToRunning: true, mode: { kind: 'macroReadOnly', macroName: 'M.Go' } })]);
    scripts.push(['f5-word', buildWordMacroLaunchScript(FILES.word, 'M.Go')]);
    scripts.push(['f5-powerpoint', buildPowerPointMacroLaunchScript(FILES.powerpoint, 'M.Go')]);
    scripts.push(['f5-access', buildAccessMacroLaunchScript(FILES.access, 'Go')]);
    scripts.push(['f5-access-form', buildAccessShowDesignScript(FILES.access, { kind: 'form', name: 'Orders' })]);
    scripts.push(['f5-access-report', buildAccessShowDesignScript(FILES.access, { kind: 'report', name: 'Sales' })]);
    scripts.push(['lock-holders', buildLockHoldersScript(FILES.excel)]);
    return scripts;
}

/**
 * Parses each file with PowerShell's parser and reports, per file, its parse
 * errors and any `else`, `elseif`, `catch` or `finally` the parser took for a
 * command - the shape that parses without an error and fails as it runs.
 */
const PARSE_CHECKER = [
    'param([string]$dir)',
    '$report = @()',
    'foreach ($file in Get-ChildItem $dir -Filter *.ps1) {',
    '  $tokens = $null',
    '  $errors = $null',
    '  $ast = [System.Management.Automation.Language.Parser]::ParseFile($file.FullName, [ref]$tokens, [ref]$errors)',
    '  $orphans = @($ast.FindAll({ param($node) $node -is [System.Management.Automation.Language.CommandAst] -and @("else", "elseif", "catch", "finally") -contains $node.GetCommandName() }, $true))',
    '  $report += [pscustomobject]@{ name = $file.BaseName; errors = @($errors | ForEach-Object { $_.Message }); orphans = $orphans.Count }',
    '}',
    '[Console]::Out.WriteLine("XLIDE_PARSE|" + (ConvertTo-Json -InputObject @($report) -Compress -Depth 4))',
].join('\r\n');

async function parseAll(scripts: ReadonlyArray<[string, string]>, folder: string): Promise<Array<{ name: string; errors: string[]; orphans: number }>> {
    const dir = path.join(SCRATCH, folder);
    fs.rmSync(dir, { recursive: true, force: true });
    fs.mkdirSync(dir, { recursive: true });
    for (const [name, script] of scripts) {
        fs.writeFileSync(path.join(dir, `${name}.ps1`), script);
    }
    const checker = path.join(SCRATCH, 'parse-checker.ps1');
    fs.writeFileSync(checker, PARSE_CHECKER);
    const result = await psFile(`& ${q(checker)} -dir ${q(dir)}`, `run-${folder}.ps1`);
    const line = result.out.find((candidate) => candidate.startsWith('XLIDE_PARSE|'));
    expect(line, result.err.join('\n')).toBeDefined();
    return JSON.parse(line!.slice('XLIDE_PARSE|'.length)) as Array<{ name: string; errors: string[]; orphans: number }>;
}

describe('the scripts XLIDE sends to PowerShell', () => {
    beforeAll(() => {
        setExtensionAssetRoot(fileURLToPath(new URL('../../', import.meta.url)));
    });

    it('catches the bug that started this - the anchor for the check below', async () => {
        // The exact shape that shipped, and its fix. A checker that cannot
        // tell these apart proves nothing about the scripts.
        const report = await parseAll([
            ['old-join', '$x = $true; if ($x) { "a" };   else { "b" }'],
            ['newline', '$x = $true\nif ($x) { "a" }\nelse { "b" }'],
        ], 'anchor');
        const byName = new Map(report.map((entry) => [entry.name, entry]));
        expect(byName.get('old-join')).toMatchObject({ errors: [], orphans: 1 });
        expect(byName.get('newline')).toMatchObject({ errors: [], orphans: 0 });
    });

    it('all parse, and none runs a clause as a command', async () => {
        const scripts = everyScript();
        const report = await parseAll(scripts, 'generated');
        expect(report).toHaveLength(scripts.length);
        const bad = report.filter((entry) => entry.errors.length > 0 || entry.orphans > 0);
        expect(bad).toEqual([]);
    });

    it('arrive whole, with their own line breaks, as the product sends them', async () => {
        const result = await ps([
            '$x = $false',
            'if ($x) {',
            '  "then"',
            '}',
            'else {',
            '  "else ran"',
            '}',
            'try {',
            '  throw "boom"',
            '}',
            'catch {',
            '  "caught " + $_.Exception.Message',
            '}',
            '"text " + [char]0x00E9 + [char]0x0416',
            'exit 3',
        ].join('\n'));
        expect(result.out).toEqual(['else ran', 'caught boom', 'text \u00e9\u0416']);
        expect(result.code).toBe(3);
    });
});
