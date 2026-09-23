// "Add VBA Project", checked by the applications themselves: each file XLIDE
// adds a project to is opened in its application with alerts on, beside a
// watcher that records and closes any dialog the application raises; its
// probe module - which reaches the new document modules by code name - is run;
// and the application saves the file, which XLIDE then reads back.
//
// Excel runs in instances the check starts itself. Word and PowerPoint only
// when neither is running, so no instance of theirs is ever touched.

import * as fs from 'fs';
import * as path from 'path';
import { describe, expect, it } from 'vitest';
import { fileURLToPath } from 'url';
import { addVbaProject, hasVbaProject, listModules, writeModule } from '../../src/vba/projectService';
import { ZipArchive } from '../../src/vba/zip';
import { dialogsSeen, fields, hostInstalled, hostIsFree, psFile, q, scratchCopy, watchLines } from './officeHarness';

const TEMPLATES = fileURLToPath(new URL('../../assets/templates/', import.meta.url));

/** Opens a workbook in an Excel of the check's own, alerts on, runs its probe and saves it. */
async function openInExcel(file: string, label: string): Promise<{ seen: Record<string, string>; dialogs: string[]; err: string[] }> {
    const log = path.join(path.dirname(file), 'dialogs.log');
    const result = await psFile([
        '$ErrorActionPreference = "Continue"',
        '$xl = [Activator]::CreateInstance([type]::GetTypeFromProgID("Excel.Application"))',
        ...watchLines('$xl.Hwnd', log),
        '$xl.DisplayAlerts = $true',
        '$wb = $null',
        `try { $wb = $xl.Workbooks.Open(${q(file)}); "opened=True" } catch { "opened=False"; "error=" + $_ }`,
        'if ($wb) {',
        '  "components=" + (@($wb.VBProject.VBComponents | ForEach-Object { $_.Name + "/" + $_.Type }) -join ",")',
        '  "codeNames=" + (@($wb.Sheets | ForEach-Object { $_.Name + ":" + $_.CodeName }) -join ",")',
        '  try { "probe=" + $xl.Run("\'" + $wb.Name + "\'!XlideProbe.Report") } catch { "probe=failed: " + $_ }',
        '  try { $wb.Save(); "saved=True" } catch { "saved=False" }',
        '  $wb.Close($false)',
        '}',
        '$xl.DisplayAlerts = $false',
        '$wb = $null',
        '$xl.Quit()',
        '$xl = $null',
        '[GC]::Collect()',
        '[GC]::WaitForPendingFinalizers()',
        'Start-Sleep -Milliseconds 800',
    ].join('\r\n'), `excel-${label}.ps1`);
    return { seen: fields(result.out), dialogs: dialogsSeen(log), err: result.err };
}

/** Adds the project and a probe that reaches its document modules by code name. */
function addProjectWithProbe(file: string, template: string, probe: string): string[] {
    const added = addVbaProject(file, path.join(TEMPLATES, template));
    writeModule(file, 'XlideProbe', `Public Function Report() As String\r\n    Report = ${probe}\r\nEnd Function\r\n`, 'standard');
    return added.modules;
}

describe.runIf(await hostInstalled('excel'))('Add VBA Project, opened in Excel', () => {
    it('anchor: a workbook Excel would have to repair does not open, and raises no dialog to miss', async () => {
        // The check below reads "opened" as "Excel had nothing to repair".
        // That holds only if a workbook that needs repair fails to open under
        // automation, which this one - a cell pointing at a shared string that
        // does not exist - does.
        const file = scratchCopy('NoVbaFixture.xlsm', 'add-vba-anchor', 'Damaged.xlsm');
        const zip = ZipArchive.read(fs.readFileSync(file));
        const sheet = zip.read('xl/worksheets/sheet1.xml').toString('utf8');
        expect(sheet).toContain('<c r="A1" t="s"><v>0</v></c>');
        zip.write('xl/worksheets/sheet1.xml', Buffer.from(sheet.replace('<c r="A1" t="s"><v>0</v></c>', '<c r="A1" t="s"><v>999</v></c>'), 'utf8'));
        fs.writeFileSync(file, zip.toBytes());

        const { seen } = await openInExcel(file, 'anchor');
        expect(seen.opened).toBe('False');
    });

    it.each([
        ['NoVbaSheetsFixture.xlsm', 'blank.xlsm', 'ThisWorkbook:100,Chart1:100,Sheet2:100,Sheet3:100,Sheet4:100,XlideProbe:1', 'Trend:Chart1,Report:Sheet2,Data:Sheet3,Sheet4:Sheet4', 'ThisWorkbook|Trend|Report|Data|Sheet4|report'],
        ['NoVbaSheetsFixture.xlsb', 'blank.xlsb', 'ThisWorkbook:100,Chart1:100,Sheet2:100,Sheet3:100,Sheet4:100,XlideProbe:1', 'Trend:Chart1,Report:Sheet2,Data:Sheet3,Sheet4:Sheet4', 'ThisWorkbook|Trend|Report|Data|Sheet4|report'],
        ['NoVbaOtherSheetsFixture.xlsm', 'blank.xlsm', 'ThisWorkbook:100,Sheet1:100,XlideProbe:1', 'Sheet1:Sheet1,Dialog1:,Macro1:', 'ThisWorkbook|Sheet1|Sheet1|'],
    ])('%s opens with nothing to repair, runs code that names its sheets by code name, and saves', async (fixture, template, components, codeNames, probe) => {
        const file = scratchCopy(fixture, `add-vba-${fixture.replace(/\W/g, '-')}`);
        const reach = fixture.startsWith('NoVbaSheets')
            ? 'ThisWorkbook.CodeName & "|" & Chart1.Name & "|" & Sheet2.Name & "|" & Sheet3.Name & "|" & Sheet4.Name & "|" & Sheet2.Range("A1").Value'
            : 'ThisWorkbook.CodeName & "|" & Sheet1.Name & "|" & Sheet1.CodeName & "|" & Sheet1.Range("A1").Value';
        addProjectWithProbe(file, template, reach);

        const { seen, dialogs, err } = await openInExcel(file, fixture.replace(/\W/g, '-'));

        expect(seen.opened, `${seen.error ?? ''} ${err.join(' ')}`).toBe('True');
        expect(dialogs).toEqual([]);
        expect(seen.components.replace(/\//g, ':')).toBe(components);
        expect(seen.codeNames).toBe(codeNames);
        expect(seen.probe).toBe(probe);
        expect(seen.saved).toBe('True');
        // Saved by Excel, with its own compiled code, and still XLIDE's to read.
        expect(hasVbaProject(file)).toBe(true);
        expect(listModules(file).map((module) => module.name)).toContain('XlideProbe');
    });
});

describe.runIf(await hostIsFree('word'))('Add VBA Project, opened in Word', () => {
    it('opens with nothing to repair, runs code in the new project, and saves', async () => {
        const file = scratchCopy('NoVbaFixture.docm', 'add-vba-word');
        expect(addProjectWithProbe(file, 'blank.docm', 'ThisDocument.Name')).toEqual(['ThisDocument']);
        const log = path.join(path.dirname(file), 'dialogs.log');
        const result = await psFile([
            '$ErrorActionPreference = "Continue"',
            '$wd = [Activator]::CreateInstance([type]::GetTypeFromProgID("Word.Application"))',
            '$wd.Visible = $true',
            '$blank = $wd.Documents.Add()',
            ...watchLines('$wd.ActiveWindow.Hwnd', log),
            '$blank.Close(0)',
            // wdAlertsAll = -1, from the Word model.
            '$wd.DisplayAlerts = -1',
            '$doc = $null',
            `try { $doc = $wd.Documents.Open(${q(file)}); "opened=True" } catch { "opened=False"; "error=" + $_ }`,
            'if ($doc) {',
            '  "components=" + (@($doc.VBProject.VBComponents | ForEach-Object { $_.Name }) -join ",")',
            '  try { "probe=" + $wd.Run("XlideProbe.Report") } catch { "probe=failed: " + $_ }',
            '  try { $doc.Save(); "saved=True" } catch { "saved=False" }',
            '  $doc.Close(0)',
            '}',
            '$wd.DisplayAlerts = 0',
            '$doc = $null',
            '$blank = $null',
            '$wd.Quit(0)',
            '$wd = $null',
            '[GC]::Collect()',
            '[GC]::WaitForPendingFinalizers()',
        ].join('\r\n'), 'word-add-vba.ps1');
        const seen = fields(result.out);
        expect(seen.opened, seen.error).toBe('True');
        expect(dialogsSeen(log)).toEqual([]);
        expect(seen.components).toBe('ThisDocument,XlideProbe');
        expect(seen.probe).toBe('NoVbaFixture.docm');
        expect(seen.saved).toBe('True');
        expect(listModules(file).map((module) => module.name)).toEqual(['ThisDocument', 'XlideProbe']);
    });
});

describe.runIf(await hostIsFree('powerpoint'))('Add VBA Project, opened in PowerPoint', () => {
    it('opens an empty project with nothing to repair, runs code added to it, and saves', async () => {
        const file = scratchCopy('NoVbaFixture.pptm', 'add-vba-powerpoint');
        expect(addProjectWithProbe(file, 'blank.pptm', 'CStr(Presentations(1).Slides.Count)')).toEqual([]);
        const log = path.join(path.dirname(file), 'dialogs.log');
        const result = await psFile([
            '$ErrorActionPreference = "Continue"',
            '$pp = [Activator]::CreateInstance([type]::GetTypeFromProgID("PowerPoint.Application"))',
            ...watchLines('$pp.HWND', log),
            // ppAlertsAll = 2, from the PowerPoint model.
            '$pp.DisplayAlerts = 2',
            '$pres = $null',
            // Presentations.Open(FileName, ReadOnly:=msoFalse, Untitled:=msoFalse, WithWindow:=msoTrue)
            `try { $pres = $pp.Presentations.Open(${q(file)}, 0, 0, -1); "opened=True" } catch { "opened=False"; "error=" + $_ }`,
            'if ($pres) {',
            '  "components=" + (@($pres.VBProject.VBComponents | ForEach-Object { $_.Name }) -join ",")',
            '  try { "probe=" + $pp.GetType().InvokeMember("Run", [Reflection.BindingFlags]::InvokeMethod, $null, $pp, @($pres.Name + "!XlideProbe.Report")) } catch { "probe=failed: " + $_ }',
            '  try { $pres.Save(); "saved=True" } catch { "saved=False" }',
            '  $pres.Close()',
            '}',
            '$pres = $null',
            '$pp.Quit()',
            '$pp = $null',
            '[GC]::Collect()',
            '[GC]::WaitForPendingFinalizers()',
        ].join('\r\n'), 'powerpoint-add-vba.ps1');
        const seen = fields(result.out);
        expect(seen.opened, seen.error).toBe('True');
        expect(dialogsSeen(log)).toEqual([]);
        expect(seen.components).toBe('XlideProbe');
        expect(seen.probe).toBe('2');
        expect(seen.saved).toBe('True');
        expect(listModules(file).map((module) => module.name)).toEqual(['XlideProbe']);
    });
});
