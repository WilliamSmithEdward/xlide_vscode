// A library reference XLIDE writes, checked by the applications themselves:
// each file gets the reference from XLIDE and a probe that names the
// library's constants early bound, then opens in its application, which
// lists the reference unbroken and runs the probe. The probe compiles only
// if the reference resolves, and starts no other application: the Outlook
// probes never start Outlook.
//
// Excel and Access run in instances the check starts itself. Word and
// PowerPoint only when neither is running, so no instance of theirs is
// ever touched.

import { describe, expect, it } from 'vitest';
import { addReference, removeReference, writeModule } from '../../src/vba/projectService';
import { fields, hostInstalled, hostIsFree, psFile, q, scratchCopy } from './officeHarness';

/** Excel's constants, early bound: `xlSheetVisible` is -1 and `XlSheetVisibility.xlSheetHidden` 0. */
const EXCEL_PROBE = [
	'Public Function Report() As String',
	'    Dim state As Excel.XlSheetVisibility',
	'    state = Excel.XlSheetVisibility.xlSheetHidden',
	'    Report = CStr(xlSheetVisible) & "/" & CStr(state)',
	'End Function',
	'',
].join('\r\n');

/** Outlook's constants, early bound: `olMailItem` is 0 and `OlItemType.olAppointmentItem` 1. */
const OUTLOOK_PROBE = [
	'Public Function Report() As String',
	'    Dim kind As Outlook.OlItemType',
	'    kind = Outlook.OlItemType.olAppointmentItem',
	'    Report = CStr(olMailItem) & "/" & CStr(kind)',
	'End Function',
	'',
].join('\r\n');

/** A copy of the fixture with XLIDE's reference to `library` and the probe module. */
function referencedCopy(fixture: string, label: string, library: string, name: string, probe: string): string {
	const file = scratchCopy(fixture, label);
	expect(addReference(file, library)).toEqual({ ok: true, added: true, name });
	writeModule(file, 'XlideProbe', probe, 'standard');
	return file;
}

/** The lines that list a project's references and run its probe, given the PowerShell names of each. */
const report = (project: string, run: string): string[] => [
	`"references=" + (@(${project}.References | ForEach-Object { $_.Name + ":" + $_.IsBroken }) -join ",")`,
	`try { "probe=" + (${run}) } catch { "probe=failed: " + $_ }`,
];

/** Opens the database in an Access of the check's own, lists its references and runs the probe. */
async function openInAccess(file: string, label: string): Promise<Record<string, string>> {
	const result = await psFile([
		'$ErrorActionPreference = "Continue"',
		'$app = [Activator]::CreateInstance([type]::GetTypeFromProgID("Access.Application"))',
		`$app.OpenCurrentDatabase(${q(file)})`,
		...report('$app', '$app.Run("Report")'),
		'$app.CloseCurrentDatabase()',
		'$app.Quit()',
		'$app = $null',
		'[GC]::Collect()',
		'[GC]::WaitForPendingFinalizers()',
	].join('\r\n'), `access-${label}.ps1`);
	return fields(result.out);
}

describe.runIf(await hostInstalled('access'))('a reference XLIDE writes, opened in Access', () => {
	it('resolves, and code naming the library runs (issue #1474)', async () => {
		const file = referencedCopy('AccessFixture.accdb', 'reference-access', 'excel', 'Excel', EXCEL_PROBE);
		const seen = await openInAccess(file, 'excel-reference');
		expect(seen.references.split(',')).toContain('Excel:False');
		expect(seen.probe).toBe('-1/0');
	});

	it('takes the reference away again, and Access no longer lists it', async () => {
		const file = scratchCopy('AccessFixture.accdb', 'reference-access-removed');
		addReference(file, 'excel');
		expect(removeReference(file, 'excel')).toMatchObject({ ok: true, removed: true });
		const seen = await openInAccess(file, 'excel-reference-removed');
		expect(seen.references).toBe('VBA:False,Access:False,stdole:False,DAO:False');
	});

	it('resolves the Outlook reference, and code naming Outlook runs', async () => {
		const file = referencedCopy('AccessFixture.accdb', 'reference-access-outlook', 'outlook', 'Outlook', OUTLOOK_PROBE);
		const seen = await openInAccess(file, 'outlook-reference');
		expect(seen.references.split(',')).toContain('Outlook:False');
		expect(seen.probe).toBe('0/1');
	});
});

describe.runIf(await hostInstalled('excel'))('the Outlook reference, opened in Excel', () => {
	it('resolves, and code naming Outlook runs', async () => {
		const file = referencedCopy('FormFixture.xlsm', 'reference-excel', 'outlook', 'Outlook', OUTLOOK_PROBE);
		const result = await psFile([
			'$ErrorActionPreference = "Continue"',
			'$xl = [Activator]::CreateInstance([type]::GetTypeFromProgID("Excel.Application"))',
			'$xl.DisplayAlerts = $false',
			`$wb = $xl.Workbooks.Open(${q(file)})`,
			...report('$wb.VBProject', '$xl.Run("\'" + $wb.Name + "\'!XlideProbe.Report")'),
			'$wb.Close($false)',
			'$wb = $null',
			'$xl.Quit()',
			'$xl = $null',
			'[GC]::Collect()',
			'[GC]::WaitForPendingFinalizers()',
		].join('\r\n'), 'excel-outlook-reference.ps1');
		const seen = fields(result.out);
		expect(seen.references.split(',')).toContain('Outlook:False');
		expect(seen.probe).toBe('0/1');
	});
});

describe.runIf(await hostIsFree('word'))('the Outlook reference, opened in Word', () => {
	it('resolves, and code naming Outlook runs', async () => {
		const file = referencedCopy('WordFixture.docm', 'reference-word', 'outlook', 'Outlook', OUTLOOK_PROBE);
		const result = await psFile([
			'$ErrorActionPreference = "Continue"',
			'$wd = [Activator]::CreateInstance([type]::GetTypeFromProgID("Word.Application"))',
			'$wd.DisplayAlerts = 0',
			`$doc = $wd.Documents.Open(${q(file)})`,
			...report('$doc.VBProject', '$wd.Run("XlideProbe.Report")'),
			'$doc.Close(0)',
			'$doc = $null',
			'$wd.Quit(0)',
			'$wd = $null',
			'[GC]::Collect()',
			'[GC]::WaitForPendingFinalizers()',
		].join('\r\n'), 'word-outlook-reference.ps1');
		const seen = fields(result.out);
		expect(seen.references.split(',')).toContain('Outlook:False');
		expect(seen.probe).toBe('0/1');
	});
});

describe.runIf(await hostIsFree('powerpoint'))('the Outlook reference, opened in PowerPoint', () => {
	it('resolves, and code naming Outlook runs', async () => {
		const file = referencedCopy('PowerPointFixture.pptm', 'reference-powerpoint', 'outlook', 'Outlook', OUTLOOK_PROBE);
		const result = await psFile([
			'$ErrorActionPreference = "Continue"',
			'$pp = [Activator]::CreateInstance([type]::GetTypeFromProgID("PowerPoint.Application"))',
			// Presentations.Open(FileName, ReadOnly:=msoFalse, Untitled:=msoFalse, WithWindow:=msoFalse)
			`$pres = $pp.Presentations.Open(${q(file)}, 0, 0, 0)`,
			...report('$pres.VBProject', '$pp.GetType().InvokeMember("Run", [Reflection.BindingFlags]::InvokeMethod, $null, $pp, @($pres.Name + "!XlideProbe.Report"))'),
			'$pres.Close()',
			'$pres = $null',
			'$pp.Quit()',
			'$pp = $null',
			'[GC]::Collect()',
			'[GC]::WaitForPendingFinalizers()',
		].join('\r\n'), 'powerpoint-outlook-reference.ps1');
		const seen = fields(result.out);
		expect(seen.references.split(',')).toContain('Outlook:False');
		expect(seen.probe).toBe('0/1');
	});
});
