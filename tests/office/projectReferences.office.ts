// A library reference XLIDE writes, checked by the applications themselves:
// each file gets the reference from XLIDE and a probe that names the
// library's constants early bound, then opens in its application, which
// lists the reference unbroken and runs the probe. The probe compiles only
// if the reference resolves, and starts no other application.
//
// Access runs in an instance the check starts itself.

import { describe, expect, it } from 'vitest';
import { addReference, removeReference, writeModule } from '../../src/vba/projectService';
import { fields, hostInstalled, psFile, q, scratchCopy } from './officeHarness';

/** Excel's constants, early bound: `xlSheetVisible` is -1 and `XlSheetVisibility.xlSheetHidden` 0. */
const EXCEL_PROBE = [
	'Public Function Report() As String',
	'    Dim state As Excel.XlSheetVisibility',
	'    state = Excel.XlSheetVisibility.xlSheetHidden',
	'    Report = CStr(xlSheetVisible) & "/" & CStr(state)',
	'End Function',
	'',
].join('\r\n');

/** Opens the database in an Access of the check's own, lists its references and runs the probe. */
async function openInAccess(file: string, label: string): Promise<Record<string, string>> {
	const result = await psFile([
		'$ErrorActionPreference = "Continue"',
		'$app = [Activator]::CreateInstance([type]::GetTypeFromProgID("Access.Application"))',
		`$app.OpenCurrentDatabase(${q(file)})`,
		'"references=" + (@($app.References | ForEach-Object { $_.Name + ":" + $_.IsBroken }) -join ",")',
		'try { "probe=" + $app.Run("Report") } catch { "probe=failed: " + $_ }',
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
		const file = scratchCopy('AccessFixture.accdb', 'reference-access');
		expect(addReference(file, 'excel')).toEqual({ ok: true, added: true, name: 'Excel' });
		writeModule(file, 'XlideProbe', EXCEL_PROBE, 'standard');
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
});
