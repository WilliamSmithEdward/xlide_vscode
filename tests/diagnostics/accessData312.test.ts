// Diagnostics tests: Access SQL literals and DAO recordsets whose failure
// the code shows (issue #312). Each sample was run through pyVBAharness on
// 2026-10-02 in Access 16.0, creating its own table first.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { getAccessObjectModel } from '../../src/analyzer/host/accessObjectModel';
import { getExcelObjectModel } from '../../src/analyzer/host/excelObjectModel';

const RS = 'Dim rs As DAO.Recordset\n    ';

function errors(body: string, host: ReturnType<typeof getExcelObjectModel> = getAccessObjectModel()): string[] {
	const source = `Option Explicit\nFunction Main() As Variant\n    ${body}\n    If IsEmpty(Main) Then Main = 1\nEnd Function\n`;
	return analyzeModule(source, { hostModel: host })
		.filter((diag) => diag.code === 'runtime-argument-value' || diag.code === 'host-argument-out-of-range')
		.map((diag) => `${diag.code} ${/Run-time error '(-?\d+)'/.exec(diag.message)?.[1] ?? '?'}`);
}

describe('Access SQL a literal shows will fail (issue #312)', () => {
	it('is reported with the error Access raises', () => {
		const cases: Array<[string, string]> = [
			['CurrentDb.Execute "SELECT * FROM T1"', '3065'],
			['CurrentDb.Execute "select ID from T1", dbFailOnError', '3065'],
			['DoCmd.RunSQL "SELECT * FROM T1"', '2342'],
			['CurrentDb.Execute ""', '3078'],
			['CurrentDb.Execute "INSERT INTO T1 (Nm) VALUES (\'a)"', '3075'],
			['DoCmd.RunSQL "INSERT INTO T1 (Nm) VALUES (\'a)"', '2342'],
			['CurrentDb.Execute "INSERT INTO T1 (Nm VALUES (\'a\')"', '3134'],
			[`${RS}Set rs = CurrentDb.OpenRecordset("SELECT * FROM T1 WHERE Nm = 'a")`, '3075'],
			['Dim q As DAO.QueryDef\n    Set q = CurrentDb.CreateQueryDef("", "SELEC 1")', '3129'],
			['Main = DLookup("Nm", "T1", "ID = ")', '2342'],
		];
		for (const [body, error] of cases) {
			expect(errors(body), body).toEqual([`runtime-argument-value ${error}`]);
		}
	});

	it('stays quiet on SQL that runs, and on a saved query named in the text', () => {
		for (const body of [
			'CurrentDb.Execute "DELETE FROM T1"',
			`${RS}Set rs = CurrentDb.OpenRecordset("SELECT * FROM T1")`,
			'Main = DLookup("Nm", "T1", "ID = 1")',
			'CurrentDb.Execute "INSERT INTO T1 (Nm) VALUES (\'it\'\'s\')"',
			'CurrentDb.Execute "UPDATE T1 SET Nm = \'(\' WHERE ID = 1"',
			'CurrentDb.Execute "qryArchive"',
			'Dim db As DAO.Database\n    Set db = CurrentDb\n    db.Execute "DELETE FROM T1"',
		]) {
			expect(errors(body), body).toEqual([]);
		}
	});
});

describe('a DAO recordset in a state that refuses the call (issue #312)', () => {
	it('is reported', () => {
		const cases: Array<[string, string]> = [
			[`${RS}Set rs = CurrentDb.OpenRecordset("T1", dbOpenDynaset)\n    rs!Nm = "x"`, '3020'],
			[`${RS}Set rs = CurrentDb.OpenRecordset("T1", dbOpenDynaset)\n    rs.Fields("Nm").Value = "x"`, '3020'],
			[`${RS}Set rs = CurrentDb.OpenRecordset("T1", dbOpenDynaset)\n    rs.Update`, '3020'],
			[`${RS}Set rs = CurrentDb.OpenRecordset("T1", dbOpenDynaset)\n    rs.Edit\n    rs!Nm = "x"\n    rs.Update\n    rs!Nm = "y"`, '3020'],
			[`${RS}Set rs = CurrentDb.OpenRecordset("T1", dbOpenSnapshot)\n    rs.Edit`, '3251'],
			[`${RS}Set rs = CurrentDb.OpenRecordset("T1", dbOpenSnapshot)\n    rs.AddNew`, '3251'],
			[`${RS}Set rs = CurrentDb.OpenRecordset("T1", dbOpenDynaset)\n    rs.Close\n    Main = rs.EOF`, '3420'],
		];
		for (const [body, error] of cases) {
			expect(errors(body), body).toEqual([`host-argument-out-of-range ${error}`]);
		}
	});

	it('stays quiet when the state allows it, or is no longer known', () => {
		for (const body of [
			`${RS}Set rs = CurrentDb.OpenRecordset("T1", dbOpenDynaset)\n    rs.Edit\n    rs!Nm = "x"\n    rs.Update`,
			`${RS}Set rs = CurrentDb.OpenRecordset("T1", dbOpenDynaset)\n    rs.AddNew\n    rs!ID = 2\n    rs.Update`,
			`${RS}Set rs = CurrentDb.OpenRecordset("T1", dbOpenDynaset)\n    Main = rs!Nm\n    rs.Close`,
			`${RS}Set rs = CurrentDb.OpenRecordset("T1", dbOpenDynaset)\n    Prepare rs\n    rs!Nm = "x"`,
			`${RS}Set rs = CurrentDb.OpenRecordset("T1", dbOpenDynaset)\n    If Main Then rs.Edit\n    rs!Nm = "x"`,
		]) {
			expect(errors(body), body).toEqual([]);
		}
	});

	it('leaves the other hosts alone', () => {
		expect(errors('CurrentDb.Execute "SELECT * FROM T1"', getExcelObjectModel())).toEqual([]);
	});
});
