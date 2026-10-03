// Diagnostics tests: #312's follow-ups (issue #611). Each sample was run
// through pyVBAharness on 2026-10-02 in Access 16.0, creating its own
// table first.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { getAccessObjectModel } from '../../src/analyzer/host/accessObjectModel';

const RS = 'Dim rs As DAO.Recordset\n    ';

function errors(body: string): string[] {
	const source = `Option Explicit\nFunction Main() As Variant\n    ${body}\n    If IsEmpty(Main) Then Main = 1\nEnd Function\n`;
	return analyzeModule(source, { hostModel: getAccessObjectModel() })
		.filter((diag) => diag.code === 'runtime-argument-value' || diag.code === 'host-argument-out-of-range')
		.map((diag) => /Run-time error '(-?\d+)'/.exec(diag.message)?.[1] ?? '?');
}

describe('Access SQL and DAO recordsets, #312\'s follow-ups (issue #611)', () => {
	it('leaves alone a make-table SELECT and a double-quoted string', () => {
		for (const body of [
			'CurrentDb.Execute "SELECT * INTO T2 FROM T1"',
			'CurrentDb.Execute "SELECT ID INTO T3 FROM T1 WHERE ID = 1"',
			'CurrentDb.Execute "INSERT INTO T1 (Nm) VALUES (""it\'s"")"',
			'CurrentDb.Execute "INSERT INTO T1 (Nm) VALUES (\'it\'\'s\')"',
			'CurrentDb.Execute "INSERT INTO T1 (Nm) VALUES (""a"")"',
			'CurrentDb.Execute "INSERT INTO T1 SELECT * FROM T1"',
			'CurrentDb.Execute "INSERT INTO T1 (ID, Nm) SELECT ID, Nm FROM T1"',
			'Main = DLookup("Nm", "T1", "Nm = ""a""")',
			'Main = DCount("*", "T1", "ID = 1")',
		]) {
			expect(errors(body), body).toEqual([]);
		}
	});

	it('reports the SQL Access refuses', () => {
		const cases: Array<[string, string]> = [
			[`${RS}Set rs = CurrentDb.OpenRecordset("SELECT * FROM T1 WHERE (ID = 1")`, '3075'],
			['CurrentDb.Execute "SELECT * FROM T1 WHERE (ID = 1"', '3075'],
			['CurrentDb.Execute "UPDATE T1 SET Nm = \'b\' WHERE (ID = 1"', '3075'],
			['CurrentDb.Execute "TRANSFORM Count(ID) SELECT Nm FROM T1 GROUP BY Nm PIVOT ID"', '3065'],
			['CurrentDb.Execute "INSERT INTO T1 (Nm) (\'a\')"', '3134'],
			['Main = DLookup("Nm", "T1", "ID = 1 AND")', '3075'],
			['Main = DLookup("Nm", "T1", "ID = 1 OR")', '3075'],
			['Main = DLookup("Nm", "T1", "Nm = \'a")', '3075'],
			['Main = DCount("*", "T1", "ID = ")', '3075'],
			['Main = DCount("*", "T1", "ID = 1 AND")', '3075'],
			['Main = DSum("ID", "T1", "ID = ")', '3075'],
			['Main = DMax("ID", "T1", "ID >")', '3075'],
		];
		for (const [body, error] of cases) {
			expect(errors(body), body).toEqual([error]);
		}
	});

	it('reports a recordset used against its state', () => {
		const cases: Array<[string, string]> = [
			[`${RS}Set rs = CurrentDb.OpenRecordset("T1", dbOpenSnapshot)\n    rs.Delete`, '3251'],
			[`${RS}Set rs = CurrentDb.OpenRecordset("T1", dbOpenForwardOnly)\n    rs.Edit`, '3251'],
			[`${RS}Set rs = CurrentDb.OpenRecordset("T1", dbOpenForwardOnly)\n    rs.AddNew`, '3251'],
			[`${RS}Set rs = CurrentDb.OpenRecordset("T1", dbOpenDynaset, dbReadOnly)\n    rs.Edit`, '3027'],
			[`${RS}Set rs = CurrentDb.OpenRecordset("T1", dbOpenDynaset, dbReadOnly)\n    rs.AddNew`, '3027'],
			[`${RS}Set rs = CurrentDb.OpenRecordset("T1", dbOpenDynaset)\n    rs.Edit\n    rs!Nm = "x"\n    rs.MoveFirst\n    rs!Nm = "y"`, '3020'],
			[`${RS}Set rs = CurrentDb.OpenRecordset("T1", dbOpenDynaset)\n    With rs\n        !Nm = "x"\n    End With`, '3020'],
			[`${RS}Dim r2 As DAO.Recordset\n    Set rs = CurrentDb.OpenRecordset("T1", dbOpenDynaset)\n    Set r2 = rs\n    r2.Close\n    Main = rs.EOF`, '3420'],
		];
		for (const [body, error] of cases) {
			expect(errors(body), body).toEqual([error]);
		}
		for (const body of [
			`${RS}Set rs = CurrentDb.OpenRecordset("T1", dbOpenDynaset)\n    rs.Delete`,
			`${RS}Set rs = CurrentDb.OpenRecordset("T1", dbOpenDynaset)\n    With rs\n        .Edit\n        !Nm = "x"\n        .Update\n    End With`,
			`${RS}Set rs = CurrentDb.OpenRecordset("T1", dbOpenDynaset)\n    rs.Edit\n    rs!Nm = "x"\n    rs.Update\n    rs.MoveFirst\n    Main = rs!Nm`,
		]) {
			expect(errors(body), body).toEqual([]);
		}
	});
});
