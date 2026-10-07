import { describe, expect, it } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { getAccessObjectModel } from '../../src/analyzer/host/accessObjectModel';
import { analyzeProjectModule } from './helpers';

const wrap = (body: string, extra = '') => `Option Explicit\nSub Main()\n${body}\nEnd Sub\n${extra}`;
const errors = (source: string, options: Parameters<typeof analyzeModule>[1] = {}) => analyzeModule(source, options).filter(d => d.severity === 'error');

describe('shared runtime resources and host events', () => {
	it('forgets file handles across a function in an assignment', () => {
		const source = wrap('Dim ok As Boolean\nOpen "xlide-audit.tmp" For Output As #1\nClose #1\nok = Reopen()\nPrint #1, "x"', 'Function Reopen() As Boolean\nOpen "xlide-audit.tmp" For Output As #1\nReopen = True\nEnd Function');
		expect(errors(source)).toEqual([]);
	});

	it('forgets file handles across cross-module functions', () => {
		const source = wrap('Dim ok As Boolean\nOpen "xlide-audit.tmp" For Output As #1\nClose #1\nok = Reopen()\nPrint #1, "x"');
		const helper = 'Public Function Reopen() As Boolean\nOpen "xlide-audit.tmp" For Output As #1\nReopen = True\nEnd Function';
		expect(analyzeProjectModule(source, [{ moduleName: 'Helpers', source: helper }], 'Caller').filter(d => d.severity === 'error')).toEqual([]);
	});

	it('does not infer a new file mode after a handled failed Open', () => {
		expect(errors(wrap('Open "xlide-audit.tmp" For Output As #1\nOn Error Resume Next\nOpen "xlide-audit.tmp" For Input As #1\nOn Error GoTo 0\nPrint #1, "x"'))).toEqual([]);
	});

	it('does not infer absence of valid file handles from the project index', () => {
		expect(errors(wrap('Print #1, "x"'), { projectOpenedFileNumbers: { any: false, numbers: new Set() } })).toEqual([]);
		expect(errors(wrap('Print #0, "x"')).map(d => d.code)).toEqual(['file-number-zero']);
	});

	it('evaluates a DAO field RHS before assuming Edit was never called', () => {
		const source = wrap('Dim rs As DAO.Recordset\nSet rs = CurrentDb.OpenRecordset("T1")\nrs!Nm = BeginEdit(rs)', 'Function BeginEdit(rs As DAO.Recordset) As String\nrs.Edit\nBeginEdit = "x"\nEnd Function');
		expect(errors(source, { hostModel: getAccessObjectModel() })).toEqual([]);
	});

	it('does not infer DAO state for fields exposed to helpers', () => {
		const source = `Option Explicit\nDim rs As DAO.Recordset\nSub Main()\nSet rs = CurrentDb.OpenRecordset("T1")\nBeginEdit\nrs!Nm = "x"\nEnd Sub\nSub BeginEdit()\nrs.Edit\nEnd Sub`;
		expect(errors(source, { hostModel: getAccessObjectModel() })).toEqual([]);
	});

	it('does not mistake a possible runtime QueryDef name for absent SQL', () => {
		for (const body of ['CurrentDb.Execute "DELET FROM T1"', 'Dim rs As DAO.Recordset\nSet rs = CurrentDb.OpenRecordset("SELEC * FROM T1")']) {
			expect(errors(wrap(body), { hostModel: getAccessObjectModel() })).toEqual([]);
		}
	});

	it('does not assume a newly added sheet is empty before events finish', () => {
		const source = wrap('Dim ws As Worksheet\nSet ws = Worksheets.Add\nDebug.Print ws.Range("A1").Find("x").Address', 'Sub Workbook_NewSheet(ByVal Sh As Object)\nSh.Range("A1").Value = "x"\nEnd Sub');
		expect(errors(source)).toEqual([]);
	});

	it('does not infer workbook identity from successive active-workbook Adds', () => {
		expect(errors(wrap('Dim a As Worksheet, b As Worksheet\nSet a = Worksheets.Add\nSet b = Worksheets.Add\na.Name = "SameName"\nb.Name = "SameName"'))).toEqual([]);
	});

	it('forgets prior sheet names when another Add can trigger an event', () => {
		expect(errors(wrap('Dim a As Worksheet, b As Worksheet\nSet a = ThisWorkbook.Worksheets.Add\na.Name = "SameName"\nSet b = ThisWorkbook.Worksheets.Add\nb.Name = "SameName"'))).toEqual([]);
	});

	it('does not infer clipboard absence from CutCopyMode', () => {
		expect(errors(wrap('Application.CutCopyMode = False\nRange("B1").PasteSpecial'))).toEqual([]);
	});
});
