import { describe, it, expect } from 'vitest';
import {
	ProjectIndex,
	resolveMemberCompletions,
	resolveHostConstant,
	resolveHostGlobal,
	resolveHostAlias,
	resolveHostMemberSignature,
	resolveMemberReturnType,
	getHostMembers,
	type HostObjectModel,
} from '../src/analyzer';

/** Offset just after the dot in the first occurrence of `marker` in `src`. */
function dotOffset(src: string, marker: string): number {
	const idx = src.indexOf(marker);
	if (idx < 0) {
		throw new Error(`marker not found: ${marker}`);
	}
	return idx + marker.length;
}

function names(src: string, marker: string, ctx = {}): string[] {
	return resolveMemberCompletions(src, dotOffset(src, marker), ctx).map((m) => m.name);
}

describe('host model resolution', () => {
	it('resolves host globals to qualified types', () => {
		expect(resolveHostGlobal('ThisWorkbook')).toBe('Excel.Workbook');
		expect(resolveHostGlobal('thisworkbook')).toBe('Excel.Workbook');
		expect(resolveHostGlobal('Application')).toBe('Excel.Application');
		expect(resolveHostGlobal('ActiveSheet')).toBe('Excel.Worksheet');
		expect(resolveHostGlobal('ActiveWorkbook')).toBe('Excel.Workbook');
		expect(resolveHostGlobal('NotAGlobal')).toBeUndefined();
	});

	it('resolves generated host enum constants case-insensitively', () => {
		expect(resolveHostConstant('xlUp')?.type).toBe('XlDirection');
		expect(resolveHostConstant('XLUP')?.name).toBe('xlUp');
		expect(resolveHostConstant('notAConstant')).toBeUndefined();
	});

	it('resolves As-type aliases case-insensitively', () => {
		expect(resolveHostAlias('Worksheet')).toBe('Excel.Worksheet');
		expect(resolveHostAlias('range')).toBe('Excel.Range');
		expect(resolveHostAlias('Excel.Workbook')).toBe('Excel.Workbook');
		expect(resolveHostAlias('MyClass')).toBeUndefined();
	});

	it('resolves chainable member return types', () => {
		expect(resolveMemberReturnType('Excel.Workbook', 'Worksheets')).toBe('Excel.Worksheets');
		expect(resolveMemberReturnType('Excel.Worksheet', 'Range')).toBe('Excel.Range');
		expect(resolveMemberReturnType('Excel.Range', 'Offset')).toBe('Excel.Range');
		expect(resolveMemberReturnType('Excel.Range', 'Parent')).toBe('Excel.Worksheet');
		expect(resolveMemberReturnType('Excel.Workbook', 'Application')).toBe('Excel.Application');
		expect(resolveMemberReturnType('Excel.Application', 'Worksheets')).toBe('Excel.Worksheets');
		expect(resolveMemberReturnType('Excel.Workbooks', 'Item')).toBe('Excel.Workbook');
		expect(resolveMemberReturnType('Excel.Worksheets', 'Item')).toBe('Excel.Worksheet');
	});

	it('resolves chaining into the broadened object model', () => {
		// Range formatting objects.
		expect(resolveMemberReturnType('Excel.Range', 'Font')).toBe('Excel.Font');
		expect(resolveMemberReturnType('Excel.Range', 'Interior')).toBe('Excel.Interior');
		expect(resolveMemberReturnType('Excel.Range', 'Borders')).toBe('Excel.Borders');
		expect(resolveMemberReturnType('Excel.Borders', 'Item')).toBe('Excel.Border');
		// Tables / list objects.
		expect(resolveMemberReturnType('Excel.Worksheet', 'ListObjects')).toBe('Excel.ListObjects');
		expect(resolveMemberReturnType('Excel.ListObjects', 'Add')).toBe('Excel.ListObject');
		expect(resolveMemberReturnType('Excel.ListObject', 'Range')).toBe('Excel.Range');
		expect(resolveMemberReturnType('Excel.ListObject', 'ListColumns')).toBe('Excel.ListColumns');
		// Windows / names / charts.
		expect(resolveMemberReturnType('Excel.Application', 'ActiveWindow')).toBe('Excel.Window');
		expect(resolveMemberReturnType('Excel.Workbook', 'Names')).toBe('Excel.Names');
		expect(resolveMemberReturnType('Excel.Names', 'Add')).toBe('Excel.Name');
		expect(resolveMemberReturnType('Excel.Worksheet', 'Shapes')).toBe('Excel.Shapes');
		// WorksheetFunction is reachable from Application.
		expect(resolveMemberReturnType('Excel.Application', 'WorksheetFunction')).toBe(
			'Excel.WorksheetFunction',
		);
	});

	it('exposes the broadened types as host aliases', () => {
		expect(resolveHostAlias('Font')).toBe('Excel.Font');
		expect(resolveHostAlias('listobject')).toBe('Excel.ListObject');
		expect(resolveHostAlias('Window')).toBe('Excel.Window');
		expect(resolveHostAlias('WorksheetFunction')).toBe('Excel.WorksheetFunction');
	});
});

describe('member completion - collections', () => {
	it('offers Workbooks collection members after the Workbooks global', () => {
		const src = 'Sub Test()\n    Workbooks.\nEnd Sub\n';
		const got = names(src, 'Workbooks.');
		expect(got).toContain('Add');
		expect(got).toContain('Open');
		expect(got).toContain('Item');
		expect(got).toContain('Count');
		expect(got).toContain('Close');
	});

	it('resolves the Workbooks global case-insensitively', () => {
		const src = 'Sub Test()\n    workbooks.\nEnd Sub\n';
		const got = names(src, 'workbooks.');
		expect(got).toContain('Add');
		expect(got).toContain('Count');
	});

	it('offers Worksheets collection members after the Worksheets global', () => {
		const src = 'Sub Test()\n    Worksheets.\nEnd Sub\n';
		const got = names(src, 'Worksheets.');
		expect(got).toContain('Add');
		expect(got).toContain('Item');
		expect(got).toContain('Count');
	});

	it('chains a Workbook collection through to its element type', () => {
		// Workbooks.Item returns a Workbook -> Workbook members follow.
		const src = 'Sub Test()\n    Workbooks.Item(1).\nEnd Sub\n';
		const got = names(src, 'Workbooks.Item(1).');
		expect(got).toContain('Worksheets');
		expect(got).toContain('Save');
	});

	it('chains ThisWorkbook.Worksheets to the Worksheets collection', () => {
		const src = 'Sub Test()\n    ThisWorkbook.Worksheets.\nEnd Sub\n';
		const got = names(src, 'ThisWorkbook.Worksheets.');
		expect(got).toContain('Add');
		expect(got).toContain('Count');
	});

	it('keeps Sheets.Item ambiguous for single-return queries', () => {
		expect(resolveMemberReturnType('Excel.Sheets', 'Item')).toBeUndefined();
	});

	it('excludes host events from object member surfaces', () => {
		const workbook = getHostMembers('Excel.Workbook').map((member) => member.name);
		const worksheet = getHostMembers('Excel.Worksheet').map((member) => member.name);
		const application = getHostMembers('Excel.Application').map((member) => member.name);
		expect(workbook).not.toContain('AfterSave');
		expect(workbook).not.toContain('Open');
		expect(worksheet).not.toContain('Change');
		expect(application).not.toContain('SheetCalculate');
		expect(resolveHostMemberSignature('Excel.Workbook', 'AfterSave')).toBeUndefined();
		expect(resolveMemberReturnType('Excel.Workbook', 'AfterSave')).toBeUndefined();

		const model: HostObjectModel = {
			source: 'test fixture',
			aliases: {},
			globals: { Thing: 'Test.Thing' },
			types: {
				'Test.Thing': {
					displayName: 'Thing',
					members: [{ name: 'Changed', kind: 'event', signature: 'Changed()' }],
				},
			},
			memberSignatures: { 'Test.Thing': { changed: 'Changed()' } },
		};
		expect(getHostMembers('Test.Thing', model)).toEqual([]);
		expect(resolveHostMemberSignature('Test.Thing', 'Changed', model)).toBeUndefined();
	});
});

describe('member completion - host globals', () => {
	it('offers verified Workbook members after ThisWorkbook.', () => {
		const src = 'Sub Test()\n    ThisWorkbook.\nEnd Sub\n';
		const got = names(src, 'ThisWorkbook.');
		expect(got).toContain('Worksheets');
		expect(got).toContain('Save');
		expect(got).toContain('SaveAs');
		expect(got).toContain('Close');
		expect(got).toContain('Name');
		expect(got).toContain('FullName');
		expect(got).toContain('VBProject');
	});

	it('offers Application members after Application.', () => {
		const src = 'Sub Test()\n    Application.\nEnd Sub\n';
		const got = names(src, 'Application.');
		expect(got).toContain('ScreenUpdating');
		expect(got).toContain('Workbooks');
		expect(got).toContain('ThisWorkbook');
		expect(got).toContain('Calculate');
	});

	it('offers Worksheet members after ActiveSheet.', () => {
		const src = 'Sub Test()\n    ActiveSheet.\nEnd Sub\n';
		const got = names(src, 'ActiveSheet.');
		expect(got).toContain('Range');
		expect(got).toContain('Cells');
		expect(got).toContain('Activate');
	});

	it('filters by the partial member prefix already typed', () => {
		const src = 'Sub Test()\n    ThisWorkbook.Sav\nEnd Sub\n';
		const got = names(src, 'ThisWorkbook.Sav');
		expect(got).toContain('Save');
		expect(got).toContain('SaveAs');
		expect(got).toContain('SaveCopyAs');
		expect(got).not.toContain('Close');
	});

	it('does not invent members', () => {
		const src = 'Sub Test()\n    ThisWorkbook.\nEnd Sub\n';
		const got = names(src, 'ThisWorkbook.');
		expect(got).not.toContain('FooBar');
		expect(got).not.toContain('DoesNotExist');
	});

	it('includes dump-backed Workbook members after ThisWorkbook.', () => {
		const src = 'Sub Test()\n    ThisWorkbook.Accept\nEnd Sub\n';
		const got = resolveMemberCompletions(src, dotOffset(src, 'ThisWorkbook.Accept'));
		const accept = got.find((member) => member.name === 'AcceptAllChanges');
		expect(accept?.owner).toBe('Excel.Workbook');
		expect(accept?.surfaceExhaustive).toBe(true);
	});

	it('includes generated members on promoted non-exhaustive host surfaces', () => {
		const appSrc = 'Sub Test()\n    Application.Centi\nEnd Sub\n';
		const app = resolveMemberCompletions(appSrc, dotOffset(appSrc, 'Application.Centi'));
		const centimetersToPoints = app.find((member) => member.name === 'CentimetersToPoints');
		expect(centimetersToPoints?.owner).toBe('Excel.Application');
		expect(centimetersToPoints?.surfaceExhaustive).toBe(false);

		const rangeSrc = 'Sub Test(rng As Range)\n    rng.Spilling\nEnd Sub\n';
		const range = resolveMemberCompletions(rangeSrc, dotOffset(rangeSrc, 'rng.Spilling'));
		const spillingToRange = range.find((member) => member.name === 'SpillingToRange');
		expect(spillingToRange?.owner).toBe('Excel.Range');
		expect(spillingToRange?.surfaceExhaustive).toBe(false);

		const sheetSrc = 'Sub Test(ws As Worksheet)\n    ws.Named\nEnd Sub\n';
		const sheet = resolveMemberCompletions(sheetSrc, dotOffset(sheetSrc, 'ws.Named'));
		const namedSheetViews = sheet.find((member) => member.name === 'NamedSheetViews');
		expect(namedSheetViews?.owner).toBe('Excel.Worksheet');
		expect(namedSheetViews?.surfaceExhaustive).toBe(true);
	});

	it('does not offer Excel events as object member completions', () => {
		expect(names('Sub Test()\n    ThisWorkbook.After\nEnd Sub\n', 'ThisWorkbook.After')).not.toContain(
			'AfterSave',
		);
		expect(names('Sub Test()\n    Application.SheetCalc\nEnd Sub\n', 'Application.SheetCalc')).not.toContain(
			'SheetCalculate',
		);
		expect(names('Sub Test()\n    ActiveSheet.Change\nEnd Sub\n', 'ActiveSheet.Change')).not.toContain(
			'Change',
		);
	});

	it('uses the dump-backed Workbook surface for ActiveWorkbook and Workbook variables', () => {
		const activeSrc = 'Sub Test()\n    ActiveWorkbook.Accept\nEnd Sub\n';
		const active = resolveMemberCompletions(
			activeSrc,
			dotOffset(activeSrc, 'ActiveWorkbook.Accept'),
		);
		expect(
			active.find((member) => member.name === 'AcceptAllChanges')?.surfaceExhaustive,
		).toBe(true);

		const variableSrc =
			'Sub Test()\n' +
			'    Dim wb As Workbook\n' +
			'    wb.Accept\n' +
			'End Sub\n';
		const variable = resolveMemberCompletions(
			variableSrc,
			dotOffset(variableSrc, 'wb.Accept'),
		);
		expect(
			variable.find((member) => member.name === 'AcceptAllChanges')?.surfaceExhaustive,
		).toBe(true);
	});
});

describe('member completion - code names and Me', () => {
	it('resolves a worksheet code name from project context', () => {
		const src = 'Sub Test()\n    Sheet1.\nEnd Sub\n';
		const ctx = { codeNames: { sheet1: 'Excel.Worksheet' } };
		const got = names(src, 'Sheet1.', ctx);
		expect(got).toContain('Range');
		expect(got).toContain('Cells');
	});

	it('resolves a chart sheet code name from project context', () => {
		const src = 'Sub Test()\n    Chart1.\nEnd Sub\n';
		const ctx = { codeNames: { chart1: 'Excel.Chart' } };
		const got = names(src, 'Chart1.', ctx);
		expect(got).toContain('ChartType');
		expect(got).toContain('SeriesCollection');
		expect(got).not.toContain('Range');
	});

	it('does not merge a dangling dot on a previous line into the chain', () => {
		// `wb.` on its own line must not combine with `Sheet3.` below it.
		const src =
			'Sub Test()\n' +
			'    Dim wb As Workbook\n' +
			'    wb.\n' +
			'\n' +
			'    Sheet3.\n' +
			'End Sub\n';
		const ctx = { codeNames: { sheet3: 'Excel.Worksheet' } };
		const got = names(src, 'Sheet3.', ctx);
		expect(got).toContain('Range');
		expect(got).toContain('Cells');
	});

	it('terminates the receiver chain at a colon separator', () => {
		const src = 'Sub Test()\n    Dim wb As Workbook\n    wb.Save : Sheet3.\nEnd Sub\n';
		const ctx = { codeNames: { sheet3: 'Excel.Worksheet' } };
		const got = names(src, 'Sheet3.', ctx);
		expect(got).toContain('Range');
	});

	it('resolves Me to the module host type (Worksheet)', () => {
		const src = 'Sub Test()\n    Me.\nEnd Sub\n';
		const got = names(src, 'Me.', { meType: 'Excel.Worksheet' });
		expect(got).toContain('Range');
		expect(got).toContain('Name');
	});

	it('resolves Me to the module host type (Workbook)', () => {
		const src = 'Sub Test()\n    Me.\nEnd Sub\n';
		const got = names(src, 'Me.', { meType: 'Excel.Workbook' });
		expect(got).toContain('Worksheets');
		expect(got).toContain('Save');
	});

	it('merges Me document source members with its host surface', () => {
		const src = 'Sub Test()\n    Me.\nEnd Sub\n';
		const got = resolveMemberCompletions(src, dotOffset(src, 'Me.'), {
			meType: 'Excel.Workbook',
			meProjectType: 'ThisWorkbook',
			projectClassMembers: [
				{
					name: 'ThisWorkbook',
					kind: 'document',
					moduleName: 'ThisWorkbook',
					exhaustive: false,
					members: [
						{ name: 'Hello', kind: 'method', moduleName: 'ThisWorkbook' },
					],
				},
			],
		});
		expect(got.map((member) => member.name)).toContain('Hello');
		expect(got.map((member) => member.name)).toContain('AcceptAllChanges');
		expect(got.find((member) => member.name === 'Hello')?.surfaceExhaustive).toBe(true);
		expect(got.find((member) => member.name === 'AcceptAllChanges')?.surfaceExhaustive).toBe(true);
	});

	it('returns nothing for Me when module type is unknown', () => {
		const src = 'Sub Test()\n    Me.\nEnd Sub\n';
		expect(names(src, 'Me.', {})).toEqual([]);
	});
});

describe('member completion - declared variables', () => {
	it('resolves a local Dim As Worksheet variable', () => {
		const src = 'Sub Test()\n    Dim ws As Worksheet\n    ws.\nEnd Sub\n';
		const got = names(src, 'ws.');
		expect(got).toContain('Range');
		expect(got).toContain('UsedRange');
	});

	it('uses the declared Worksheet type after assignment from Sheets(index)', () => {
		const src =
			'Sub Test()\n' +
			'    Dim ws As Worksheet\n' +
			'    Set ws = Workbooks(1).Sheets(1)\n' +
			'    ws.\n' +
			'End Sub\n';
		const got = names(src, 'ws.');
		expect(got).toContain('Range');
		expect(got).toContain('Cells');
		expect(got).not.toContain('ChartType');
	});

	it('uses the declared Chart type after assignment from Sheets(index)', () => {
		const src =
			'Sub Test()\n' +
			'    Dim ch As Chart\n' +
			'    Set ch = Workbooks(1).Sheets(1)\n' +
			'    ch.\n' +
			'End Sub\n';
		const got = names(src, 'ch.');
		expect(got).toContain('ChartType');
		expect(got).toContain('SeriesCollection');
		expect(got).not.toContain('Range');
	});

	it('refines Object variables from simple Set assignments to known object expressions', () => {
		const src =
			'Sub Test()\n' +
			'    Dim obj As Object\n' +
			'    Set obj = Workbooks(1).Worksheets(1)\n' +
			'    obj.\n' +
			'End Sub\n';
		const got = names(src, 'obj.');
		expect(got).toContain('Range');
		expect(got).toContain('Cells');
		expect(got).not.toContain('ChartType');
	});

	it('keeps Object variables assigned from Sheets(index) on the merged item surface', () => {
		const src =
			'Sub Test()\n' +
			'    Dim obj As Object\n' +
			'    Set obj = Workbooks(1).Sheets(1)\n' +
			'    obj.\n' +
			'End Sub\n';
		const got = names(src, 'obj.');
		expect(got).toContain('Range');
		expect(got).toContain('Cells');
		expect(got).toContain('ChartType');
	});

	it('refines Variant variables from Set assignments and supports downstream chains', () => {
		const src =
			'Sub Test()\n' +
			'    Dim obj As Variant\n' +
			'    Set obj = Workbooks(1).Sheets(1)\n' +
			'    obj.Range("A1").\n' +
			'End Sub\n';
		const got = names(src, 'obj.Range("A1").');
		expect(got).toContain('Value');
		expect(got).toContain('Offset');
	});

	it('uses the latest preceding Set assignment for generic object variables', () => {
		const src =
			'Sub Test()\n' +
			'    Dim obj As Object\n' +
			'    Set obj = Workbooks(1).Worksheets(1)\n' +
			'    Set obj = ActiveWorkbook\n' +
			'    obj.\n' +
			'End Sub\n';
		const got = names(src, 'obj.');
		expect(got).toContain('Save');
		expect(got).toContain('Worksheets');
		expect(got).not.toContain('Range');
	});

	it('resolves a parameter typed As Range', () => {
		const src = 'Sub Test(rng As Range)\n    rng.\nEnd Sub\n';
		const got = names(src, 'rng.');
		expect(got).toContain('Value');
		expect(got).toContain('Offset');
	});

	it('resolves a module-level variable', () => {
		const src = 'Private wb As Workbook\n\nSub Test()\n    wb.\nEnd Sub\n';
		const got = names(src, 'wb.');
		expect(got).toContain('Worksheets');
		expect(got).toContain('Save');
	});

	it('resolves a variable declared inside a block', () => {
		const src =
			'Sub Test()\n    If True Then\n        Dim ws As Worksheet\n    End If\n    ws.\nEnd Sub\n';
		const got = names(src, 'ws.');
		expect(got).toContain('Range');
	});
});

describe('member completion - workbook classes', () => {
	const projectClassMembers = [
		{
			name: 'Person',
			kind: 'class' as const,
			moduleName: 'Person',
			members: [
				{ name: 'Name', kind: 'property' as const, returns: 'String', moduleName: 'Person' },
				{
					name: 'Age',
					kind: 'property' as const,
					returns: 'Integer',
					writable: true,
					writeType: 'Integer',
					moduleName: 'Person',
					doc: {
						summary: 'Age in whole years.',
						params: [],
						source: 'inline' as const,
					},
				},
				{ name: 'Save', kind: 'method' as const, moduleName: 'Person' },
				{ name: 'Manager', kind: 'method' as const, returns: 'Person', moduleName: 'Person' },
			],
		},
	];

	it('offers members for a variable declared as a project class', () => {
		const src = 'Sub Test()\n    Dim p As Person\n    p.\nEnd Sub\n';
		const got = names(src, 'p.', { projectClassMembers });
		expect(got).toContain('Name');
		expect(got).toContain('Save');
	});

	it('offers source-backed current class members through Me', () => {
		const src = 'Sub Test()\n    Me.\nEnd Sub\n';
		const got = names(src, 'Me.', { meProjectType: 'Person', projectClassMembers });
		expect(got).toContain('Name');
		expect(got).toContain('Save');
	});

	it('includes inline documentation for source-backed project class members', () => {
		const src = 'Sub Test()\n    Dim p As Person\n    p.Ag\nEnd Sub\n';
		const got = resolveMemberCompletions(src, dotOffset(src, 'p.Ag'), {
			projectClassMembers,
		});
		const age = got.find((m) => m.name === 'Age');
		expect(age?.documentation).toContain('Age in whole years.');
		expect(age?.writable).toBe(true);
		expect(age?.writeType).toBe('Integer');
		expect(age?.surfaceExhaustive).toBe(true);
	});

	it('offers public class fields and excludes invalid public constants', () => {
		const person =
			'Public Age As Integer\n' +
			'Public Const Species As String = "Human"\n';
		const index = new ProjectIndex();
		index.setModule({
			moduleName: 'Person',
			moduleKind: 'class',
			source: person,
		});
		const src = 'Sub Test()\n    Dim p As Person\n    p.\nEnd Sub\n';
		const got = resolveMemberCompletions(src, dotOffset(src, 'p.'), {
			projectClassMembers: index.projectClassMembers(),
		});
		const age = got.find((member) => member.name === 'Age');
		const species = got.find((member) => member.name === 'Species');
		expect(age?.writable).toBe(true);
		expect(age?.writeType).toBe('Integer');
		expect(species).toBeUndefined();
	});

	it('carries default-member attributes for source-backed project class members', () => {
		const person = [
			'Public Property Get Value() As String',
			'End Property',
			'Attribute Value.VB_UserMemId = 0',
		].join('\n');
		const index = new ProjectIndex();
		index.setModule({
			moduleName: 'Person',
			moduleKind: 'class',
			source: person,
		});
		const src = 'Sub Test()\n    Dim p As Person\n    p.Val\nEnd Sub\n';
		const got = resolveMemberCompletions(src, dotOffset(src, 'p.Val'), {
			projectClassMembers: index.projectClassMembers(),
		});
		const value = got.find((member) => member.name === 'Value');
		expect(value?.defaultMember).toBe(true);
		expect(value?.attributes?.[0]?.name).toBe('VB_UserMemId');
	});

	it('carries source definition locations for project class members', () => {
		const person =
			'Public Sub Save()\n' +
			'End Sub\n';
		const index = new ProjectIndex();
		index.setModule({
			moduleName: 'Person',
			moduleKind: 'class',
			source: person,
		});
		const src = 'Sub Test()\n    Dim p As Person\n    p.Sav\nEnd Sub\n';
		const got = resolveMemberCompletions(src, dotOffset(src, 'p.Sav'), {
			projectClassMembers: index.projectClassMembers(),
		});
		const save = got.find((member) => member.name === 'Save');
		expect(save?.definitions).toHaveLength(1);
		const def = save?.definitions?.[0];
		expect(def?.moduleName).toBe('Person');
		expect(def ? person.slice(def.nameSpan.start, def.nameSpan.end) : '').toBe('Save');
	});

	it('keeps same-named project class member definitions tied to the receiver type', () => {
		const index = new ProjectIndex();
		index.setModule({
			moduleName: 'Person',
			moduleKind: 'class',
			source: 'Public Property Get FirstName() As String\nEnd Property\n',
		});
		index.setModule({
			moduleName: 'Class1',
			moduleKind: 'class',
			source: 'Public Property Get FirstName() As String\nEnd Property\n',
		});
		const src = [
			'Sub Test()',
			'    Dim p As Person',
			'    Dim c As Class1',
			'    p.First',
			'    c.First',
			'End Sub',
		].join('\n');
		const projectClassMembers = index.projectClassMembers();

		const personFirst = resolveMemberCompletions(src, dotOffset(src, 'p.First'), {
			projectClassMembers,
		}).find((member) => member.name === 'FirstName');
		const classFirst = resolveMemberCompletions(src, dotOffset(src, 'c.First'), {
			projectClassMembers,
		}).find((member) => member.name === 'FirstName');

		expect(personFirst?.definitions?.map((definition) => definition.moduleName)).toEqual(['Person']);
		expect(classFirst?.definitions?.map((definition) => definition.moduleName)).toEqual(['Class1']);
	});

	it('chains through project class members that return a project class', () => {
		const src = 'Sub Test()\n    Dim p As Person\n    p.Manager.\nEnd Sub\n';
		const got = names(src, 'p.Manager.', { projectClassMembers });
		expect(got).toContain('Name');
		expect(got).toContain('Save');
	});

	it('does not resolve ambiguous project class member surfaces', () => {
		const src = 'Sub Test()\n    Dim p As Person\n    p.\nEnd Sub\n';
		expect(names(src, 'p.', {
			projectClassMembers: [
				...projectClassMembers,
				{ ...projectClassMembers[0], moduleName: 'OtherPerson' },
			],
		})).toEqual([]);
	});

	it('marks source-only document module member surfaces as non-exhaustive', () => {
		const src = 'Sub Test()\n    Dim wb As ThisWorkbook\n    wb.H\nEnd Sub\n';
		const got = resolveMemberCompletions(src, dotOffset(src, 'wb.H'), {
			projectClassMembers: [
				{
					name: 'ThisWorkbook',
					kind: 'document',
					moduleName: 'ThisWorkbook',
					exhaustive: false,
					members: [
						{ name: 'Hello', kind: 'method', moduleName: 'ThisWorkbook' },
					],
				},
			],
		});
		expect(got.map((member) => member.name)).toContain('Hello');
		expect(got[0]?.surfaceExhaustive).toBe(false);
	});

	it('merges ThisWorkbook source members with the exhaustive Workbook host surface', () => {
		const src = 'Sub Test()\n    ThisWorkbook.\nEnd Sub\n';
		const got = resolveMemberCompletions(src, dotOffset(src, 'ThisWorkbook.'), {
			projectClassMembers: [
				{
					name: 'ThisWorkbook',
					kind: 'document',
					moduleName: 'ThisWorkbook',
					exhaustive: false,
					members: [
						{ name: 'Hello', kind: 'method', moduleName: 'ThisWorkbook' },
					],
				},
			],
		});
		expect(got.map((member) => member.name)).toContain('Hello');
		expect(got.map((member) => member.name)).toContain('AcceptAllChanges');
		expect(got.find((member) => member.name === 'Hello')?.surfaceExhaustive).toBe(true);
		expect(got.find((member) => member.name === 'AcceptAllChanges')?.surfaceExhaustive).toBe(true);
	});
});

describe('member completion - chaining', () => {
	it('walks a member chain through return types', () => {
		const src = 'Sub Test()\n    ThisWorkbook.ActiveSheet.\nEnd Sub\n';
		const got = names(src, 'ThisWorkbook.ActiveSheet.');
		expect(got).toContain('Range');
		expect(got).toContain('Cells');
	});

	it('walks chains across call parentheses', () => {
		const src = 'Sub Test(ws As Worksheet)\n    ws.Range("A1").Offset(1, 0).\nEnd Sub\n';
		const got = names(src, 'ws.Range("A1").Offset(1, 0).');
		expect(got).toContain('Value');
		expect(got).toContain('Resize');
	});

	it('walks through collection default Item for Worksheets(index).Range', () => {
		const src = 'Sub Test()\n    ThisWorkbook.Worksheets(1).Range("A1").\nEnd Sub\n';
		const got = names(src, 'ThisWorkbook.Worksheets(1).Range("A1").');
		expect(got).toContain('Value');
		expect(got).toContain('Offset');
	});

	it('walks through global Workbooks(index).Worksheets(index).Range', () => {
		const src = 'Sub Test()\n    Workbooks(1).Worksheets(1).Range("A1").\nEnd Sub\n';
		const got = names(src, 'Workbooks(1).Worksheets(1).Range("A1").');
		expect(got).toContain('Value');
		expect(got).toContain('Resize');
	});

	it('carries verified call signatures for callable member completions', () => {
		const src = 'Sub Test()\n    Workbooks(1).Worksheets(1).Ran\nEnd Sub\n';
		const got = resolveMemberCompletions(
			src,
			dotOffset(src, 'Workbooks(1).Worksheets(1).Ran'),
		);
		const range = got.find((m) => m.name === 'Range');
		expect(range?.signature).toBe('Range(Cell1, [Cell2]) As Range');
	});

	it('carries generated reference signatures and docs for promoted host members', () => {
		const src = 'Sub Test()\n    Application.Calc\nEnd Sub\n';
		const got = resolveMemberCompletions(src, dotOffset(src, 'Application.Calc'));
		const calculate = got.find((m) => m.name === 'Calculate');
		expect(calculate?.signature).toBe('Calculate()');
		expect(calculate?.documentation).toContain('Calculates all open workbooks');
		expect(calculate?.doc?.source).toBe('external');
	});

	it('uses generated reference metadata ahead of fallback signatures', () => {
		const src = 'Sub Test()\n    Workbooks.Op\nEnd Sub\n';
		const got = resolveMemberCompletions(src, dotOffset(src, 'Workbooks.Op'));
		const open = got.find((m) => m.name === 'Open');
		expect(open?.signature).toContain('[ReadOnly As Variant]');
		expect(open?.doc?.params.find((p) => p.name === 'ReadOnly')?.text).toContain(
			'read-only mode',
		);
	});

	it('walks through indexed Sheets into merged sheet object members', () => {
		const sheetSrc = 'Sub Test()\n    Workbooks(1).Sheets(1).\nEnd Sub\n';
		const sheetMembers = names(sheetSrc, 'Workbooks(1).Sheets(1).');
		expect(sheetMembers).toContain('Range');
		expect(sheetMembers).toContain('Cells');
		expect(sheetMembers).toContain('ChartType');

		const src = 'Sub Test()\n    Workbooks(1).Sheets(1).Range("A1").\nEnd Sub\n';
		const got = names(src, 'Workbooks(1).Sheets(1).Range("A1").');
		expect(got).toContain('Value');
		expect(got).toContain('Offset');
	});

	it('walks through explicit Sheets.Item into worksheet members', () => {
		const src = 'Sub Test()\n    Workbooks(1).Sheets.Item(1).Range("A1").\nEnd Sub\n';
		const got = names(src, 'Workbooks(1).Sheets.Item(1).Range("A1").');
		expect(got).toContain('Value');
		expect(got).toContain('Resize');
	});

	it('resolves Range.Worksheet back to a worksheet', () => {
		const src = 'Sub Test(rng As Range)\n    rng.Worksheet.\nEnd Sub\n';
		const got = names(src, 'rng.Worksheet.');
		expect(got).toContain('Cells');
		expect(got).toContain('Range');
	});
});

describe('member completion - user-defined types', () => {
	function userTypeIndex(): ProjectIndex {
		const index = new ProjectIndex();
		index.setModule({
			moduleName: 'Types',
			moduleKind: 'standard',
			source: [
				'Public Type TPoint',
				'    X As Long',
				'    Y As Long',
				'End Type',
				'Public Type TBox',
				'    Corner As TPoint',
				'End Type',
				'Private Type THidden',
				'    Secret As String',
				'End Type',
			].join('\n'),
		});
		index.setModule({ moduleName: 'Caller', moduleKind: 'standard', source: '' });
		return index;
	}

	it('offers fields for a variable declared as a visible UDT', () => {
		const index = userTypeIndex();
		const src = 'Sub Test()\n    Dim p As TPoint\n    p.\nEnd Sub\n';
		const got = resolveMemberCompletions(src, dotOffset(src, 'p.'), {
			projectClassMembers: index.projectMemberSurfaces('Caller'),
		});
		expect(got.map((member) => member.name)).toEqual(['X', 'Y']);
		expect(got.find((member) => member.name === 'X')?.writeType).toBe('Long');
		expect(got.find((member) => member.name === 'X')?.surfaceExhaustive).toBe(true);
	});

	it('chains through nested UDT fields', () => {
		const index = userTypeIndex();
		const src = 'Sub Test()\n    Dim box As TBox\n    box.Corner.\nEnd Sub\n';
		const got = names(src, 'box.Corner.', {
			projectClassMembers: index.projectMemberSurfaces('Caller'),
		});
		expect(got).toContain('X');
		expect(got).toContain('Y');
	});

	it('resolves leading-dot fields inside With blocks', () => {
		const index = userTypeIndex();
		const src = 'Sub Test()\n    Dim p As TPoint\n    With p\n        .\n    End With\nEnd Sub\n';
		const got = names(src, '        .', {
			projectClassMembers: index.projectMemberSurfaces('Caller'),
		});
		expect(got).toContain('X');
		expect(got).toContain('Y');
	});

	it('does not expose private UDT fields outside their module', () => {
		const index = userTypeIndex();
		const src = 'Sub Test()\n    Dim hidden As THidden\n    hidden.\nEnd Sub\n';
		const got = names(src, 'hidden.', {
			projectClassMembers: index.projectMemberSurfaces('Caller'),
		});
		expect(got).toEqual([]);
	});
});

describe('member completion - With blocks', () => {
	it('resolves a leading-dot member against the active With receiver', () => {
		const src =
			'Sub Test(rng As Range)\n' +
			'    With rng\n' +
			'        .Va\n' +
			'    End With\n' +
			'End Sub\n';
		const got = names(src, '.Va');
		expect(got).toContain('Value');
		expect(got).toContain('Value2');
	});

	it('walks leading-dot chains from the active With receiver', () => {
		const src =
			'Sub Test(ws As Worksheet)\n' +
			'    With ws\n' +
			'        .Range("A1").\n' +
			'    End With\n' +
			'End Sub\n';
		const got = names(src, '.Range("A1").');
		expect(got).toContain('Value');
		expect(got).toContain('Offset');
	});

	it('keeps resolving inside an unfinished With block while editing', () => {
		const src =
			'Sub Test(rng As Range)\n' +
			'    With rng\n' +
			'        .Off\n' +
			'End Sub\n';
		const got = names(src, '.Off');
		expect(got).toContain('Offset');
	});
});

describe('member completion - negative cases', () => {
	it('returns nothing when not in a member-access position', () => {
		const src = 'Sub Test()\n    ThisWorkbook\nEnd Sub\n';
		const offset = src.indexOf('ThisWorkbook') + 'ThisWorkbook'.length;
		expect(resolveMemberCompletions(src, offset)).toEqual([]);
	});

	it('returns nothing for an unknown receiver', () => {
		const src = 'Sub Test()\n    Foo.\nEnd Sub\n';
		expect(names(src, 'Foo.')).toEqual([]);
	});

	it('returns nothing when chaining through a non-chainable member', () => {
		const src = 'Sub Test()\n    ThisWorkbook.Name.\nEnd Sub\n';
		expect(names(src, 'ThisWorkbook.Name.')).toEqual([]);
	});

	it('host member lists are non-empty and well-formed', () => {
		for (const type of [
			'Excel.Application',
			'Excel.Workbook',
			'Excel.Worksheet',
			'Excel.Range',
			'Excel.Workbooks',
			'Excel.Worksheets',
			'Excel.Sheets',
		]) {
			const members = getHostMembers(type);
			expect(members.length).toBeGreaterThan(0);
			for (const mem of members) {
				expect(mem.name).toMatch(/^[A-Za-z_][A-Za-z0-9_]*$/);
				expect(['property', 'method']).toContain(mem.kind);
			}
		}
	});

	it('marks promoted non-Workbook host member surfaces as non-exhaustive', () => {
		const src = 'Sub Test()\n    Application.Work\nEnd Sub\n';
		const got = resolveMemberCompletions(src, dotOffset(src, 'Application.Work'));
		const workbooks = got.find((member) => member.name === 'Workbooks');
		expect(workbooks?.owner).toBe('Excel.Application');
		expect(workbooks?.surfaceExhaustive).toBe(false);
	});

	it('marks generated Worksheet host member surfaces as exhaustive', () => {
		const src = 'Sub Test()\n    ActiveSheet.Named\nEnd Sub\n';
		const got = resolveMemberCompletions(src, dotOffset(src, 'ActiveSheet.Named'));
		const namedSheetViews = got.find((member) => member.name === 'NamedSheetViews');
		expect(namedSheetViews?.owner).toBe('Excel.Worksheet');
		expect(namedSheetViews?.surfaceExhaustive).toBe(true);
	});

	it('marks dump-backed Workbook host member surfaces as exhaustive', () => {
		const src = 'Sub Test()\n    ThisWorkbook.Sav\nEnd Sub\n';
		const got = resolveMemberCompletions(src, dotOffset(src, 'ThisWorkbook.Sav'));
		const save = got.find((member) => member.name === 'Save');
		expect(save?.owner).toBe('Excel.Workbook');
		expect(save?.surfaceExhaustive).toBe(true);
	});

	it('can mark a verified exhaustive host member surface', () => {
		const model: HostObjectModel = {
			source: 'test fixture',
			aliases: {},
			globals: { Thing: 'Test.Thing' },
			types: {
				'Test.Thing': {
					displayName: 'Thing',
					exhaustive: true,
					members: [{ name: 'Known', kind: 'method' }],
				},
			},
		};
		const src = 'Sub Test()\n    Thing.K\nEnd Sub\n';
		const got = resolveMemberCompletions(src, dotOffset(src, 'Thing.K'), { model });
		expect(got[0]?.name).toBe('Known');
		expect(got[0]?.surfaceExhaustive).toBe(true);
	});
});
