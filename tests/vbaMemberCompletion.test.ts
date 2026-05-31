import { describe, it, expect } from 'vitest';
import {
	resolveMemberCompletions,
	resolveHostGlobal,
	resolveHostAlias,
	resolveMemberReturnType,
	getHostMembers,
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

	it('does not chain Sheets.Item because it is a mixed Object collection', () => {
		expect(resolveMemberReturnType('Excel.Sheets', 'Item')).toBeUndefined();
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
});

describe('member completion - code names and Me', () => {
	it('resolves a worksheet code name from project context', () => {
		const src = 'Sub Test()\n    Sheet1.\nEnd Sub\n';
		const ctx = { codeNames: { sheet1: 'Excel.Worksheet' } };
		const got = names(src, 'Sheet1.', ctx);
		expect(got).toContain('Range');
		expect(got).toContain('Cells');
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

	it('resolves Range.Worksheet back to a worksheet', () => {
		const src = 'Sub Test(rng As Range)\n    rng.Worksheet.\nEnd Sub\n';
		const got = names(src, 'rng.Worksheet.');
		expect(got).toContain('Cells');
		expect(got).toContain('Range');
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
		for (const type of ['Excel.Application', 'Excel.Workbook', 'Excel.Worksheet', 'Excel.Range']) {
			const members = getHostMembers(type);
			expect(members.length).toBeGreaterThan(0);
			for (const mem of members) {
				expect(mem.name).toMatch(/^[A-Za-z][A-Za-z0-9]*$/);
				expect(['property', 'method', 'event']).toContain(mem.kind);
			}
		}
	});
});
