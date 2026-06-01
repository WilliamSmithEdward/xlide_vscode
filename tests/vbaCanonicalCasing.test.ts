import { describe, it, expect } from 'vitest';
import { resolveCanonicalCaseEdit, type CanonicalCaseContext } from '../src/analyzer';

function editAtMarker(src: string, ctx: CanonicalCaseContext = {}) {
	const offset = src.indexOf('|');
	if (offset < 0) {
		throw new Error('Missing | marker');
	}
	return resolveCanonicalCaseEdit(src.replace('|', ''), offset, ctx);
}

describe('canonical casing edits', () => {
	it('canonicalizes host member names after a boundary', () => {
		const edit = editAtMarker(
			'Sub T()\n    Workbooks(1).Worksheets(1).Range("A1").value| = 1\nEnd Sub\n',
		);
		expect(edit?.text).toBe('Value');
	});

	it('canonicalizes leading-dot host member names inside With', () => {
		const edit = editAtMarker(
			'Sub T()\n    With Range("A1")\n        .value| = 1\n    End With\nEnd Sub\n',
		);
		expect(edit?.text).toBe('Value');
	});

	it('canonicalizes runtime functions before an argument list', () => {
		const edit = editAtMarker('Sub T()\n    x = left|("test", 3)\nEnd Sub\n');
		expect(edit?.text).toBe('Left');
	});

	it('canonicalizes exported project procedures', () => {
		const edit = editAtMarker('Sub T()\n    mysub|\nEnd Sub\n', {
			identifier: {
				projectProcedures: [
					{
						name: 'mySub',
						moduleName: 'Helpers',
						kind: 'sub',
						params: [],
					},
				],
			},
		});
		expect(edit?.text).toBe('mySub');
	});

	it('canonicalizes type names in declaration and New expression positions', () => {
		const ctx: CanonicalCaseContext = {
			type: {
				projectTypes: [{ name: 'Person', kind: 'class' }],
			},
		};
		expect(editAtMarker('Sub T()\n    Dim p As person|\nEnd Sub\n', ctx)?.text).toBe('Person');
		expect(editAtMarker('Sub T()\n    Set p = New person|\nEnd Sub\n', ctx)?.text).toBe('Person');
		expect(editAtMarker('Sub T()\n    Dim amount As currency|\nEnd Sub\n', ctx)?.text).toBe('Currency');
	});

	it('canonicalizes source-backed current class members through Me', () => {
		const edit = editAtMarker('Sub T()\n    Me.save|\nEnd Sub\n', {
			member: {
				meProjectType: 'Person',
				projectClassMembers: [
					{
						name: 'Person',
						kind: 'class',
						moduleName: 'Person',
						members: [{ name: 'Save', kind: 'method', moduleName: 'Person' }],
					},
				],
			},
		});
		expect(edit?.text).toBe('Save');
	});

	it('canonicalizes keywords without touching comments or strings', () => {
		expect(editAtMarker('sub| T()\nEnd Sub\n')?.text).toBe('Sub');
		expect(editAtMarker('Sub T()\n    MsgBox "left|("\nEnd Sub\n')).toBeUndefined();
		expect(editAtMarker("Sub T()\n    ' left|(\nEnd Sub\n")).toBeUndefined();
	});
});
