import { describe, it, expect } from 'vitest';
import {
	resolveCanonicalCaseEdit,
	resolveCanonicalCaseEdits,
	type CanonicalCaseContext,
} from '../src/analyzer';

function editAtMarker(src: string, ctx: CanonicalCaseContext = {}) {
	const offset = src.indexOf('|');
	if (offset < 0) {
		throw new Error('Missing | marker');
	}
	return resolveCanonicalCaseEdit(src.replace('|', ''), offset, ctx);
}

function editsOnLine(src: string, lineNumber: number, ctx: CanonicalCaseContext = {}) {
	const lines = src.split('\n');
	const start = lines.slice(0, lineNumber).reduce((sum, line) => sum + line.length + 1, 0);
	const end = start + lines[lineNumber].replace(/\r$/, '').length;
	return resolveCanonicalCaseEdits(src, { start, end }, ctx);
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

	it('canonicalizes every safe token on a completed line', () => {
		const src =
			'Sub T()\n' +
			'    dim p as person: set p = new person\n' +
			'End Sub\n';
		const edits = editsOnLine(src, 1, {
			type: {
				projectTypes: [{ name: 'Person', kind: 'class' }],
			},
		});

		expect(edits.map((edit) => [src.slice(edit.start, edit.end), edit.text])).toEqual([
			['dim', 'Dim'],
			['as', 'As'],
			['person', 'Person'],
			['set', 'Set'],
			['new', 'New'],
			['person', 'Person'],
		]);
	});

	it('line canonicalization keeps strings and comments unchanged', () => {
		const src =
			'Sub T()\n' +
			'    msgbox "left" \' application.calculate\n' +
			'End Sub\n';
		const edits = editsOnLine(src, 1);

		expect(edits.map((edit) => [src.slice(edit.start, edit.end), edit.text])).toEqual([
			['msgbox', 'MsgBox'],
		]);
	});
});
