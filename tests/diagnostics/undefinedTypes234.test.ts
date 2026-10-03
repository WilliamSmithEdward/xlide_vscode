// Diagnostics tests: an As clause naming a type no type of the project and
// none of its referenced libraries holds (issue #234). Measured on
// 2026-10-03 in Excel 16.0 (build 20430), a workbook referencing VBA, Excel,
// stdole and Office.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

const DEFAULT_REFERENCES = ['VBA', 'Excel', 'stdole', 'Office'];

function errors(type: string, referencedLibraries: readonly string[] | null = DEFAULT_REFERENCES, extra = ''): string[] {
	const src = `Option Explicit\n${extra}Private Function F() As ${type}\nEnd Function\n`;
	return analyzeModule(src, referencedLibraries === null ? {} : { referencedLibraries }).filter((diag) => diag.severity === 'error').map((diag) => diag.code);
}

describe('a type no referenced library holds (issue #234)', () => {
	it('reports a name none of them spells', () => {
		expect(errors('Qwerty')).toEqual(['invalid-as-type-name']);
		// MSForms is referenced only once a form is inserted.
		expect(errors('DataObject')).toEqual(['invalid-as-type-name']);
	});

	it('takes the names the libraries hold, hidden ones included', () => {
		for (const type of ['CommandBar', 'StdFont', 'XlDirection', 'ErrObject', 'Collection', 'IUnknown', 'Range', 'Long', 'Variant', 'Object', 'IFontDisp', 'VbMsgBoxResult', 'Workbook', 'FileDialog']) {
			expect(errors(type), type).toEqual([]);
		}
		expect(errors('DataObject', [...DEFAULT_REFERENCES, 'MSForms'])).toEqual([]);
	});

	it('stays quiet where the references are not known or not all read', () => {
		expect(errors('Qwerty', null)).toEqual([]);
		expect(errors('Qwerty', [...DEFAULT_REFERENCES, 'ADODB'])).toEqual([]);
		expect(errors('Qwerty', DEFAULT_REFERENCES, 'Private Type Qwerty\n    v As Long\nEnd Type\n')).toEqual([]);
	});

	it('leaves the Scripting Runtime to missing-library-reference', () => {
		expect(errors('Dictionary')).toEqual(['missing-library-reference']);
	});
});
