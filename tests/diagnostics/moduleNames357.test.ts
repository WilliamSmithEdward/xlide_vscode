// Diagnostics tests: module names (issue #357). Measured in Excel 16.0
// (build 20326, 2026-10-01), each module added with VBComponents.Add and
// renamed, and called qualified from another module.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { byCode, expectDiagnostic } from '../helpers/diagnostics';
import { analyzeProjectModule } from './helpers';

const CALLER = (name: string): string => `Option Explicit\nFunction Main() As Variant\n    Main = ${name}.Hi()\nEnd Function\n`;
const CALLEE = 'Option Explicit\nPublic Function Hi() As Long\n    Hi = 7\nEnd Function\n';

function errors(src: string, name: string): string[] {
	return analyzeProjectModule(src, [{ moduleName: name, source: CALLEE }], 'Module1')
		.filter((diag) => diag.severity === 'error')
		.map((diag) => `${diag.code}: ${diag.message}`);
}

describe('a project module named like a library or a keyword', () => {
	it('is called through its name, ahead of a library the project does not reference', () => {
		for (const name of ['Word', 'Access', 'PowerPoint', 'Property']) {
			expect(errors(CALLER(name), name), name).toEqual([]);
		}
	});

	it('still reports the unreferenced library when no module bears its name', () => {
		const src = CALLER('Word');
		expectDiagnostic(src, byCode(analyzeModule(src), 'missing-library-reference'), 'missing-library-reference', { span: 'Word.Hi' });
	});
});

describe('a name the VBE refuses for a module', () => {
	it('reports the libraries every Excel project references, and six more keywords', () => {
		for (const [name, message] of [
			['Excel', 'object library'],
			['VBA', 'object library'],
			['Office', 'object library'],
			['stdole', 'object library'],
			['False', 'cannot name a module'],
			['WithEvents', 'cannot name a module'],
			['LongLong', 'cannot name a module'],
			['LongPtr', 'cannot name a module'],
			['Decimal', 'cannot name a module'],
			['CDate', 'cannot name a module'],
		] as const) {
			const src = `Attribute VB_Name = "${name}"\n${CALLEE}`;
			expectDiagnostic(src, byCode(analyzeModule(src, { moduleName: name }), 'invalid-declaration-name'), 'invalid-declaration-name', { span: `"${name}"`, message });
		}
	});

	it('accepts the names Excel lets a module take', () => {
		for (const name of ['Word', 'Access', 'PowerPoint', 'Outlook', 'MSForms', 'Line', 'Name', 'Property', 'Collection', 'Range', 'Application', 'Value']) {
			const src = `Attribute VB_Name = "${name}"\n${CALLEE}`;
			expect(byCode(analyzeModule(src, { moduleName: name }), 'invalid-declaration-name'), name).toEqual([]);
		}
	});
});
