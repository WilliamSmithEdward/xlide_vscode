// Diagnostics tests: Implements member completeness (issue #125). Measured
// in Excel 16.0 (build 20326, 2026-09-25): a class that implements IFoo
// without IFoo_Name is refused, and so is IFoo_Name with another signature.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import type { VbaProjectClassMembers } from '../../src/analyzer/symbols/symbolModel';
import { byCode, expectDiagnostic } from '../helpers/diagnostics';

const IFOO: VbaProjectClassMembers = {
	name: 'IFoo',
	kind: 'class',
	moduleName: 'IFoo',
	exhaustive: true,
	members: [
		{ name: 'Name', kind: 'method', moduleName: 'IFoo', returns: 'String', signature: 'Name() As String' },
		{ name: 'Size', kind: 'method', moduleName: 'IFoo', signature: 'Size(ByVal n As Long)' },
	],
};

const analyze = (src: string) => analyzeModule(src, { moduleName: 'CFoo', moduleKind: 'class', projectClassMembers: [IFOO] });

describe('implements-member-missing and implements-member-signature (issue #125)', () => {
	it('flags every interface member the class leaves out', () => {
		const src = 'Option Explicit\nImplements IFoo\n';
		const hits = byCode(analyze(src), 'implements-member-missing');
		expect(hits.map((hit) => hit.message)).toEqual([
			expect.stringContaining("implement 'Name' for interface 'IFoo'"),
			expect.stringContaining("implement 'Size' for interface 'IFoo'"),
		]);
	});

	it('flags an implementation whose parameters or return differ', () => {
		const src = 'Option Explicit\nImplements IFoo\nPrivate Function IFoo_Name(ByVal extra As Long) As String\nEnd Function\nPrivate Sub IFoo_Size(ByVal n As Integer)\nEnd Sub\n';
		const hits = byCode(analyze(src), 'implements-member-signature');
		expect(hits.map((hit) => hit.message)).toEqual([
			expect.stringContaining('takes 0 parameters, this procedure 1'),
			expect.stringContaining('parameter 1 is Integer here and long on the interface'),
		]);
	});

	it('stays quiet when every member matches, and outside object modules', () => {
		const src = 'Option Explicit\nImplements IFoo\nPrivate Function IFoo_Name() As String\n    IFoo_Name = "x"\nEnd Function\nPrivate Sub IFoo_Size(ByVal n As Long)\nEnd Sub\n';
		expect(byCode(analyze(src), 'implements-member-missing')).toHaveLength(0);
		expect(byCode(analyze(src), 'implements-member-signature')).toHaveLength(0);
		expectDiagnostic('Option Explicit\nImplements IFoo\n', analyzeModule('Option Explicit\nImplements IFoo\n', { moduleKind: 'standard', projectClassMembers: [IFOO] }), 'implements-statement-placement');
	});
});
