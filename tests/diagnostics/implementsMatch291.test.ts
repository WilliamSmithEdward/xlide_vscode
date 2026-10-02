// Diagnostics tests: what Implements matches beyond names and types (issue
// #291). Measured in Excel 16.0 64-bit (2026-10-02) through pyVBAharness,
// with an interface class IFoo and a class Impl that implements it.

import { describe, it, expect } from 'vitest';
import { analyzeVbaModuleSource } from '../../src/vbaModuleAnalysis';
import { buildVbaProjectIndex, projectAnalysisOptionsForModule, projectProcedureSignatures } from '../../src/vbaProjectAnalysis';

function found(iface: string, impl: string): string[] {
	const modules = [
		{ moduleName: 'IFoo', type: 'class', source: `Option Explicit\n${iface}\n` },
		{ moduleName: 'Impl', type: 'class', source: `Option Explicit\nImplements IFoo\n${impl}\n` },
	];
	const project = buildVbaProjectIndex(modules, undefined, { conditionalCompilation: { projectConstants: {} } });
	const procedures = projectProcedureSignatures(project);
	const { diagnostics } = analyzeVbaModuleSource({
		source: modules[1].source,
		moduleName: 'Impl',
		moduleKind: 'class',
		...projectAnalysisOptionsForModule(project, 'Impl', procedures),
	} as never);
	return diagnostics.filter((diag) => diag.severity === 'error').map((diag) => diag.code);
}

const prop = (kind: string, name: string, params: string, type = ''): string => `Private Property ${kind} IFoo_${name}(${params})${type}\nEnd Property`;

describe('an Implements class', () => {
	it('owes nothing for a Friend member', () => {
		expect(found('Public Sub Go()\nEnd Sub\nFriend Sub Hidden()\nEnd Sub', 'Private Sub IFoo_Go()\nEnd Sub')).toEqual([]);
	});

	it('passes, defaults and returns as the interface does', () => {
		expect(found('Public Sub S(ByVal n As Long)\nEnd Sub', 'Private Sub IFoo_S(ByRef n As Long)\nEnd Sub')).toEqual(['implements-member-signature']);
		expect(found('Public Sub S(Optional ByVal n As Long = 1)\nEnd Sub', 'Private Sub IFoo_S(Optional ByVal n As Long = 2)\nEnd Sub')).toEqual(['implements-member-signature']);
		expect(found('Public Sub S(Optional ByVal n As Long = 1)\nEnd Sub', 'Private Sub IFoo_S(ByVal n As Long)\nEnd Sub')).toEqual(['implements-member-signature']);
		expect(found('Public Function F() As Long\nEnd Function', 'Private Sub IFoo_F()\nEnd Sub')).toEqual(['implements-member-signature']);
		expect(found('Public Sub S(ByRef n As Long)\nEnd Sub', 'Private Sub IFoo_S(n As Long)\nEnd Sub')).toEqual([]);
		expect(found('Public Sub S(ByVal n As Long)\nEnd Sub', 'Private Sub IFoo_S(ByVal m As Long)\nEnd Sub')).toEqual([]);
	});

	it('implements a Public variable with the accessors and passing its type needs', () => {
		expect(found('Public N As Long', `${prop('Get', 'N', '', ' As Long')}\n${prop('Let', 'N', 'ByVal x As Long')}`)).toEqual([]);
		expect(found('Public N As Long', `${prop('Get', 'N', '', ' As Long')}\n${prop('Let', 'N', 'ByRef x As Long')}`)).toEqual(['implements-member-signature']);
		expect(found('Public C As Collection', `${prop('Get', 'C', '', ' As Collection')}\n${prop('Set', 'C', 'ByVal x As Collection')}`)).toEqual([]);
		expect(found('Public C As Collection', `${prop('Get', 'C', '', ' As Collection')}\n${prop('Let', 'C', 'ByVal x As Collection')}`)).toEqual(['implements-member-missing']);
		expect(found('Public V As Variant', `${prop('Get', 'V', '', ' As Variant')}\n${prop('Let', 'V', 'x As Variant')}\n${prop('Set', 'V', 'x As Variant')}`)).toEqual([]);
		expect(found('Public V As Variant', `${prop('Get', 'V', '', ' As Variant')}\n${prop('Let', 'V', 'x As Variant')}`)).toEqual(['implements-member-missing']);
		expect(found('Public Name As String', `${prop('Get', 'Name', '', ' As Long')}\n${prop('Let', 'Name', 'ByVal x As Long')}`)).toEqual(['implements-member-signature']);
	});
});
