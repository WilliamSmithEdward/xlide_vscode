// Diagnostics tests: compile errors from declaration scope (issue #490).
// Each sample was compiled through pyVBAharness on 2026-10-02 in Excel 16.0
// (build 20430).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { analyzeProjectModule } from './helpers';

const errors = (diagnostics: ReturnType<typeof analyzeModule>): string[] =>
	diagnostics.filter((d) => d.severity === 'error').map((d) => d.code ?? '');

const MAIN = (body: string): string => `Option Explicit\nFunction Main() As Variant\n    ${body}\n    Main = 1\nEnd Function\n`;

function inModule1(body: string, module2: string): string[] {
	const source = MAIN(body);
	return errors(analyzeProjectModule(source, [
		{ moduleName: 'Module1', source },
		{ moduleName: 'Module2', source: module2 },
	], 'Module1'));
}

describe('declaration scope (issue #490)', () => {
	it('takes a Type or Declare with no keyword in a class as Public', () => {
		const type = 'Option Explicit\nType T\n    x As Long\nEnd Type\n';
		const declare = 'Option Explicit\nDeclare PtrSafe Function GT Lib "kernel32" Alias "GetTickCount" () As Long\n';
		expect(errors(analyzeModule(type, { moduleKind: 'class' }))).toEqual(['object-module-public-member']);
		expect(errors(analyzeModule(declare, { moduleKind: 'class' }))).toEqual(['object-module-public-member']);
		const quiet = 'Option Explicit\nEnum E\n    eA\nEnd Enum\nConst K As Long = 1\n';
		expect(errors(analyzeModule(quiet, { moduleKind: 'class' }))).toEqual([]);
	});

	it('refuses a Private Type or Enum of another module, bare or qualified', () => {
		const privateType = 'Option Explicit\nPrivate Type T\n    x As Long\nEnd Type\n';
		expect(inModule1('Dim t As T', privateType)).toEqual(['invalid-as-type-name']);
		expect(inModule1('Dim t As Module2.T', privateType)).toEqual(['invalid-as-type-name']);
		expect(inModule1('Dim e As E', 'Option Explicit\nPrivate Enum E\n    eA\nEnd Enum\n')).toEqual(['invalid-as-type-name']);
		expect(inModule1('Dim t As T', 'Option Explicit\nPublic Type T\n    x As Long\nEnd Type\n')).toEqual([]);
		expect(inModule1('Dim t As Module2.T', 'Option Explicit\nType T\n    x As Long\nEnd Type\n')).toEqual([]);
	});

	it('refuses a variable as a type qualifier', () => {
		expect(inModule1('Dim c As New Collection\n    Dim t As c.T', 'Option Explicit\nPublic Type T\n    x As Long\nEnd Type\n')).toEqual(['invalid-as-type-name']);
	});
});
