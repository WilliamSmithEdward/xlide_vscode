// Diagnostics tests: a name that is also a standard module's name, used as a
// line label or before the module's own Function (issue #403). Every case was
// measured in Excel 16.0 (build 20326, 2026-10-01), compiled with the VBE's
// Debug > Compile.

import { describe, it, expect } from 'vitest';
import { byCode, expectDiagnostic } from '../helpers/diagnostics';
import { analyzeProjectModule, type ProjectTestModule } from './helpers';

function source(...lines: string[]): string {
	return `Option Explicit\nFunction Main() As Variant\n${lines.map((line) => `    ${line}`).join('\n')}\nEnd Function\n`;
}

describe('a line label named like a standard module', () => {
	const foo: ProjectTestModule[] = [{ moduleName: 'Foo', source: 'Option Explicit\nPublic Sub Bar()\nEnd Sub\n' }];
	it('compiles, as a label and as a GoTo, GoSub or Resume target', () => {
		for (const lines of [
			['GoTo Foo', 'Foo:', 'Main = 1'],
			['Main = 1', 'Foo:'],
			['Foo: Main = 1'],
			['On Error GoTo Foo', 'Main = 1', 'Exit Function', 'Foo:', 'Main = 2'],
			['On Error GoTo H', 'Main = 1', 'Foo:', 'Exit Function', 'H:', 'Resume Foo'],
			['GoSub Foo', 'Exit Function', 'Foo:', 'Main = 1', 'Return'],
		]) {
			const src = source(...lines);
			expect(byCode(analyzeProjectModule(src, foo, 'Module1'), 'malformed-statement'), lines.join(' / ')).toEqual([]);
		}
	});

	it('still reports the module name called bare', () => {
		const src = source('Foo', 'Main = 1');
		expectDiagnostic(src, byCode(analyzeProjectModule(src, foo, 'Module1'), 'malformed-statement'), 'malformed-statement', { span: 'Foo', message: 'not module' });
	});
});

describe('a module qualifier named like its own Function', () => {
	it('is the module, not the Function value', () => {
		const foo: ProjectTestModule[] = [{ moduleName: 'Foo', source: 'Option Explicit\nPublic Function Foo() As Long\n    Foo = 1\nEnd Function\n' }];
		const src = 'Option Explicit\nPublic Sub Z()\n    Dim n As Long\n    n = Foo.Foo()\nEnd Sub\n';
		expect(byCode(analyzeProjectModule(src, foo, 'Module2'), 'scalar-member-access')).toEqual([]);
	});

	it('still reports a Long local of that name', () => {
		const foo: ProjectTestModule[] = [{ moduleName: 'Foo', source: 'Option Explicit\nPublic Function Foo() As Long\n    Foo = 1\nEnd Function\n' }];
		const src = 'Option Explicit\nPublic Sub Z()\n    Dim Foo As Long, n As Long\n    n = Foo.Bar\nEnd Sub\n';
		expectDiagnostic(src, byCode(analyzeProjectModule(src, foo, 'Module2'), 'scalar-member-access'), 'scalar-member-access', { span: 'Foo.' });
	});
});
