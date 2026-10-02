// Diagnostics tests: WithEvents of a class that raises no events, and an
// Optional default that names nothing (issue #445). Measured in Excel 16.0
// (build 20326, 2026-10-02) with Debug > Compile.

import { describe, it, expect } from 'vitest';
import { byCode, expectDiagnostic } from '../helpers/diagnostics';
import { analyzeProjectModule } from './helpers';

const PLAIN = 'Option Explicit\nPublic N As Long\n';
const EVENTFUL = 'Option Explicit\nPublic Event Done()\nPublic Sub Fire()\n    RaiseEvent Done\nEnd Sub\n';

function holder(decl: string): string[] {
	const src = `Option Explicit\n${decl}\n`;
	return analyzeProjectModule(src, [
		{ moduleName: 'Holder', source: src, type: 'class' },
		{ moduleName: 'Plain', source: PLAIN, type: 'class' },
		{ moduleName: 'Eventful', source: EVENTFUL, type: 'class' },
	], 'Holder', { moduleKind: 'class' }).filter((diag) => diag.code === 'withevents-declaration').map((diag) => diag.message);
}

describe('WithEvents of a project class', () => {
	it('is refused when the class raises no events', () => {
		expect(holder('Private WithEvents s As Plain')).toEqual([expect.stringContaining('does not source automation events')]);
	});

	it('compiles when the class has an Event, and without WithEvents', () => {
		expect(holder('Private WithEvents s As Eventful')).toEqual([]);
		expect(holder('Private s As Plain')).toEqual([]);
	});
});

describe('an Optional default that names nothing', () => {
	it('is Variable not defined', () => {
		const src = 'Option Explicit\nFunction Main() As Variant\n    Main = F()\nEnd Function\nPrivate Function F(Optional x As Long = y) As Long\n    F = x\nEnd Function\n';
		expectDiagnostic(src, byCode(analyzeProjectModule(src, [], 'Module1'), 'undeclared-variable'), 'undeclared-variable', { span: 'y', message: "Optional parameter's default" });
	});

	it('compiles when it names a Const', () => {
		const src = 'Option Explicit\nPrivate Const y = 3\nFunction Main() As Variant\n    Main = F()\nEnd Function\nPrivate Function F(Optional x As Long = y) As Long\n    F = x\nEnd Function\n';
		expect(byCode(analyzeProjectModule(src, [], 'Module1'), 'undeclared-variable')).toEqual([]);
	});
});
