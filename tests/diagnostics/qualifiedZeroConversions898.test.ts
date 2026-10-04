import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { analyzeVbaModuleSource } from '../../src/vbaModuleAnalysis';
import { byCode, expectDiagnostic } from '../helpers/diagnostics';

const wrap = (expression: string): string =>
	`Option Explicit\nFunction Main() As Variant\n    Main = ${expression}\nEnd Function\n`;

// Issue #898: Excel 16.0 build 20430.20092 raises 11 for
// 10 / VBA.CDbl(0), while 10 / VBA.CDbl(0.4) returns 25.
describe.each([
	['raw analyzer', (source: string) => analyzeModule(source)],
	['module wrapper', (source: string) => analyzeVbaModuleSource({ source, moduleName: 'Module1', moduleType: 'standard' }).diagnostics],
] as const)('VBA-qualified zero conversions (%s, issue #898)', (_name, analyze) => {
	it.each([
		['10 / VBA.CDbl(0)', 'VBA.CDbl(0)', "'11'"],
		['10 / vBa.cDbL(0)', 'vBa.cDbL(0)', "'11'"],
		['10 / +VBA.CDbl(0)', '+VBA.CDbl(0)', "'11'"],
		['10 / -VBA.CDbl(0)', '-VBA.CDbl(0)', "'11'"],
		['10 / (VBA.CDbl(0))', 'VBA.CDbl(0)', "'11'"],
		['10 / VBA.CLng(0.4)', 'VBA.CLng(0.4)', "'11'"],
		['10 \\ VBA.CDbl(0)', 'VBA.CDbl(0)', "'11'"],
		['10 Mod VBA.CDbl(0)', 'VBA.CDbl(0)', "'11'"],
		['0 / VBA.CDbl(0)', 'VBA.CDbl(0)', "'6'"],
	])('reports %s', (expression, span, error) => {
		const source = wrap(expression);
		expectDiagnostic(source, analyze(source), 'division-by-zero', { span, message: error });
	});

	it.each([
		'10 / VBA.CDbl(0.4)',
		'10 / VBA.CLng(0.6)',
		'10 / Other.CDbl(0)',
		'10 / Other.VBA.CDbl(0)',
		'10 / VBA.CDbl(0).Value',
		'10 / VBA.CDbl(0)(1)',
	])('does not report %s as a zero divisor', (expression) => {
		expect(byCode(analyze(wrap(expression)), 'division-by-zero')).toHaveLength(0);
	});
});
