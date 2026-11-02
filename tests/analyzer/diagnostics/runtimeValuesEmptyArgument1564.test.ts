// Issue #1564. An empty argument where the runtime value rule reads an
// integer, such as `Mid(s, ) = "z"`, threw inside the rule. Its visitor was
// then switched off for the rest of the module, so findings in later
// procedures went missing. Each shape below is in procedure A; procedure B's
// `Err.Raise 0` must still be reported.

import { describe, expect, it } from 'vitest';
import { analyzeVbaModuleSource } from '../../../src/vbaModuleAnalysis';

/** The trimmed line of each runtime-argument-value finding. */
function findingLines(procedureA: string[]): string[] {
	const source = [
		'Option Explicit',
		'Public Sub A()',
		'    Dim s As String',
		'    Dim d As Date',
		'    s = "abc"',
		...procedureA.map((line) => `    ${line}`),
		'End Sub',
		'Public Sub B()',
		'    Err.Raise 0',
		'End Sub',
		'',
	].join('\r\n');
	const lines = source.split('\r\n');
	return analyzeVbaModuleSource({ source, moduleName: 'Module1', moduleKind: 'standard' }).diagnostics
		.filter((diagnostic) => diagnostic.code === 'runtime-argument-value')
		.map((diagnostic) => lines[source.slice(0, diagnostic.span.start).split('\r\n').length - 1].trim());
}

describe('an empty argument the runtime value rule reads (issue #1564)', () => {
	it('reports Err.Raise 0 when procedure A has no empty argument', () => {
		expect(findingLines(['Mid(s, 2) = "z"'])).toEqual(['Err.Raise 0']);
	});

	it.each([
		['a Mid statement with an empty start', 'Mid(s, ) = "z"'],
		['a Mid$ statement with an empty start', 'Mid$(s, ) = "z"'],
		['a MidB statement with an empty start', 'MidB(s, ) = "z"'],
		['a MidB$ statement on a literal with an empty start', 'MidB$("abc", ) = "z"'],
		['a Mid statement with an empty start and a length', 'Mid(s, , 1) = "z"'],
		['DateAdd with an empty number', 'd = DateAdd("d", , d)'],
		['DateSerial with an empty month inside DateAdd', 'd = DateAdd("d", 1, DateSerial(2026, , 1))'],
		['DateSerial with an empty day inside DateAdd', 'd = DateAdd("d", 1, DateSerial(2026, 1, ))'],
	])('still reports Err.Raise 0 after %s', (_name, line) => {
		expect(findingLines([line])).toContain('Err.Raise 0');
	});
});
