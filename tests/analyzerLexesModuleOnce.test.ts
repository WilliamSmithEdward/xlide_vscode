import { describe, expect, it } from 'vitest';
import { analyzeModule } from '../src/analyzer';
import {
	startTokenizeMissLogForTests,
	stopTokenizeMissLogForTests,
} from '../src/analyzer/lexer/tokenize';

// Issue #139. The statement-token cache is keyed by source string and holds
// two of them; the module token cache holds eight. Rules that lex an
// expression the parser carried as its own string (an If condition, a For Each
// source, an Enum member value, a parameter default, a Const value) must lex
// it uncached. Sending those strings through the cached lexer evicted the
// module, so the next ordinary statement re-lexed the whole module: once per
// procedure, quadratic in module size, and 17x slower on real projects.

function build(procedures: number): string {
	const lines: string[] = [
		'Option Explicit',
		'',
		'Public Enum Colours',
	];
	for (let i = 0; i < procedures; i += 1) {
		lines.push(`    Colour${i} = ${i * 3}`);
	}
	lines.push('End Enum', '');
	for (let i = 0; i < procedures; i += 1) {
		lines.push(
			`Private Const LIMIT${i} As Long = ${i} * 7 + 1`,
			`Public Sub Proc${i}(ByVal a As Long, Optional ByVal b As Integer = ${i + 100})`,
			'    Dim total As Long',
			'    Dim item As Variant',
			`    total = a + LIMIT${i}`,
			`    If total <> ${i} Then`,
			`        total = b \\ total`,
			`    ElseIf a > ${i + 1} Then`,
			'        total = 0',
			'    End If',
			`    For Each item In Array(${i}, total)`,
			'        total = total + item',
			'    Next item',
			'End Sub',
			'',
		);
	}
	return lines.join('\r\n');
}

describe('one analysis pass lexes the module once (issue #139)', () => {
	it('does not evict the module while lexing derived expression strings', () => {
		const source = build(60);
		startTokenizeMissLogForTests();
		let misses: number[];
		try {
			analyzeModule(source, { host: 'excel' });
		} finally {
			misses = stopTokenizeMissLogForTests();
		}
		expect(misses.filter((length) => length === source.length)).toHaveLength(1);
	});
});
