// github.com/WilliamSmithEdward/xlide_vscode/issues/322.
//
// Two procedure shapes grew with the square of their size. A long If/ElseIf
// chain: the handler-flow checks read the procedure's label references once
// per branch, 2000 arms taking 1.7 s. And many locals: every statement holds
// every local's reaching value, and the value rules built, and walked, the
// whole map at every statement, 2000 locals taking 2.4 s.
//
// Ratio tests, as undeclaredVariableCost.test.ts is: absolute times vary by
// machine, but quadratic growth does not hide in a ratio.

import { describe, expect, it } from 'vitest';
import { analyzeModule } from '../src/analyzer';

// Measured on this ladder: 4.5 (chain) and 4.9 (locals) with the fix, 12.7
// and 14.8 without. Linear is 4x. The ceiling sits between them with room
// for a loaded machine.
const RATIO_CEILING = 9;

function elseIfChain(arms: number): string {
	const lines = ['Option Explicit', 'Function F(ByVal k As Long) As Long', '    Dim t As Long', '    If k = 0 Then', '        t = t + 0'];
	for (let n = 1; n <= arms; n++) {
		lines.push(`    ElseIf k = ${n} Then`, `        t = t + ${n}`);
	}
	lines.push('    End If', '    F = t', 'End Function', '');
	return lines.join('\r\n');
}

function manyLocals(count: number): string {
	const lines = ['Option Explicit', 'Function F() As Long'];
	for (let n = 0; n < count; n++) {
		lines.push(`    Dim v${n} As Long`);
	}
	for (let s = 0; s < count; s++) {
		lines.push(`    v${s % count} = v${(s * 7 + 3) % count} + ${s % 50}`);
	}
	lines.push('    F = v0', 'End Function', '');
	return lines.join('\r\n');
}

function medianMs(source: string, runs: number): number {
	analyzeModule(source);
	const samples: number[] = [];
	for (let i = 0; i < runs; i++) {
		const started = performance.now();
		analyzeModule(source);
		samples.push(performance.now() - started);
	}
	samples.sort((a, b) => a - b);
	return samples[Math.floor(samples.length / 2)];
}

describe('analysis cost stays near linear in a procedure\'s size (issue #322)', () => {
	it('for a long If/ElseIf chain across a 4x step', () => {
		const ratio = medianMs(elseIfChain(2000), 3) / Math.max(medianMs(elseIfChain(500), 3), 1);
		expect(ratio).toBeLessThan(RATIO_CEILING);
	}, 300000);

	it('for many locals across a 4x step', () => {
		const ratio = medianMs(manyLocals(2000), 3) / Math.max(medianMs(manyLocals(500), 3), 1);
		expect(ratio).toBeLessThan(RATIO_CEILING);
	}, 300000);
});
