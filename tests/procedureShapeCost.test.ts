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

// Measured on this ladder, as bestRatio reads it: 3.8 to 4.6 (chain) and 4.5
// to 5.7 (locals) with the fix, idle or under load; 11 and 15 without. Linear
// is 4x. The ceiling sits between them with room for a loaded machine.
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

function timeMs(source: string): number {
	const started = performance.now();
	analyzeModule(source);
	return performance.now() - started;
}

/**
 * The large source's time over the small one's, each its fastest of `runs`.
 * Both are warmed before either is timed, and the samples alternate, so the
 * JIT warming as the runs go on (about 2x on the small source) speeds both
 * alike. Timing the large source first and the small one after it read the
 * small one warmer, and once lifted the ratio past the ceiling on a shared
 * CI runner (9.9, against 4 to 5 here). The fastest run is the one other
 * load interrupted least; quadratic growth still shows in it.
 */
function bestRatio(small: string, large: string, runs: number): number {
	timeMs(small);
	timeMs(large);
	let smallMs = Number.POSITIVE_INFINITY;
	let largeMs = Number.POSITIVE_INFINITY;
	for (let i = 0; i < runs; i++) {
		smallMs = Math.min(smallMs, timeMs(small));
		largeMs = Math.min(largeMs, timeMs(large));
	}
	return largeMs / Math.max(smallMs, 1);
}

describe('analysis cost stays near linear in a procedure\'s size (issue #322)', () => {
	it('for a long If/ElseIf chain across a 4x step', () => {
		expect(bestRatio(elseIfChain(500), elseIfChain(2000), 5)).toBeLessThan(RATIO_CEILING);
	}, 300000);

	it('for many locals across a 4x step', () => {
		expect(bestRatio(manyLocals(500), manyLocals(2000), 5)).toBeLessThan(RATIO_CEILING);
	}, 300000);
});
