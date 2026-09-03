// github.com/WilliamSmithEdward/xlide_vscode/issues/62.
//
// The quick fix added in #59 typed the right-hand side of EVERY assignment
// before asking whether the target was undeclared, and typing one costs a bind
// of the whole module. A module whose variables are all declared - which is
// most modules - paid that for no findings, and the cost grew with the square
// of the module, so a 21,000-line module went from half a second to 37.
//
// This is a ratio test, not a stopwatch: absolute times vary by machine, but
// quadratic growth does not hide in a ratio. It follows the shape of the
// pass-cost guard that caught the original.

import { describe, expect, it } from 'vitest';
import { analyzeModule } from '../src/analyzer';
import { analyzeProjectModule, projectOptions } from './diagnostics/helpers';

/**
 * Procedures whose assignments all MISS the integer-literal fast path, which
 * is what hid the regression in synthetic modules full of `n = 1`.
 */
function module(procedures: number): string {
	const parts = ['Option Explicit', ''];
	for (let i = 0; i < procedures; i++) {
		parts.push(
			`Public Function V${i}(ByVal seed As Long) As Long`,
			'    Dim total As Long',
			'    total = seed * 4',
			'    total = total - CLng(seed)',
			`    V${i} = total`,
			'End Function',
			'',
		);
	}
	return parts.join('\r\n');
}

function medianMs(fn: () => void, runs: number): number {
	const samples: number[] = [];
	for (let i = 0; i < runs; i++) {
		const started = performance.now();
		fn();
		samples.push(performance.now() - started);
	}
	samples.sort((a, b) => a - b);
	return samples[Math.floor(samples.length / 2)];
}

describe('analysis cost stays linear in module size', () => {
	it('does not grow superlinearly across an 8x size step', () => {
		// Project options are what switch the rule ON: without knownIdentifiers
		// it never runs, and a probe that omits them measures nothing. That is
		// how the first version of this guard passed against the regression.
		const analyze = (source: string) => {
			const options = projectOptions([{ moduleName: 'PassCost', source }], 'PassCost');
			return () => {
				analyzeModule(source, { moduleName: 'PassCost', ...options });
			};
		};
		const small = module(200);
		const large = module(1600);
		analyze(small)();
		analyze(large)();

		const smallMs = medianMs(analyze(small), 3);
		const largeMs = medianMs(analyze(large), 3);
		// Measured on this ladder: 7.7x with the fix, 56.5x without. Linear is
		// ~8x. The ceiling sits between them with room for a loaded machine.
		expect(largeMs / Math.max(smallMs, 1)).toBeLessThan(20);
	}, 300000);

	it('still attaches a typed declaration to a name that IS undeclared', () => {
		const source = [
			'Option Explicit', 'Sub T()', '    Dim n As Long', '    n = 1',
			'    missing = n * 2', 'End Sub',
		].join('\r\n') + '\r\n';
		const hits = analyzeProjectModule(source, [{ moduleName: 'M', source }], 'M')
			.filter((d) => d.code === 'undeclared-variable');
		expect(hits).toHaveLength(1);
		expect(hits[0].data?.declareVariable?.variableName).toBe('missing');
		expect(hits[0].data?.declareVariable?.declaredType).toBeTruthy();
	});
});
