import { describe, expect, it } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { tokenizeCached } from '../../src/analyzer/lexer/tokenize';
import { parseModule } from '../../src/analyzer/parser/parseModule';
import { statementTokensCached } from '../../src/analyzer/lexer/tokenHelpers';

// Issue #45. The analyzer's memos find their entry by comparing a whole source
// string. Within one pass that is free, because the caller hands back the very
// instance the entry was stored under and V8 settles `===` on the pointer.
// Across passes it is not: a host that re-materialises module text - a worker
// boundary, a pipe, a re-read - produces a NEW string with the same content, so
// every lookup compared the whole module. statementTokensCached is asked
// hundreds of thousands of times per pass, which made every analysis after the
// first quadratic in module size.

/** A distinct string instance with identical content, as a transport produces. */
function fresh(source: string): string {
	return Buffer.from(source, 'utf8').toString('utf8');
}

function build(procedures: number, name = 'Proc'): string {
	const lines: string[] = ['Option Explicit', ''];
	for (let i = 0; i < procedures; i += 1) {
		lines.push(
			`Public Sub ${name}${i}(ByVal a As Long)`,
			'    Dim total As Long',
			'    Dim ws As Worksheet',
			'    Set ws = ActiveWorkbook.Worksheets(1)',
			`    total = a + ${i}`,
			'    If total > 0 Then',
			'        ws.Range("A1").Value = total',
			'    End If',
			'End Sub',
			'',
		);
	}
	return lines.join('\n');
}

describe('the analyzer memos hit on content, not only on instance', () => {
	it('serves a fresh instance of the same text from the cache', () => {
		const source = build(3);
		const first = tokenizeCached(source);
		// Identity, not just equality: a miss would build a second array.
		expect(tokenizeCached(fresh(source))).toBe(first);
		expect(tokenizeCached(fresh(source))).toBe(first);

		const module = parseModule(source);
		expect(parseModule(fresh(source))).toBe(module);

		const span = { start: 0, end: source.indexOf('\n') };
		const statement = statementTokensCached(source, span);
		expect(statementTokensCached(fresh(source), span)).toBe(statement);
	});

	it('reports the same findings whichever instance it is handed', () => {
		const source = build(40);
		const findings = () => analyzeModule(fresh(source), { host: 'excel' })
			.map((diagnostic) => `${diagnostic.code}@${diagnostic.span.start}`);
		const baseline = analyzeModule(source, { host: 'excel' })
			.map((diagnostic) => `${diagnostic.code}@${diagnostic.span.start}`);
		expect(findings()).toEqual(baseline);
		expect(findings()).toEqual(baseline);
	});

	it('analyses a repeat pass in the same order of time as the first', () => {
		// The defect showed as a repeat pass costing MULTIPLES of the cold one -
		// 18x on a 64,802-line module, growing with its square. At 1,000
		// procedures it still costs about 4-7x, while a sound repeat costs about
		// as much as the cold pass. The margin is deliberately loose: this
		// guards the shape, not a millisecond budget.
		//
		// A busy machine can stall any one pass, so a single cold sample is not
		// a baseline. Each round times a cold pass and then a repeat over a fresh
		// instance of the same text, and the median of three rounds decides. Each
		// round names its procedures afresh, so no memo can reuse an earlier
		// round's tokens or parse and every cold pass is a full one.
		const time = (source: string): number => {
			const started = process.hrtime.bigint();
			analyzeModule(fresh(source), { host: 'excel' });
			return Number(process.hrtime.bigint() - started) / 1e6;
		};
		const round = (index: number): number => {
			const source = build(1000, `Round${index}Proc`);
			const cold = time(source);
			return time(source) / cold;
		};
		time(build(40, 'WarmUp')); // reduces JIT start-up in the first cold pass
		const ratios = [round(0), round(1)];
		// Two rounds on the same side of 3x already settle the median of three.
		if ((ratios[0] < 3) !== (ratios[1] < 3)) { ratios.push(round(2)); }
		ratios.sort((a, b) => a - b);
		expect(ratios[1]).toBeLessThan(3);
	}, 120_000);
});
