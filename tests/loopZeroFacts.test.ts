import { describe, expect, it } from 'vitest';
import { loopCountersAt } from '../src/analyzer/diagnostics/loopCounters';
import { parseModule } from '../src/analyzer/parser/parseModule';

function fixture(source: string) {
	const proc = parseModule(source).members.find(member => member.kind === 'Procedure');
	if (!proc || proc.kind !== 'Procedure') throw new Error('Missing procedure');
	return proc;
}
function loop(name: string): string {
	return `Do While ${name} < 3\nx = ${name} * 1\n${name} = ${name} + 1\nLoop`;
}

describe('procedure-wide zero-start facts', () => {
	it('reads declaration names once across distinct zero-initialized loops', () => {
		const count = 300;
		const source = ['Sub P()', ...Array.from({ length: count }, (_, i) => `Dim c${i} As Long`),
			...Array.from({ length: count }, (_, i) => loop(`c${i}`)), 'End Sub'].join('\n');
		const proc = fixture(source);
		let reads = 0;
		for (const node of proc.body) {
			if (node.kind !== 'VariableGroup') continue;
			for (const decl of node.declarations) {
				const name = decl.name;
				Object.defineProperty(decl, 'name', { get: () => { reads++; return name; } });
			}
		}
		expect(loopCountersAt(source, proc.body).size).toBe(count);
		expect(reads).toBeLessThan(count * 10);
	});

	it('ignores unrelated earlier loop names while rejecting a reused counter', () => {
		const source = ['Sub P()', 'Dim c1 As Long, c10 As Long', loop('c1'), loop('c10'), loop('c1'), 'End Sub'].join('\n');
		const proc = fixture(source);
		const names = [...loopCountersAt(source, proc.body).values()].map(counters => [...counters.keys()]);
		expect(names).toEqual([['c1'], ['c10']]);
	});

	it.each(['État', 'İ', 'µ', 'ΟΣ'])('preserves prefix checks for Unicode counter %s', name => {
		const source = ['Sub P()', `Dim ${name} As Long`, 'x = "ABCİDEF"', loop(name), 'End Sub'].join('\n');
		expect(loopCountersAt(source, fixture(source).body).size).toBe(1);
		const read = source.replace('x = "ABCİDEF"', `x = ${name}`);
		expect(loopCountersAt(read, fixture(read).body).size).toBe(0);
	});

	it('keeps first declaration within a group and last group eligibility', () => {
		const first = ['Sub P()', 'Dim i As Long, i As Object', loop('i'), 'End Sub'].join('\n');
		expect(loopCountersAt(first, fixture(first).body).size).toBe(1);
		const last = first.replace('Dim i As Long, i As Object', 'Dim i As Long\nDim i As Object');
		expect(loopCountersAt(last, fixture(last).body).size).toBe(0);
	});
});
