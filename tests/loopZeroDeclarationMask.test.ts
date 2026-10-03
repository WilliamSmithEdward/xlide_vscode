import { afterEach, describe, expect, it, vi } from 'vitest';
import { loopCountersAt } from '../src/analyzer/diagnostics/loopCounters';
import { parseModule } from '../src/analyzer/parser/parseModule';

function fixture(declarations: string, before = '', after = '') {
	const source = `Sub P()\n${declarations}\n${before}Do While i < 3\nx = i * 1\ni = i + 1\nLoop\n${after}End Sub`;
	const proc = parseModule(source).members.find(member => member.kind === 'Procedure');
	if (!proc || proc.kind !== 'Procedure') throw new Error('Missing procedure');
	return { source, proc };
}
afterEach(() => { vi.restoreAllMocks(); });

describe('zero-start declaration masking', () => {
	it('copies the procedure prefix once rather than once per declaration', () => {
		const { source, proc } = fixture('Dim i As Long\n' + Array.from({ length: 1000 }, (_, i) => `Dim D${i} As Long`).join('\n'));
		const original = String.prototype.slice;
		let copied = 0;
		vi.spyOn(String.prototype, 'slice').mockImplementation(function (this: string, start, end) {
			const result = original.call(this, start, end);
			if (String(this).length > 1000) copied += result.length;
			return result;
		});
		const facts = loopCountersAt(source, [...proc.body]);
		expect(facts.size).toBe(1);
		expect([...facts.values()][0].get('i')?.first.offset).toBe(0);
		expect(copied).toBeLessThan(source.length * 20);
	});

	it.each([
		['Dim i As Long', '', '', true],
		['Dim padding As Long, _\n i As Long', '', '', true],
		['Dim i As Long\nDim d As Long', "' unrelated comment\n", '', true],
		['', '', 'Dim i As Long\n', true],
		['Dim i As Long', 'x = i\n', '', false],
		['Dim i As Long', "' mention i in comment\n", '', false],
		['Static i As Long', '', '', false],
		['Dim i As Object', '', '', false],
		['Dim i(3) As Long', '', '', false],
		['Dim i As Long', '', 'GoTo Again\nAgain:\n', false],
	] as const)('preserves zero-start eligibility for %s', (declarations, before, after, known) => {
		const { source, proc } = fixture(declarations, before, after);
		expect(loopCountersAt(source, proc.body).size > 0).toBe(known);
	});
});
