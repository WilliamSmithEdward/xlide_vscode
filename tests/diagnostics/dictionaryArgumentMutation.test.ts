import { describe, expect, it } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

function errors(body: string, extra = ''): string[] {
	const source = `Option Explicit\nSub Main()\nDim d As Object\nSet d = CreateObject("Scripting.Dictionary")\n${body}\nEnd Sub\n${extra}`;
	return analyzeModule(source).filter(d => d.severity === 'error').map(d => d.code);
}

const RESET = 'Function Reset(d As Object) As Long\nd.RemoveAll\nReset = 1\nEnd Function\n';

describe('Dictionary Add argument evaluation', () => {
	it.each([
		'd.Add "k", 1\nd.Add "x", Reset(d)\nd.Add "k", 2',
		'd.Add "k", 1\nd.Add "k", Reset(d)',
		'd.Add "k", 1\nd.Add "k", Reset(d) + 0',
	])('discards keys before a helper can clear them: %s', body => {
		expect(errors(body, RESET)).toEqual([]);
	});

	it.each(['d("b")', 'd.Item("b")'])('discards contents when an item read can add a key: %s', item => {
		expect(errors(`d.Add "a", ${item}\nd.Remove "b"`)).toEqual([]);
	});

	it('discards state of other dictionaries passed to arguments', () => {
		expect(errors('Dim other As Object\nSet other = CreateObject("Scripting.Dictionary")\nother.Add "k", 1\nd.Add "x", Reset(other)\nother.Add "k", 2', RESET)).toEqual([]);
	});

	it('does not treat a bare property getter as a literal value', () => {
		expect(errors('d.Add "k", 1\nd.Add "k", ResetValue', 'Property Get ResetValue() As Long\nResetValue = 1\nEnd Property\n')).toEqual([]);
	});

	it.each(['1', '"value"', '-1', 'True'])('retains duplicate-key checks with a literal value: %s', item => {
		expect(errors(`d.Add "k", ${item}\nd.Add "k", 2`)).toEqual(['collection-key-in-use']);
	});
});
