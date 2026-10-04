import {parseModule} from '../src/analyzer/parser/parseModule';
import {collectModuleLiteralIntegerConstants} from '../src/analyzer/diagnostics/constExpr';
import {ProjectIndex} from '../src/analyzer/symbols/projectIndex';
import {describe, expect, it, vi} from 'vitest';
import {resolveRawIntegerConstants, evaluateIntegerConstantExpression} from '../src/analyzer/constants/integerConstantExpression';

describe('constant dependency stack', () => {
	it.each(['alias', 'arithmetic', 'cycle'] as const)('resolves 10,000 forward %s dependencies without native stack growth', mode => {
		const count = 10000;
		const raw = new Map(Array.from({length: count}, (_, i) => ['c' + i, i === count - 1 ? (mode === 'cycle' ? 'c0' : '1') : 'c' + (i + 1) + (mode === 'arithmetic' ? ' + 1' : '')]));
		const expected = Array.from({length: count}, (_, i) => ['c' + (count - 1 - i), mode === 'cycle' ? undefined : mode === 'alias' ? 1 : i + 1]);
		if (mode === 'cycle') { expected.unshift(['c0', undefined]); expected.pop(); }
		expect([...resolveRawIntegerConstants(raw)]).toEqual(expected);
	});

	it('does not replay a wide expression while resolving its dependencies', () => {
		const count = 10000;
		const raw = new Map([['root', Array.from({length: count}, (_, i) => 'unit + c' + i).join(' + ')], ...Array.from({length: count}, (_, i) => ['c' + i, '1'])]);
		const read = vi.spyOn(raw, 'get');
		const base = new Map([['unit', 1]]), baseRead = vi.spyOn(base, 'get');
		const values = resolveRawIntegerConstants(raw, base);
		expect([...values]).toEqual([...Array.from({length: count}, (_, i) => ['c' + i, 1]), ['root', count * 2]]);
		expect(baseRead.mock.calls).toEqual(Array.from({length: count}, () => ['unit']));
		expect(read.mock.calls).toEqual([['root'], ...Array.from({length: count}, (_, i) => ['c' + i])]);
	});

	it('keeps expression nesting independent of constant dependency nesting', () => {
		const count = 100, wrap = (text: string) => '('.repeat(100) + text + ')'.repeat(100);
		const raw = new Map(Array.from({length: count}, (_, i) => ['c' + i, wrap(i === count - 1 ? '7' : 'c' + (i + 1))]));
		expect([...resolveRawIntegerConstants(raw)]).toEqual(Array.from({length: count}, (_, i) => ['c' + (count - 1 - i), 7]));
	});

	it('preserves early unknowns, ambiguity, cycles, function keys and base lookup order', () => {
		const raw = new Map<string, string | undefined>([['unknown', 'missing + hidden'], ['hidden', '3'], ['dup', undefined], ['self', 'self'], ['qualified', 'Module.Value + F(Arg)'], ['upper', 'dup + 1'], ['UpperOnly', '2']]);
		const base = new Map([['module.value', 10], ['arg', 2], ['f(2)', 4]]), read = vi.spyOn(base, 'get');
		expect([...resolveRawIntegerConstants(raw, base)]).toEqual([['unknown', undefined], ['hidden', 3], ['dup', undefined], ['self', undefined], ['qualified', 14], ['upper', undefined]]);
		expect(read.mock.calls).toEqual([['missing'], ['module.value'], ['arg'], ['f(2)'], ['upperonly']]);
	});

	it('preserves lookup order and Not parsing for one-token and qualified paths', () => {
		const seen: string[] = [], lookup = {get: (name: string) => {seen.push(name); return 2;}};
		expect(evaluateIntegerConstantExpression('Not', lookup)).toBeUndefined();
		expect(evaluateIntegerConstantExpression('Not . Member', lookup)).toBeUndefined();
		expect(evaluateIntegerConstantExpression('[Not]', lookup)).toBe(2);
		expect(evaluateIntegerConstantExpression('Module.Value', lookup)).toBe(2);
		expect(evaluateIntegerConstantExpression('name + 1', lookup)).toBe(3);
		expect(seen).toEqual(['not', 'module.value', 'name']);
	});
});




it.each(['\n', '\r\n', '\r'])('resolves forward module constants through diagnostic and project callers with EOL %j', eol => {
	const count = 2000;
	const source = Array.from({length: count}, (_, i) => `Public Const C${i} As Long = ${i === count - 1 ? '1' : 'C' + (i + 1) + ' + 1'}`).join(eol);
	const expected = Array.from({length: count}, (_, i) => ['c' + (count - 1 - i), i + 1]);
	expect([...collectModuleLiteralIntegerConstants(parseModule(source), undefined)]).toEqual(expected);
	const project = new ProjectIndex();
	project.setModule({moduleName: 'Constants', moduleKind: 'standard', source});
	const exported = project.visibleExternalIntegerConstantExpressions('Caller');
	expect(exported.size).toBe(count * 2);
	for (let i = 0; i < count; i++) {
		expect(exported.get('c' + i)).toBe(String(count - i));
		expect(exported.get('constants.c' + i)).toBe(String(count - i));
	}
});
