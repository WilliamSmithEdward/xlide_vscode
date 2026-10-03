import { afterEach, describe, expect, it, vi } from 'vitest';
import { checkUndeclaredVariables } from '../src/analyzer/diagnostics/rules/undeclared';
import { parseModule } from '../src/analyzer/parser/parseModule';
import { buildModuleSymbols } from '../src/analyzer/symbols/buildModuleSymbols';
import * as sourceScan from '../src/vbaSourceScan';

type Hit = Parameters<Parameters<typeof checkUndeclaredVariables>[13]>;
function fixture(source: string) {
	const mod = parseModule(source);
	const symbols = buildModuleSymbols('M', 'standard', source, { parsedModule: mod });
	const run = () => {
		const hits: Hit[] = [];
		checkUndeclaredVariables(source, mod, symbols, undefined, new Set(), undefined,
			undefined, undefined, undefined, 'standard', undefined, undefined, undefined,
			(...hit) => { hits.push(hit); });
		return hits;
	};
	return { mod, run };
}
afterEach(() => { vi.restoreAllMocks(); });

describe('undeclared declaration insertion setup', () => {
	it('searches the body once rather than once per missing assignment', () => {
		const count = 1000;
		const source = ['Option Explicit', 'Sub P()',
			...Array.from({ length: count }, (_, i) => `Dim D${i} As Object`),
			...Array.from({ length: count }, (_, i) => `Set Missing${i} = Nothing`),
			'End Sub'].join('\n');
		const { mod, run } = fixture(source);
		const proc = mod.members.find(member => member.kind === 'Procedure');
		if (!proc || proc.kind !== 'Procedure') throw new Error('Missing procedure');
		let reads = 0;
		proc.body = new Proxy(proc.body, {
			get(target, key, receiver) {
				if (typeof key === 'string' && /^\d+$/.test(key)) reads++;
				return Reflect.get(target, key, receiver);
			},
		});
		const hits = run();
		expect(hits).toHaveLength(count);
		expect(hits.every(hit => hit[3]?.declareVariable?.edit.span.start === source.indexOf('Set Missing0'))).toBe(true);
		expect(reads).toBeLessThan(count * 30);
	});

	it.each(['\n', '\r\n'])('keeps each procedure insertion point, indent, type and %j line ending', eol => {
		const source = ['Option Explicit', 'Sub A()', '    Dim d As Long', "    ' comment",
			'    first = 1', '    Set missingObject = Nothing', 'End Sub', 'Sub B()',
			'\tthird = "text"', 'End Sub'].join(eol) + eol;
		const { run } = fixture(source);
		const detect = vi.spyOn(sourceScan, 'detectEol');
		const fixes = run().map(hit => hit[3]?.declareVariable);
		expect(fixes).toEqual([
			{ variableName: 'first', declaredType: 'Long', edit: {
				span: { start: source.indexOf('    first'), end: source.indexOf('    first') },
				newText: `    Dim first As Long${eol}` } },
			{ variableName: 'missingObject', declaredType: 'Object', edit: {
				span: { start: source.indexOf('    first'), end: source.indexOf('    first') },
				newText: `    Dim missingObject As Object${eol}` } },
			{ variableName: 'third', declaredType: 'String', edit: {
				span: { start: source.indexOf('\tthird'), end: source.indexOf('\tthird') },
				newText: `\tDim third As String${eol}` } },
		]);
		expect(detect).toHaveBeenCalledTimes(1);
	});

	it('does not prepare edits for declared targets or read-only diagnostics', () => {
		const { run } = fixture('Option Explicit\nSub P()\nDim d As Object\nSet d = Nothing\nDebug.Print missing\nEnd Sub');
		const detect = vi.spyOn(sourceScan, 'detectEol');
		const hits = run();
		expect(hits).toHaveLength(1);
		expect(hits[0][3]).toBeUndefined();
		expect(detect).not.toHaveBeenCalled();
	});
});
