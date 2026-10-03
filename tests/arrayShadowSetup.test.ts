import { describe, expect, it } from 'vitest';
import { checkFixedArraySubscriptBounds } from '../src/analyzer/diagnostics/rules/arrays';
import { parseModule } from '../src/analyzer/parser/parseModule';
import { buildModuleSymbols } from '../src/analyzer/symbols/buildModuleSymbols';

type Hit = Parameters<Parameters<typeof checkFixedArraySubscriptBounds>[4]>;
function fixture(source: string) {
	const mod = parseModule(source);
	const symbols = buildModuleSymbols('M', 'standard', source, { parsedModule: mod });
	const run = () => {
		const hits: Hit[] = [];
		checkFixedArraySubscriptBounds(source, mod, symbols, undefined, (...hit) => { hits.push(hit); });
		return hits;
	};
	return { symbols, run };
}

describe('module-array local shadow setup', () => {
	it('does not rebuild every local name for every module array', () => {
		const count = 500;
		const source = [...Array.from({ length: count }, (_, i) => `Dim A${i}(0 To 1) As Long`),
			'Sub P()', ...Array.from({ length: count }, (_, i) => `Dim L${i} As Long`),
			'A0(5) = 1', 'End Sub'].join('\n');
		const { symbols, run } = fixture(source);
		let reads = 0;
		for (const symbol of symbols.all.filter(symbol => symbol.kind === 'localVariable')) {
			const name = symbol.name;
			Object.defineProperty(symbol, 'name', { get: () => { reads++; return name; } });
		}
		const hits = run();
		expect(hits).toHaveLength(1);
		expect(source.slice(hits[0][2].start, hits[0][2].end)).toBe('5');
		expect(reads).toBeLessThan(count * 30);
	});

	it('keeps local, parameter and unshadowed module bounds separate per procedure', () => {
		const source = ['Dim Items(0 To 1) As Long', 'Dim Other(0 To 2) As Long',
			'Sub LocalScope()', 'Dim items(0 To 9) As Long', 'items(5) = 1', 'Other(5) = 1', 'End Sub',
			'Sub ParamScope(ByRef ITEMS() As Long)', 'ITEMS(5) = 1', 'Other(4) = 1', 'End Sub',
			'Sub ModuleScope()', 'Items(3) = 1', 'Other(1) = 1', 'End Sub'].join('\r\n');
		const hits = fixture(source).run();
		expect(hits.map(hit => hit[2])).toEqual(['Other(5)', 'Other(4)', 'Items(3)'].map(text => {
			const start = source.indexOf(text) + text.indexOf('(') + 1;
			return { start, end: start + 1 };
		}));
	});
});
