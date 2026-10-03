import { describe, expect, it } from 'vitest';
import { rememberProjectWrittenNames, untouchedModuleVariables, untouchedModuleVariablesIn } from '../src/analyzer/diagnostics/moduleState';
import { parseModule } from '../src/analyzer/parser/parseModule';
import { buildModuleSymbols } from '../src/analyzer/symbols/buildModuleSymbols';

function fixture(source: string) {
	const mod = parseModule(source);
	return { symbols: buildModuleSymbols('M', 'standard', source, { parsedModule: mod }),
		procs: mod.members.filter(member => member.kind === 'Procedure') };
}

describe('shared module-state facts', () => {
	it('does not rescan module declarations for every procedure and consumer', () => {
		const count = 200;
		const source = [...Array.from({ length: count }, (_, i) => `Private A${i} As Long`),
			...Array.from({ length: 20 }, (_, i) => `Sub P${i}()\nx = 1\nEnd Sub`)].join('\n');
		const { symbols, procs } = fixture(source);
		let reads = 0;
		for (const symbol of symbols.all.filter(symbol => symbol.kind === 'moduleVariable')) {
			const name = symbol.name;
			Object.defineProperty(symbol, 'name', { get: () => { reads++; return name; } });
		}
		for (let consumer = 0; consumer < 3; consumer++) {
			for (const proc of procs) expect(untouchedModuleVariablesIn(source, symbols, proc).size).toBe(count);
		}
		expect(reads).toBeLessThan(count * 10);
	});

	it('preserves module writes, exclusions and procedure-local shadowing', () => {
		const source = ['Private a As Long', 'Private b As Long', 'Private changed As Long',
			'Private auto As New Collection', 'Private fixed As String * 3', 'Public shared As Long',
			'Sub LocalScope()', 'Dim A As Long', 'changed = 1', 'End Sub',
			'Sub ParamScope(ByVal B As Long)', 'x = 1', 'End Sub', 'Sub Unshadowed()', 'x = 1', 'End Sub'].join('\n');
		const { symbols, procs } = fixture(source);
		rememberProjectWrittenNames(symbols, new Set());
		for (let consumer = 0; consumer < 3; consumer++) {
			expect([...untouchedModuleVariablesIn(source, symbols, procs[0]).keys()]).toEqual(['b', 'shared']);
			expect([...untouchedModuleVariablesIn(source, symbols, procs[1]).keys()]).toEqual(['a', 'shared']);
			expect([...untouchedModuleVariablesIn(source, symbols, procs[2]).keys()]).toEqual(['a', 'b', 'shared']);
		}
	});

	it('refreshes public state when project writes change or become unavailable', () => {
		const source = 'Public shared As Long\nSub P()\nx = 1\nEnd Sub';
		const { symbols, procs } = fixture(source);
		const names = () => [...untouchedModuleVariablesIn(source, symbols, procs[0]).keys()];
		expect(names()).toEqual([]);
		const writes = new Set<string>();
		rememberProjectWrittenNames(symbols, writes);
		expect(names()).toEqual(['shared']);
		writes.add('shared');
		rememberProjectWrittenNames(symbols, writes);
		expect(names()).toEqual([]);
		rememberProjectWrittenNames(symbols, new Set());
		expect(names()).toEqual(['shared']);
		rememberProjectWrittenNames(symbols, undefined);
		expect(names()).toEqual([]);
	});

	it('refreshes module write facts for changed source with the same declaration metadata', () => {
		const source = 'Private a As Long\nSub P()\nx = 1\nEnd Sub';
		const { symbols, procs } = fixture(source);
		expect([...untouchedModuleVariablesIn(source, symbols, procs[0]).keys()]).toEqual(['a']);
		const edited = source.replace('x = 1', 'a = 1');
		expect([...untouchedModuleVariablesIn(edited, symbols, procs[0]).keys()]).toEqual([]);
		expect([...untouchedModuleVariables(source, symbols).keys()]).toEqual(['a']);
	});
});
