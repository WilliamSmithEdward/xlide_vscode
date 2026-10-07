import { describe, expect, it } from 'vitest';
import { parseModule } from '../src/analyzer/parser/parseModule';
import { buildModuleSymbols } from '../src/analyzer/symbols/buildModuleSymbols';
import { objectLetStateAt } from '../src/analyzer/diagnostics/rules/objectState';
import type { MemberCompletionContext } from '../src/analyzer/completion/memberAccess';

describe('object-state bound symbol contexts', () => {
	for (const eol of ['\n', '\r\n', '\r']) {
		for (const change of ['auto-instantiated', 'static'] as const) {
			it(`refreshes a retained procedure for ${change} symbols with ${JSON.stringify(eol)}`, () => {
				const source = ['Sub Run()', 'Dim value As Object', 'value = 3', 'End Sub'].join(eol);
				const module = parseModule(source);
				const proc = module.members.find(node => node.kind === 'Procedure');
				if (proc?.kind !== 'Procedure') { throw new Error('Missing procedure'); }
				const base = buildModuleSymbols('M', 'standard', source, { parsedModule: module });
				const changed = buildModuleSymbols('M', 'standard', source, { parsedModule: module });
				const local = changed.root.children?.find(symbol => symbol.name === 'Run')?.children?.find(symbol => symbol.name === 'value');
				if (!local) { throw new Error('Missing local'); }
				if (change === 'auto-instantiated') { local.isAutoInstantiated = true; }
				else { local.visibility = 'Static'; }
				const context: MemberCompletionContext = {};
				const offset = source.indexOf('value =');
				const state = (symbols: typeof base, ctx = context) => objectLetStateAt(source, module, proc, symbols, ctx, undefined, offset);
				// Independently recomputed controls establish the two symbol contexts.
				expect(state(base, {})).toBe('unset');
				expect(state(changed, {})).toBe('unknown');
				for (const [symbols, expected] of [[base, 'unset'], [changed, 'unknown'], [base, 'unset']] as const) {
					expect(state(symbols)).toBe(expected);
					expect(state(symbols)).toBe(expected);
				}
			});
		}
	}
	it('reuses each symbol snapshot when alternating retained contexts', () => {
		const source = 'Sub Reuse()\nDim item As Object\nitem = 1\nEnd Sub';
		const module = parseModule(source), proc = module.members.find(node => node.kind === 'Procedure');
		if (proc?.kind !== 'Procedure') { throw new Error('Missing procedure'); }
		const snapshots = [buildModuleSymbols('M', 'standard', source, { parsedModule: module }), buildModuleSymbols('M', 'standard', source, { parsedModule: module })];
		let reads = 0;
		for (const symbols of snapshots) {
			const local = symbols.root.children?.find(symbol => symbol.name === 'Reuse')?.children?.find(symbol => symbol.name === 'item');
			if (!local) { throw new Error('Missing local'); }
			Object.defineProperty(local, 'isAutoInstantiated', { get: () => { reads++; return false; } });
		}
		const context: MemberCompletionContext = {};
		const state = (symbols: typeof snapshots[number]) => objectLetStateAt(source, module, proc, symbols, context, undefined, source.indexOf('item ='));
		for (const symbols of snapshots) { expect(state(symbols)).toBe('unset'); }
		expect(reads).toBeGreaterThan(0);
		reads = 0;
		for (let round = 0; round < 10; round++) {
			for (const symbols of snapshots) { expect(state(symbols)).toBe('unset'); }
		}
		expect(reads).toBe(0);
	});
});
