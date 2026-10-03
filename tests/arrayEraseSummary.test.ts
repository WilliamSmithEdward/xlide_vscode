import { afterEach, describe, expect, it, vi } from 'vitest';
import { checkUnallocatedDynamicArrayAccess } from '../src/analyzer/diagnostics/rules/arrays';
import { parseModule } from '../src/analyzer/parser/parseModule';
import { buildModuleSymbols } from '../src/analyzer/symbols/buildModuleSymbols';
import { createConditionalActivityTracker } from '../src/analyzer/conditional/conditionalCompilation';

type Hit = Parameters<Parameters<typeof checkUnallocatedDynamicArrayAccess>[4]>;
function fixture(source: string, vba7?: boolean) {
	const mod = parseModule(source);
	const environment = vba7 === undefined ? undefined : { compilerConstants: { VBA7: vba7 } };
	const symbols = buildModuleSymbols('M', 'standard', source, { parsedModule: mod, conditionalCompilation: environment });
	const activity = environment ? createConditionalActivityTracker(mod, environment) : undefined;
	const run = () => {
		const hits: Hit[] = [];
		checkUnallocatedDynamicArrayAccess(source, mod, symbols, activity, (...hit) => { hits.push(hit); });
		return hits;
	};
	return { mod, run };
}
function module(body: string, parameter = 'ByRef p() As Long'): string {
	return `Sub Free(${parameter})\n${body}\nEnd Sub\nSub Caller()\nDim a() As Long\nReDim a(1)\nFree a\nx = UBound(a)\nEnd Sub`;
}
afterEach(() => { vi.restoreAllMocks(); });

describe('array-erasure procedure summaries', () => {
	it('does not read statement spans for procedures without eligible array parameters', () => {
		const { mod, run } = fixture('Sub P()\n' + 'x = 1\n'.repeat(1000) + 'End Sub');
		const proc = mod.members.find(member => member.kind === 'Procedure');
		if (!proc || proc.kind !== 'Procedure') throw new Error('Missing procedure');
		let reads = 0;
		for (const node of proc.body) {
			const span = node.span;
			Object.defineProperty(node, 'span', { get: () => { reads++; return span; } });
		}
		expect(run()).toEqual([]);
		expect(reads).toBeLessThan(20);
	});

	it('collects mentions once across many array parameters', () => {
		const count = 200;
		const params = Array.from({ length: 25 }, (_, i) => `ByRef p${i}() As Long`).join(',');
		const { run } = fixture(`Sub Free(${params})\n${'x = 1\n'.repeat(count)}${Array.from({ length: 25 }, (_, i) => `Erase p${i}`).join('\n')}\nEnd Sub`);
		const original = String.prototype.toLowerCase;
		let reads = 0;
		vi.spyOn(String.prototype, 'toLowerCase').mockImplementation(function (this: string) {
			if (String(this) === 'x') reads++;
			return original.call(this);
		});
		expect(run()).toEqual([]);
		expect(reads).toBeLessThan(count * 10);
	});

	it.each([
		['x = p(0)\nErase p', 'ByRef p() As Long', 1],
		['Erase p\nx = p(0)', 'ByRef p() As Long', 0],
		['Exit Sub\nErase p', 'ByRef p() As Long', 0],
		['If True Then\nx = p(0)\nEnd If\nErase p', 'ByRef p() As Long', 0],
		['Erase p', 'ByVal p() As Long', 0],
		['Erase p', 'ByRef p As Variant', 0],
	] as const)('preserves the effect of %s', (body, parameter, hits) => {
		expect(fixture(module(body, parameter)).run()).toHaveLength(hits);
	});

	it.each([false, true])('keeps inactive nested mentions conservative with VBA7=%s', vba7 => {
		const nested = '#If VBA7 Then\nIf True Then\nx = p(0)\nEnd If\n#End If\nErase p';
		expect(fixture(module(nested), vba7).run()).toHaveLength(0);
		const leaf = '#If VBA7 Then\nx = p(0)\n#End If\nErase p';
		expect(fixture(module(leaf), vba7).run()).toHaveLength(1);
	});
});
