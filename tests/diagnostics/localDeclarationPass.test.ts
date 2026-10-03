import { describe, expect, it, vi } from 'vitest';
import { checkLocalDeclarationOrder } from '../../src/analyzer/diagnostics/rules/localDeclarationOrder';
import { parseModule } from '../../src/analyzer/parser/parseModule';
import { buildModuleSymbols } from '../../src/analyzer/symbols/buildModuleSymbols';
import type { Span } from '../../src/analyzer/parser/nodes';

describe('local declaration pass', () => {
	it.each([true, false])('does not repeatedly enumerate all declarations (explicit=%s)', explicit => {
		const count = 400;
		const body = Array.from({ length: count }, (_, i) => explicit ? 'Debug.Print B' + i + '\n' : 'Const A' + i + ' = B' + i + '\n').join('');
		const source = (explicit ? 'Option Explicit\n' : '') + 'Sub Main()\n' + body + Array.from({ length: count }, (_, i) => 'Dim B' + i + ' As Long\n').join('') + 'End Sub';
		const module = parseModule(source);
		const symbols = buildModuleSymbols('Module', 'standard', source, { parsedModule: module });
		let visits = 0;
		const original = Map.prototype.values;
		const spy = vi.spyOn(Map.prototype, 'values').mockImplementation(function* (this: Map<unknown, unknown>) {
			for (const value of original.call(this)) { visits++; yield value; }
		});
		try {
			const push = vi.fn();
			checkLocalDeclarationOrder(source, module, symbols, undefined, undefined, undefined, push);
			expect(push).toHaveBeenCalledTimes(count);
			expect(push.mock.calls.every(call => call[0] === (explicit ? 'undeclaredVariable' : 'constValueNotConstant'))).toBe(true);
			expect(visits).toBeLessThan(count * 20);
		} finally { spy.mockRestore(); }
	});

	it.each(['\n', '\r\n'])('preserves physical Const line coverage for colon-separated uses with %j', eol => {
		const source = ['Sub Main()', 'Const A = B: Const C = D: Debug.Print E', 'Dim B As Long, D As Long, E As Long', 'End Sub'].join(eol);
		const module = parseModule(source);
		const symbols = buildModuleSymbols('Module', 'standard', source, { parsedModule: module });
		const found: Array<{ kind: string; span: Span }> = [];
		checkLocalDeclarationOrder(source, module, symbols, undefined, undefined, undefined, (kind, _message, span) => { found.push({ kind, span }); });
		expect(found.map(item => item.kind)).toEqual(['constValueNotConstant', 'constValueNotConstant', 'constValueNotConstant']);
		expect(found.map(item => source.slice(item.span.start, item.span.end))).toEqual(['B', 'D', 'E']);
	});
});
