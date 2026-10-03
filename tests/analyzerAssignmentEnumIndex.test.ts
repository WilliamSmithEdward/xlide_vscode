import { describe, expect, it } from 'vitest';
import { checkAssignmentTypes } from '../src/analyzer/diagnostics/rules/assignments';
import { parseModule } from '../src/analyzer/parser/parseModule';
import { buildModuleSymbols } from '../src/analyzer/symbols/buildModuleSymbols';
import type { VbaSymbol } from '../src/analyzer/symbols/symbolModel';

function messages(source: string, projectSymbols?: readonly VbaSymbol[]): string[] {
	const module = parseModule(source);
	const symbols = buildModuleSymbols('M', 'standard', source, { parsedModule: module });
	const found: string[] = [];
	checkAssignmentTypes(source, module, symbols, projectSymbols, {}, undefined, (_rule, message) => found.push(message));
	return found;
}

describe('assignment enum name index', () => {
	it('does not reread every unrelated symbol kind for every assignment', () => {
		const members = 100, assignments = 200;
		const source = 'Option Explicit\n' + Array.from({ length: members }, (_, i) => 'Private v' + i + ' As Long\n').join('')
			+ 'Sub Main(ByVal x As Long)\nDim total As Long\n'
			+ Array.from({ length: assignments }, () => 'total = x + 1\n').join('') + 'End Sub';
		const module = parseModule(source);
		const base = buildModuleSymbols('M', 'standard', source, { parsedModule: module });
		let reads = 0;
		const symbols = { ...base, root: { ...base.root, children: base.root.children?.map(symbol => {
			const copy = { ...symbol };
			Object.defineProperty(copy, 'kind', { get() { if (symbol.kind === 'moduleVariable') { reads++; } return symbol.kind; } });
			return copy;
		}) } };
		const found: string[] = [];
		checkAssignmentTypes(source, module, symbols, undefined, {}, undefined, (_rule, message) => found.push(message));
		expect(found).toEqual([]);
		expect(reads).toBeLessThan(members * 30);
	});

	it('keeps case-insensitive module enums as Long assignment targets', () => {
		const source = 'Private Enum Status\nReady\nEnd Enum\nSub Main()\nDim value As sTaTuS\nvalue = "abc"\nEnd Sub';
		expect(messages(source)).toEqual([expect.stringContaining('expects Long')]);
	});

	it('recognizes qualified project enums and rebuilds for replaced visible symbols', () => {
		const enumSource = 'Public Enum Status\nReady\nEnd Enum';
		const external = buildModuleSymbols('Definitions', 'standard', enumSource).root.children!;
		const source = 'Sub Main()\nDim value As Definitions.Status\nvalue = "abc"\nEnd Sub';
		for (const visible of [external, [], external]) {
			const found = messages(source, visible);
			expect(found.some(message => message.includes('expects Long'))).toBe(visible.length > 0);
		}
	});
});
