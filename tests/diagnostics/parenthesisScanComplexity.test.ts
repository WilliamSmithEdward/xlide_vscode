import { describe, expect, it, vi } from 'vitest';
import { checkParentheses } from '../../src/analyzer/diagnostics/rules/parentheses';
import { parseModule } from '../../src/analyzer/parser/parseModule';
import { buildModuleSymbols } from '../../src/analyzer/symbols/buildModuleSymbols';
import { statementTokens } from '../../src/analyzer/diagnostics/walker';

const counted = vi.hoisted(() => ({ reads: 0 }));
vi.mock('../../src/analyzer/diagnostics/walker', async () => {
	const actual = await vi.importActual<typeof import('../../src/analyzer/diagnostics/walker')>('../../src/analyzer/diagnostics/walker');
	return { ...actual, statementTokens: vi.fn((...args: Parameters<typeof actual.statementTokens>) =>
		actual.statementTokens(...args).map(token => ({ ...token, get rawText() { counted.reads++; return token.rawText; } }))) };
});

describe('parenthesis scan complexity', () => {
	it('visits nested token ranges a bounded number of times', () => {
		const depth = 400;
		const source = 'Sub Main()\nDim value As Long\nvalue = ' + '('.repeat(depth) + '1' + ')'.repeat(depth) + '\nEnd Sub';
		const module = parseModule(source);
		const symbols = buildModuleSymbols('Module', 'standard', source, { parsedModule: module });
		counted.reads = 0;
		const push = vi.fn();
		checkParentheses(source, module, symbols, {}, undefined, push);
		expect(push).not.toHaveBeenCalled();
		expect(statementTokens).toHaveBeenCalled();
		expect(counted.reads).toBeLessThan(depth * 40);
	});
});
