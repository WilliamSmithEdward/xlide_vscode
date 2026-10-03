import { beforeEach, describe, expect, it, vi } from 'vitest';
import { checkDeclarationOrder } from '../../src/analyzer/diagnostics/rules/declarationOrder';
import { parseModule } from '../../src/analyzer/parser/parseModule';
import { tokenize } from '../../src/analyzer/lexer/tokenize';

vi.mock('../../src/analyzer/lexer/tokenize', async () => {
	const actual = await vi.importActual<typeof import('../../src/analyzer/lexer/tokenize')>('../../src/analyzer/lexer/tokenize');
	return { ...actual, tokenize: vi.fn(actual.tokenize) };
});
beforeEach(() => { vi.mocked(tokenize).mockClear(); });

describe('prepared dependency walks', () => {
	it('lexes shared dependency expressions once per scope in a rule pass', () => {
		const source = Array.from({ length: 100 }, (_, i) => 'Public Const A' + i + ' = Other.B0\n').join('');
		const module = parseModule(source);
		const external = new Map<string, string>();
		for (let i = 0; i < 100; i++) { external.set('other.b' + i, i === 99 ? '1' : 'B' + (i + 1)); }
		vi.mocked(tokenize).mockClear();
		const push = vi.fn();
		checkDeclarationOrder(source, module, 'Module', external, undefined, undefined, push);
		expect(push).not.toHaveBeenCalled();
		expect(vi.mocked(tokenize).mock.calls.length).toBeLessThan(300);
	});

	it('handles deep cross-module cycles without consuming the call stack', () => {
		const source = 'Public Const Root = Other.B0\n';
		const module = parseModule(source);
		const external = new Map<string, string>();
		for (let i = 0; i < 10000; i++) { external.set('other.b' + i, i === 9999 ? 'Module.Root' : 'B' + (i + 1)); }
		external.set('module.root', 'Other.B0');
		const push = vi.fn();
		checkDeclarationOrder(source, module, 'Module', external, undefined, undefined, push);
		expect(push).toHaveBeenCalledTimes(1);
		expect(push.mock.calls[0][0]).toBe('circularDeclarationDependency');
		expect(push.mock.calls[0][1]).toContain("through 'Other.B0'");
	});

	it('keeps ordered foreign provenance and isolates changed project tables', () => {
		const source = 'Public Const Root = Other.First + Other.Second\n';
		const module = parseModule(source);
		const external = new Map([['other.first', 'Other.Second'], ['other.second', 'Module.Root'], ['module.root', 'Other.First + Other.Second']]);
		const first = vi.fn();
		checkDeclarationOrder(source, module, 'Module', external, undefined, undefined, first);
		expect(first.mock.calls[0][1]).toContain("through 'Other.First'");
		external.set('other.first', '1');
		const second = vi.fn();
		checkDeclarationOrder(source, module, 'Module', external, undefined, undefined, second);
		expect(second.mock.calls[0][1]).toContain("through 'Other.Second'");
	});
});
