import { beforeEach, describe, expect, it, vi } from 'vitest';
import { resolveCanonicalCaseEdits } from '../src/analyzer/completion/canonicalCasing';
import { createIdentifierCompletionResolver, resolveIdentifierCompletions } from '../src/analyzer/completion/identifierCompletion';
import { buildModuleSymbols } from '../src/analyzer/symbols/buildModuleSymbols';

vi.mock('../src/analyzer/symbols/buildModuleSymbols', async () => {
	const actual = await vi.importActual<typeof import('../src/analyzer/symbols/buildModuleSymbols')>('../src/analyzer/symbols/buildModuleSymbols');
	return { ...actual, buildModuleSymbols: vi.fn(actual.buildModuleSymbols) };
});
beforeEach(() => { vi.mocked(buildModuleSymbols).mockClear(); });

describe('request-local identifier symbols', () => {
	it('builds symbols once for a bulk casing request', () => {
		const source = 'Sub Main()\nDim Value As Long\n' + 'value = value + 1\n'.repeat(30) + 'End Sub';
		const edits = resolveCanonicalCaseEdits(source, { start: 0, end: source.length },
			{ identifier: { includeGlobals: false, includeRuntime: false } });
		expect(edits).toHaveLength(60);
		expect(buildModuleSymbols).toHaveBeenCalledTimes(1);
	});

	it('selects the scope at each offset and returns independent completion records', () => {
		const source = 'Sub First()\nDim LocalOne As Long\nlocal\nEnd Sub\nSub Second()\nDim LocalTwo As Long\nlocal\nEnd Sub';
		const ctx = { includeGlobals: false, includeRuntime: false };
		const at = createIdentifierCompletionResolver(source, ctx);
		const firstOffset = source.indexOf('local') + 5;
		const secondOffset = source.lastIndexOf('local') + 5;
		const first = at(firstOffset);
		expect(first.map(item => item.name)).toEqual(['LocalOne']);
		expect(at(secondOffset).map(item => item.name)).toEqual(['LocalTwo']);
		expect(buildModuleSymbols).toHaveBeenCalledTimes(1);
		first[0].name = 'Changed';
		expect(at(firstOffset)).toEqual(resolveIdentifierCompletions(source, firstOffset, ctx));
	});

	it('does not build symbols for declaration names and isolates source revisions', () => {
		const source = 'Sub Main()\nDim OldName As Long\nold\nEnd Sub';
		const at = createIdentifierCompletionResolver(source, { includeGlobals: false, includeRuntime: false });
		expect(at(source.indexOf('OldName') + 7)).toEqual([]);
		expect(buildModuleSymbols).not.toHaveBeenCalled();
		expect(at(source.indexOf('old') + 3).map(item => item.name)).toEqual(['OldName']);
		const changed = source.replace('OldName', 'NewName').replace('old', 'new');
		const next = createIdentifierCompletionResolver(changed, { includeGlobals: false, includeRuntime: false });
		expect(next(changed.indexOf('new') + 3).map(item => item.name)).toEqual(['NewName']);
		expect(buildModuleSymbols).toHaveBeenCalledTimes(2);
	});
});
