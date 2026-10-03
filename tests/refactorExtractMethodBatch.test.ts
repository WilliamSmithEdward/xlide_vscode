import { beforeEach, describe, expect, it, vi } from 'vitest';
import { extractMethod } from '../src/analyzer/refactor/extractMethod';
import { applyVbaTextEdits } from '../src/analyzer/refactor/refactorTypes';
import { classifyReferenceKinds } from '../src/analyzer/references/referenceKinds';
import { findIdentifierOccurrences, findIdentifierOccurrencesForNames } from '../src/vbaSourceScan';

vi.mock('../src/analyzer/references/referenceKinds', async () => {
	const actual = await vi.importActual<typeof import('../src/analyzer/references/referenceKinds')>('../src/analyzer/references/referenceKinds');
	return { ...actual, classifyReferenceKinds: vi.fn(actual.classifyReferenceKinds) };
});
beforeEach(() => { vi.mocked(classifyReferenceKinds).mockClear(); });

describe('batched extract-method references', () => {
	it('classifies all selected locals in one token-stream pass', () => {
		const prefix = 'Option Explicit\nSub Main()\n' + Array.from({ length: 100 }, (_, i) => 'Dim v' + i + ' As Long\n').join('');
		const selection = 'v0 = v0 + 1\nv1 = v1 + 1';
		const source = prefix + selection + '\nDebug.Print v0, v1\nEnd Sub';
		const result = extractMethod({ source, span: { start: prefix.length, end: prefix.length + selection.length }, name: 'Work' });
		expect(result.ok).toBe(true);
		if (!result.ok) { throw new Error(result.reason); }
		expect(applyVbaTextEdits(source, result.edits)).toContain('Private Sub Work(ByRef v0 As Long, ByRef v1 As Long)');
		expect(classifyReferenceKinds).toHaveBeenCalledTimes(1);
	});

	it('matches whole-name Unicode occurrence semantics and ignores strings/comments', () => {
		const source = "Dim Café, ΔΕΛΤΑ As Long\r\nCafé = ΔΕΛΤΑ\nDebug.Print \"Café ΔΕΛΤΑ\" ' Café\robj.CAFÉ = [ΔΕΛΤΑ]\nCaféteria = 0\n' continued _\nCafé = ΔΕΛΤΑ\n";
		const found = findIdentifierOccurrencesForNames(source, ['CAFÉ', 'δελτα', 'missing', 'café']);
		expect([...found.keys()]).toEqual(['café', 'δελτα', 'missing']);
		expect(found.get('café')?.map(occ => occ.text)).toEqual(['Café', 'Café', 'CAFÉ']);
		expect(found.get('δελτα')?.map(occ => occ.text)).toEqual(['ΔΕΛΤΑ', 'ΔΕΛΤΑ', 'ΔΕΛΤΑ']);
		for (const name of ['CAFÉ', 'δελτα', 'missing']) {
			expect(found.get(name.toLowerCase())).toEqual(findIdentifierOccurrences(source, name));
		}
		expect(findIdentifierOccurrencesForNames(source, [])).toEqual(new Map());
	});

	it('keeps returned occurrence arrays independent across calls', () => {
		const first = findIdentifierOccurrencesForNames('x = x + 1', ['x']);
		first.get('x')!.pop();
		expect(findIdentifierOccurrencesForNames('x = x + 1', ['X']).get('x')).toHaveLength(2);
	});
});
