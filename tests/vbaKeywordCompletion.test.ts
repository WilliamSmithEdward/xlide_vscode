import { describe, it, expect } from 'vitest';
import { resolveKeywordCompletions } from '../src/analyzer';

function at(src: string, marker: string): number {
	const idx = src.indexOf(marker);
	if (idx < 0) {
		throw new Error(`marker not found: ${marker}`);
	}
	return idx + marker.length;
}

function labels(src: string, marker: string) {
	return resolveKeywordCompletions(src, at(src, marker)).items.map((item) => item.label);
}

describe('keyword completion - statement snippets', () => {
	it('offers block snippets at statement start without being exclusive', () => {
		const src = 'Sub T()\n    \nEnd Sub\n';
		const result = resolveKeywordCompletions(src, at(src, '    '));
		expect(result.exclusive).toBe(false);
		expect(result.items.map((item) => item.label)).toContain('If');
		expect(result.items.map((item) => item.label)).toContain('With');
		expect(result.items.find((item) => item.label === 'With')?.insertText).toBe(
			'With ${1:object}\n    $0\nEnd With',
		);
	});

	it('filters statement snippets by typed prefix', () => {
		const src = 'Sub T()\n    Wi\nEnd Sub\n';
		expect(labels(src, '    Wi')).toEqual(['With']);
	});

	it('matches compound keyword snippets when typed without spaces', () => {
		const src = 'Sub T()\n    forea\nEnd Sub\n';
		expect(labels(src, '    forea')).toEqual(['For Each']);
	});

	it('offers the active block closer first on a blank line', () => {
		const src = 'Sub T()\n    With rng\n        \n';
		const result = resolveKeywordCompletions(src, at(src, '        '));
		expect(result.items[0]?.label).toBe('End With');
		expect(result.items[0]?.sortText).toBe('000:close');
	});

	it('mirrors loop iterator placeholders in the generated Next statement', () => {
		const src = 'Sub T()\n    For\nEnd Sub\n';
		const result = resolveKeywordCompletions(src, at(src, '    For'));
		expect(result.items.find((item) => item.label === 'For')?.insertText).toBe(
			'For ${1:i} = ${2:1} To ${3:10}\n    $0\nNext ${1:i}',
		);
		expect(result.items.find((item) => item.label === 'For Each')?.insertText).toBe(
			'For Each ${1:item} In ${2:collection}\n    $0\nNext ${1:item}',
		);
	});
});

describe('keyword completion - narrow grammar contexts', () => {
	it('offers only Option clauses after Option', () => {
		const src = 'Option ';
		const result = resolveKeywordCompletions(src, at(src, 'Option '));
		expect(result.exclusive).toBe(true);
		expect(result.items.map((item) => item.label)).toEqual([
			'Explicit',
			'Base 0',
			'Base 1',
			'Compare Binary',
			'Compare Text',
		]);
	});

	it('offers Option Compare values after Option Compare', () => {
		const src = 'Option Compare ';
		const result = resolveKeywordCompletions(src, at(src, 'Option Compare '));
		expect(result.exclusive).toBe(true);
		expect(result.items.map((item) => item.label)).toEqual(['Binary', 'Text']);
	});

	it('offers declaration snippets after access modifiers', () => {
		const src = 'Private ';
		const result = resolveKeywordCompletions(src, at(src, 'Private '));
		expect(result.exclusive).toBe(true);
		expect(result.items.map((item) => item.label)).toContain('Sub');
		expect(result.items.find((item) => item.label === 'Property Get')?.insertText).toContain(
			'End Property',
		);
	});

	it('offers On Error continuations', () => {
		const src = 'Sub T()\n    On Error ';
		const result = resolveKeywordCompletions(src, at(src, 'On Error '));
		expect(result.exclusive).toBe(true);
		expect(result.items.map((item) => item.label)).toEqual([
			'GoTo 0',
			'GoTo -1',
			'GoTo label',
			'Resume Next',
		]);
	});

	it('suggests the matching End form after End', () => {
		const src = 'Sub T()\n    If ok Then\n        End ';
		const result = resolveKeywordCompletions(src, at(src, 'End '));
		expect(result.exclusive).toBe(true);
		expect(result.items.map((item) => item.label)).toEqual(['If']);
		expect(result.items[0]?.insertText).toBe('If');
	});

	it('suppresses End completions when the innermost open block closes another way', () => {
		const src = 'Sub T()\n    For i = 1 To 3\n        End ';
		const result = resolveKeywordCompletions(src, at(src, 'End '));
		expect(result.exclusive).toBe(true);
		expect(result.items).toEqual([]);
	});

	it('offers directive snippets after #', () => {
		const src = '#';
		const result = resolveKeywordCompletions(src, at(src, '#'));
		expect(result.exclusive).toBe(true);
		expect(result.items.map((item) => item.label)).toContain('#If');
	});
});

describe('keyword completion - suppressed positions', () => {
	it('does not offer snippets inside comments or strings', () => {
		expect(resolveKeywordCompletions("' If", 4).items).toEqual([]);
		expect(resolveKeywordCompletions('"If', 3).items).toEqual([]);
	});

	it('does not offer statement snippets inside an expression', () => {
		const src = 'Sub T()\n    x = ';
		expect(resolveKeywordCompletions(src, at(src, 'x = ')).items).toEqual([]);
	});
});
