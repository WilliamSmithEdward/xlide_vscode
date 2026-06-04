import { describe, it, expect } from 'vitest';
import { materializeKeywordSnippet, resolveKeywordCompletions } from '../src/analyzer';
import {
	detectSmartBlockOpener,
	VBA_BLOCK_INDENT_UNIT,
	VBA_SMART_BLOCK_SNIPPETS,
} from '../src/vbaStructuralAnalysis';

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
			'With ${1:object}\n\n\t.$0\n\nEnd With',
		);
	});

	it('preserves literal tab indentation when materializing nested snippet bodies', () => {
		const snippet = 'For ${1:i} = ${2:1} To ${3:10}\n\n\t$0\n\nNext ${1/(.*)/$1/}';

		expect(materializeKeywordSnippet(snippet, '    ')).toBe(
			'For ${1:i} = ${2:1} To ${3:10}\n\n    \t$0\n\n    Next ${1/(.*)/$1/}',
		);
	});

	it('uses one literal tab body unit across block archetype snippets', () => {
		const statementItems = resolveKeywordCompletions('Sub T()\n    wh\nEnd Sub\n', at('Sub T()\n    wh\nEnd Sub\n', '    wh')).items;
		expect(statementItems.find((item) => item.label === 'While')?.insertText).toBe(
			'While ${1:condition}\n\n\t$0\n\nWend',
		);

		const all = resolveKeywordCompletions('Sub T()\n    \nEnd Sub\n', at('Sub T()\n    \nEnd Sub\n', '    ')).items;
		for (const label of ['If', 'With', 'For', 'For Each', 'Do While', 'Do Until', 'While', 'Select Case', 'Sub', 'Function', 'Type', 'Enum']) {
			const item = all.find((candidate) => candidate.label === label);
			expect(item?.kind, label).toBe('snippet');
			expect(item?.insertText, label).toContain('\n\t');
			expect(item?.insertText, label).not.toContain('\n    ');
		}
	});

	it('projects compact block snippets from the same smart-block catalogue', () => {
		const result = resolveKeywordCompletions(
			'Sub T()\n    \nEnd Sub\n',
			at('Sub T()\n    \nEnd Sub\n', '    '),
			{ blockLayout: 'compact' },
		);

		expect(result.items.find((item) => item.label === 'With')?.insertText).toBe(
			'With ${1:object}\n\t.$0\nEnd With',
		);
		expect(result.items.find((item) => item.label === 'If Else')?.insertText).toBe(
			'If ${1:condition} Then\n\t$2\nElse\n\t$0\nEnd If',
		);
		expect(result.items.find((item) => item.label === 'Select Case')?.insertText).toBe(
			'Select Case ${1:expression}\n\tCase ${2:value}\n\t\t$0\nEnd Select',
		);
	});

	it('projects statement block snippets from the shared smart-block catalogue', () => {
		const result = resolveKeywordCompletions('Sub T()\n    \nEnd Sub\n', at('Sub T()\n    \nEnd Sub\n', '    '));
		const labels = result.items.map((item) => item.label);
		const expected = VBA_SMART_BLOCK_SNIPPETS
			.filter((spec) => spec.contexts.includes('statement'))
			.map((spec) => spec.label);

		for (const label of expected) {
			expect(labels, label).toContain(label);
		}
	});

	it('keeps shared block snippets aligned with Smart Enter openers', () => {
		for (const spec of VBA_SMART_BLOCK_SNIPPETS) {
			if (!spec.smartEnterExample) {
				continue;
			}
			const opener = detectSmartBlockOpener(spec.smartEnterExample);
			expect(opener?.endKeyword, spec.label).toBe(spec.smartEnterCloser);
			expect(spec.insertText, spec.label).toContain(`\n${VBA_BLOCK_INDENT_UNIT}`);
			expect(spec.insertText, spec.label).not.toContain('\n    ');
			if (opener?.bodyPrefix) {
				expect(spec.insertText.split('\n')[2], spec.label).toContain(
					VBA_BLOCK_INDENT_UNIT + opener.bodyPrefix,
				);
			}
		}
	});

	it('filters statement snippets by typed prefix', () => {
		const src = 'Sub T()\n    Wi\nEnd Sub\n';
		expect(labels(src, '    Wi')).toEqual(['With']);
	});

	it('matches compound keyword snippets when typed without spaces', () => {
		const src = 'Sub T()\n    forea\nEnd Sub\n';
		expect(labels(src, '    forea')).toEqual(['For Each']);
	});

	it('matches retired static snippet aliases through analyzer snippets', () => {
		expect(labels('Sub T()\n    ifelse\nEnd Sub\n', '    ifelse')).toEqual(['If Else']);
		expect(labels('Sub T()\n    dountil\nEnd Sub\n', '    dountil')).toEqual(['Do Loop Until']);
		expect(labels('Sub T()\n    propget\nEnd Sub\n', '    propget')).toEqual(['Property Get']);
		expect(labels('Sub T()\n    dp\nEnd Sub\n', '    dp')).toEqual(['Debug.Print']);
		expect(labels('Sub T()\n    onerror\nEnd Sub\n', '    onerror')).toEqual(['On Error GoTo Handler']);
	});

	it('offers common single-line statement snippets', () => {
		const src = 'Sub T()\n    \nEnd Sub\n';
		const result = resolveKeywordCompletions(src, at(src, '    '));
		const itemLabels = result.items.map((item) => item.label);

		for (const label of ['Dim', 'Set', 'ReDim', 'ReDim Preserve', 'Exit Sub', 'Exit Function', 'Exit For', 'Exit Do']) {
			expect(itemLabels, label).toContain(label);
		}
		expect(result.items.find((item) => item.label === 'Dim')?.insertText).toBe(
			'Dim ${1:name} As ${2:Variant}',
		);
		expect(result.items.find((item) => item.label === 'Set')?.insertText).toBe(
			'Set ${1:object} = ${2:value}',
		);
	});

	it('matches common compact single-line statement aliases', () => {
		expect(labels('Sub T()\n    redimpreserve\nEnd Sub\n', 'redimpreserve')).toEqual(['ReDim Preserve']);
		expect(labels('Sub T()\n    exitf\nEnd Sub\n', 'exitf')).toEqual([
			'Exit Function',
			'Exit For',
		]);
	});

	it('owns former declaration snippets in the analyzer', () => {
		expect(labels('Ty', 'Ty')).toEqual(['Type']);
		expect(labels('En', 'En')).toEqual(['Enum']);
	});

	it('offers the active block closer first on a blank line', () => {
		const src = 'Sub T()\n    With rng\n        \n';
		const result = resolveKeywordCompletions(src, at(src, '        '));
		expect(result.items[0]?.label).toBe('End With');
		expect(result.items[0]?.sortText).toBe('000:close');
	});

	it('keeps full block shortcut snippets for explicit completion gestures', () => {
		const src = 'Sub T()\n    For\nEnd Sub\n';
		const result = resolveKeywordCompletions(src, at(src, '    For'));
		expect(result.items.find((item) => item.label === 'For')?.insertText).toBe(
			'For ${1:i} = ${2:1} To ${3:10}\n\n\t$0\n\nNext ${1/(.*)/$1/}',
		);
		expect(result.items.find((item) => item.label === 'For Each')?.insertText).toBe(
			'For Each ${1:item} In ${2:collection}\n\n\t$0\n\nNext ${1/(.*)/$1/}',
		);
	});

	it('suggests the active loop closer with the iterator from the opener', () => {
		const src = 'Sub T()\n    For Each cell In Selection\n        \n';
		const result = resolveKeywordCompletions(src, at(src, '        '));
		expect(result.items[0]?.label).toBe('Next cell');
		expect(result.items[0]?.insertText).toBe('Next cell');
		expect(result.items[0]?.sortText).toBe('000:close');
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
		expect(result.items.map((item) => item.label)).toContain('Declare PtrSafe Function');
		expect(result.items.find((item) => item.label === 'Declare PtrSafe Function')?.insertText)
			.toBe('Declare PtrSafe Function ${1:ProcedureName} Lib "${2:library}" (ByVal ${3:argument} As ${4:LongPtr}) As ${5:LongPtr}');
	});

	it('offers pointer-safe Declare aliases after access modifiers', () => {
		const src = 'Public ptrsafe';
		const result = resolveKeywordCompletions(src, at(src, 'Public ptrsafe'));
		expect(result.exclusive).toBe(true);
		expect(result.items.map((item) => item.label)).toEqual([
			'Declare PtrSafe Sub',
			'Declare PtrSafe Function',
		]);
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

	it('offers Exit targets after Exit', () => {
		const src = 'Sub T()\n    Exit ';
		const result = resolveKeywordCompletions(src, at(src, 'Exit '));
		expect(result.exclusive).toBe(true);
		expect(result.items.map((item) => item.label)).toEqual([
			'Sub',
			'Function',
			'Property',
			'For',
			'Do',
		]);
		expect(labels('Sub T()\n    Exit F\nEnd Sub\n', 'Exit F')).toEqual(['Function', 'For']);
		expect(resolveKeywordCompletions('Sub T()\n    Exit For ', at('Sub T()\n    Exit For ', 'Exit For ')).items)
			.toEqual([]);
	});

	it('offers Do and Loop condition forms in their keyword slots', () => {
		const doResult = resolveKeywordCompletions('Sub T()\n    Do ', at('Sub T()\n    Do ', 'Do '));
		expect(doResult.exclusive).toBe(true);
		expect(doResult.items.map((item) => item.label)).toEqual(['While', 'Until']);
		expect(doResult.items.find((item) => item.label === 'While')?.insertText).toBe(
			'While ${1:condition}',
		);

		const loopResult = resolveKeywordCompletions('Sub T()\n    Loop ', at('Sub T()\n    Loop ', 'Loop '));
		expect(loopResult.exclusive).toBe(true);
		expect(loopResult.items.map((item) => item.label)).toEqual(['While', 'Until']);
	});

	it('offers Select Case and For Each In continuations', () => {
		const selectResult = resolveKeywordCompletions('Sub T()\n    Select ', at('Sub T()\n    Select ', 'Select '));
		expect(selectResult.exclusive).toBe(true);
		expect(selectResult.items.map((item) => item.label)).toEqual(['Case']);

		const forEachResult = resolveKeywordCompletions(
			'Sub T()\n    For Each cell ',
			at('Sub T()\n    For Each cell ', 'For Each cell '),
		);
		expect(forEachResult.exclusive).toBe(true);
		expect(forEachResult.items.map((item) => item.label)).toEqual(['In']);
	});

	it('offers Case helpers without hiding expression completions', () => {
		const src = 'Sub T()\n    Select Case value\n        Case ';
		const result = resolveKeywordCompletions(src, at(src, '        Case '));
		expect(result.exclusive).toBe(false);
		expect(result.items.map((item) => item.label)).toEqual(['Else', 'Is']);
		expect(result.items.find((item) => item.label === 'Is')?.insertText).toBe(
			'Is ${1:operator} ${2:value}',
		);
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
		expect(result.items.map((item) => item.label)).toContain('#If VBA7 Declare Function');
		expect(result.items.map((item) => item.label)).toContain('#If Win64 Then');
		expect(result.items.find((item) => item.label === '#If')?.insertText).toBe(
			'#If ${1:condition} Then\n\n\t$0\n\n#End If',
		);
		expect(result.items.find((item) => item.label === '#If VBA7 Declare Sub')?.insertText).toBe(
			[
				'#If VBA7 Then',
				'Public Declare PtrSafe Sub ${1:ProcedureName} Lib "${2:library}" (ByVal ${3:argument} As ${4:LongPtr})',
				'#Else',
				'Public Declare Sub ${1/(.*)/$1/} Lib "${2/(.*)/$1/}" (ByVal ${3/(.*)/$1/} As ${5:Long})',
				'#End If',
			].join('\n'),
		);
	});

	it('matches conditional Declare snippets through compact aliases', () => {
		const src = '#vba7declare';
		const result = resolveKeywordCompletions(src, at(src, '#vba7declare'));
		expect(result.items.map((item) => item.label)).toEqual(['#If VBA7 Declare Sub']);
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
