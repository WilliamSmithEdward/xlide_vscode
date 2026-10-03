import { describe, expect, it } from 'vitest';
import { leadingDocLines, parseDocBody, scanDocTags } from '../src/analyzer/docs/docComment';

function tagsOf(source: string) {
	return scanDocTags(leadingDocLines(source, source.indexOf('Sub P')))!;
}

describe('documentation scanner recovery', () => {
	it('skips a long unknown attribute word and preserves earlier valid attributes', () => {
		const body = '<param name="Good" type="Long" ' + 'x'.repeat(20000) + ' />';
		expect(parseDocBody(body, 'inline').params).toEqual([{name: 'Good', type: 'Long', text: ''}]);
		const tags = tagsOf("''' " + body + '\nSub P()\nEnd Sub');
		expect(tags[0]).toMatchObject({name: 'Good', hasHints: true, text: ''});
	});

	it('retains matching after invalid leading characters and failed unquoted values', () => {
		const body = '<param 12name="Digit" other= name="Nested" NAME = "Last" type/="No" unit="N" />';
		expect(parseDocBody(body, 'inline').params).toEqual([{name: 'Last', unit: 'N', text: ''}]);
		const source = "''' " + body + '\nSub P()\nEnd Sub';
		const tag = tagsOf(source)[0];
		expect(tag.name).toBe('Last');
		expect(source.slice(tag.nameSpan!.start, tag.nameSpan!.end)).toBe('Last');
	});

	it.each(['\n', '\r\n'])('preserves raw name spans while decoding entity values (%j)', (eol) => {
		const source = ["''' <param", "''' name=\"A&amp;B\" type=\"Long\" />", 'Sub P()', 'End Sub'].join(eol);
		const tag = tagsOf(source)[0];
		expect(tag.name).toBe('A&B');
		expect(source.slice(tag.nameSpan!.start, tag.nameSpan!.end)).toBe('A&amp;B');
		expect(tag.open).toEqual({start: source.indexOf('<param'), end: source.indexOf('/>') + 2});
	});

	it('keeps valid tags before a long incomplete opening suffix', () => {
		const source = "''' <summary/>\n" + "''' <param\n".repeat(10000) + 'Sub P()\nEnd Sub';
		const tags = tagsOf(source);
		expect(tags).toHaveLength(1);
		expect(tags[0]).toMatchObject({tag: 'summary', text: ''});
	});

	it('keeps the earliest incomplete opener when a later nested opener supplies its ending', () => {
		const source = ["''' <param", "''' <param name=\"Found\"/>", 'Sub P()', 'End Sub'].join('\n');
		const tags = tagsOf(source);
		expect(tags).toHaveLength(1);
		expect(tags[0]).toMatchObject({name: 'Found', open: {start: source.indexOf('<param'), end: source.indexOf('/>') + 2}});
	});

	it('retains parsed hints before an unterminated quoted attribute', () => {
		const body = '<returns type="Long" name="unfinished>';
		expect(parseDocBody(body + '</returns>', 'inline')).toMatchObject({returns: '', returnsType: 'Long'});
		const tag = tagsOf("''' " + body + '</returns>\nSub P()\nEnd Sub')[0];
		expect(tag.name).toBeUndefined();
		expect(tag.hasHints).toBe(true);
	});
});
