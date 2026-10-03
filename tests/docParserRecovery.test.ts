import { describe, expect, it } from 'vitest';
import { parseDocBody } from '../src/analyzer/docs/docComment';
import { parseMetadataFile } from '../src/analyzer/docs/externalDoc';

describe('documentation parser recovery', () => {
	it('finds a later self-closing parameter after unclosed pairs', () => {
		const body = '<param name="Lost">unfinished\n'.repeat(2000) + '<param name="Found" type="Long"/>';
		expect(parseDocBody(body, 'inline').params).toEqual([{name: 'Found', type: 'Long', text: ''}]);
	});

	it('keeps a paired parameter body non-overlapping even when it contains another opening', () => {
		const body = '<param name="Outer">first <param name="Inner">second</param> tail</param><param name="After"/>';
		expect(parseDocBody(body, 'inline').params).toEqual([
			{name: 'Outer', text: 'first <param name="Inner">second'},
			{name: 'After', text: ''},
		]);
	});

	it('retains the first paired summary rather than treating a reopening as a boundary', () => {
		expect(parseDocBody('<summary>first <summary>second</summary>', 'inline').summary)
			.toBe('first <summary>second');
	});

	it('keeps lenient opening-tag matching and last-attribute-wins behavior', () => {
		const body = '<parametric name="No"/><param-name="Odd">x</param><param name="First" name="Last" type="Long"//><returns type="Currency"//>';
		const doc = parseDocBody(body, 'inline');
		expect(doc.params).toEqual([{name: 'Odd', text: 'x'}, {name: 'Last', type: 'Long', text: ''}]);
		expect(doc.returns).toBe('');
		expect(doc.returnsType).toBe('Currency');
	});

	it('preserves malformed nested opening text in attributes', () => {
		const doc = parseDocBody('<param broken <param name="A"/>', 'inline');
		expect(doc.params).toEqual([{name: 'A', text: ''}]);
	});

	it('uses original string coordinates when casing Unicode text changes its length', () => {
		const doc = parseDocBody('İ<SuMmArY>İ &amp; 中文</SUMMARY><PARAM name="ß">İ</param>', 'external');
		expect(doc).toMatchObject({summary: 'İ & 中文', params: [{name: 'ß', text: 'İ'}], source: 'external'});
	});

	it('stops on repeated incomplete opening prefixes without inventing content', () => {
		const body = '<param name="A" '.repeat(10000);
		expect(parseDocBody(body, 'inline')).toEqual({params: [], source: 'inline'});
	});

	it('recovers a later summary when an earlier opening has no close', () => {
		const doc = parseDocBody('<summary>unfinished <summary/>', 'inline');
		expect(doc.summary).toBe('');
	});

	it('preserves example layout and parses the same recovery inside external metadata', () => {
		const xml = '<member name="M.P"><param name="Lost">unfinished <param name="Found"/><example>\r\nDim x As Long\r\nx = 1\r\n</example></member>';
		const entries = parseMetadataFile(xml);
		expect(entries).toHaveLength(1);
		expect(entries[0].doc.params).toEqual([{name: 'Found', text: ''}]);
		expect(entries[0].doc.example).toBe('Dim x As Long\nx = 1');
	});
});
