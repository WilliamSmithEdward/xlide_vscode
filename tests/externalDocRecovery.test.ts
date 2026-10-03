import { describe, expect, it } from 'vitest';
import { parseMetadataFile } from '../src/analyzer/docs/externalDoc';

describe('external documentation member recovery', () => {
	it('returns no entries for a long suffix of unclosed members', () => {
		const xml = '<member name="A">text\n'.repeat(10000);
		expect(parseMetadataFile(xml)).toEqual([]);
	});

	it('consumes the first close even when an unclosed outer member encloses another opener', () => {
		const xml = '<member name="Outer">unfinished<member name="Inner"><summary>Inner text</summary></member><member name="Tail">tail</member>';
		expect(parseMetadataFile(xml)).toEqual([
			{name: 'Outer', doc: {summary: 'Inner text', params: [], source: 'external'}},
			{name: 'Tail', doc: {summary: 'tail', params: [], source: 'external'}},
		]);
	});

	it('consumes empty-name pairs without exposing nested members', () => {
		const xml = '<member name=" "><member name="Hidden">hidden</member></member><member name="Shown">shown</member>';
		expect(parseMetadataFile(xml).map(entry => entry.name)).toEqual(['Shown']);
	});

	it('preserves case-insensitive tags, whitespace, original coordinates and literal name entities', () => {
		const xml = 'İ<MEMBER\r\nname = " A&amp;B " ><summary>İ 中文</summary></mEmBeR>';
		expect(parseMetadataFile(xml)).toEqual([{name: 'A&amp;B', doc: {summary: 'İ 中文', params: [], source: 'external'}}]);
	});

	it('retains the accepted member opening syntax while recovering later pairs', () => {
		const xml = '<member name="Extra" other="x">bad</member><member other="x" name="Order">bad</member>'
			+ "<member name='Single'>bad</member><member name=\"Self\"/><member name=\"Good\">good</member>";
		expect(parseMetadataFile(xml).map(entry => entry.name)).toEqual(['Good']);
	});

	it('keeps already-parsed members when later content cannot close', () => {
		const xml = '<member name="Good">good</member>' + '<member name="Lost">unfinished'.repeat(1000);
		expect(parseMetadataFile(xml).map(entry => entry.name)).toEqual(['Good']);
	});
});
