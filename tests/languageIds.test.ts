// The language ids in src/languageIds.ts are the ones package.json
// contributes: a constant that drifts from the manifest would match no
// document, and nothing else would say so.

import { readFileSync } from 'fs';
import { describe, expect, it } from 'vitest';
import {
	FORM_MARKUP_LANGUAGE_ID,
	VBA_CODE_FENCE,
	VBA_LANGUAGE_ID,
	XLIDE_VBA_LANGUAGE_ID,
} from '../src/languageIds';

interface Manifest {
	contributes: {
		languages: Array<{ id: string; extensions?: string[] }>;
		grammars: Array<{ language: string }>;
	};
}

describe('language ids', () => {
	const { languages, grammars } = (JSON.parse(readFileSync('package.json', 'utf8')) as Manifest).contributes;
	const extensionsOf = (id: string) => languages.find((language) => language.id === id)?.extensions ?? [];

	it('name every language package.json contributes, and no other', () => {
		expect(languages.map((language) => language.id).sort())
			.toEqual([FORM_MARKUP_LANGUAGE_ID, VBA_LANGUAGE_ID, XLIDE_VBA_LANGUAGE_ID].sort());
	});

	it('belong to the files they are named for', () => {
		expect(extensionsOf(VBA_LANGUAGE_ID)).toEqual(expect.arrayContaining(['.bas', '.cls', '.frm']));
		expect(extensionsOf(FORM_MARKUP_LANGUAGE_ID)).toEqual(['.form']);
		// A module inside an Office file has no extension of its own.
		expect(extensionsOf(XLIDE_VBA_LANGUAGE_ID)).toEqual([]);
	});

	it('have the VBA grammar for both kinds of VBA document', () => {
		expect(grammars.map((grammar) => grammar.language).sort())
			.toEqual([VBA_LANGUAGE_ID, XLIDE_VBA_LANGUAGE_ID].sort());
	});

	it('open a Markdown code block with the VBA language', () => {
		expect(VBA_CODE_FENCE).toBe('```vba');
	});
});
