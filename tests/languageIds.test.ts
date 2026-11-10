import { readFileSync, readdirSync } from 'fs';
import * as path from 'path';
import { describe, expect, it } from 'vitest';
import {
    VBA_CODE_FENCE,
    VBA_LANGUAGE_ID,
    XLIDE_VBA_LANGUAGE_ID,
    XML_LANGUAGE_ID,
} from '../src/analyzer/languageIds';

// Code that means a language names it through src/analyzer/languageIds.ts,
// so a bare 'vba' in the source is always the VBA library (`VBA.Left`). Both
// spellings are common, so without this a new language check written as a
// bare string reads exactly like the library name.

const root = path.join(__dirname, '..');

interface Manifest {
    contributes: {
        languages: Array<{ id: string; extensions?: string[] }>;
        grammars: Array<{ language: string }>;
    };
}

const manifest = JSON.parse(readFileSync(path.join(root, 'package.json'), 'utf8')) as Manifest;

// Every place the source states a language: a document's language id, a
// document selector, setting a document's language, a code block's language,
// language-scoped settings, and a Markdown code fence.
const LANGUAGE_POSITIONS: RegExp[] = [
    /languageId\s*[!=]==?\s*['"`]/,
    /\blanguage:\s*['"`]/,
    /\blanguageId:\s*['"`]/,
    /setTextDocumentLanguage\([^,]+,\s*['"`]/,
    /appendCodeblock\([^,]+,\s*['"`]/,
    /```(?:vba|xlide-vba|xml)\b/,
];

function sourceFiles(dir: string): string[] {
    return readdirSync(dir, { withFileTypes: true }).flatMap((entry) => {
        const full = path.join(dir, entry.name);
        if (entry.isDirectory()) {
            // The integration tests assert the ids a document really has.
            return entry.name === 'test' && dir === path.join(root, 'src') ? [] : sourceFiles(full);
        }
        return entry.name.endsWith('.ts') ? [full] : [];
    });
}

describe('language ids', () => {
    it('are the ones package.json contributes', () => {
        const ids = manifest.contributes.languages.map((language) => language.id);
        expect(ids).toEqual(expect.arrayContaining([VBA_LANGUAGE_ID, XLIDE_VBA_LANGUAGE_ID, XML_LANGUAGE_ID]));
        const vba = manifest.contributes.languages.find((language) => language.id === VBA_LANGUAGE_ID);
        expect(vba?.extensions).toContain('.bas');
        const xml = manifest.contributes.languages.find((language) => language.id === XML_LANGUAGE_ID);
        expect(xml?.extensions).toContain('.form');
        expect(manifest.contributes.grammars.map((grammar) => grammar.language))
            .toEqual(expect.arrayContaining([VBA_LANGUAGE_ID, XLIDE_VBA_LANGUAGE_ID]));
        expect(VBA_CODE_FENCE).toBe('```vba');
    });

    it('are never written as bare strings where the source means a language', () => {
        const constants = path.join(root, 'src', 'analyzer', 'languageIds.ts');
        const found: string[] = [];
        for (const file of sourceFiles(path.join(root, 'src'))) {
            if (file === constants) {
                continue;
            }
            readFileSync(file, 'utf8').split('\n').forEach((line, index) => {
                if (LANGUAGE_POSITIONS.some((pattern) => pattern.test(line))) {
                    found.push(`${path.relative(root, file)}:${index + 1}: ${line.trim()}`);
                }
            });
        }
        expect(found).toEqual([]);
    });
});
