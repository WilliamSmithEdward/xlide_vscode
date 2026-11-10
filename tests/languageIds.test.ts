import { readFileSync, readdirSync } from 'fs';
import * as path from 'path';
import { describe, expect, it } from 'vitest';
import {
    VBA_CODE_FENCE,
    VBA_LANGUAGE_ID,
    XLIDE_VBA_LANGUAGE_ID,
    XML_LANGUAGE_ID,
} from '../src/analyzer/languageIds';

// Code that names one of XLIDE's languages uses src/analyzer/languageIds.ts,
// so a bare 'vba' in the source is the VBA library (`VBA.Left`), or the
// diagnostic collection's name. This catches the usual slip of writing the id
// out by hand, line by line; it is a tripwire, not a parser.

const root = path.join(__dirname, '..');
const ID = `['"\`](?:vba|xlide-vba|xml)['"\`]`;
const BARE_LANGUAGE = [
    new RegExp(`\\blanguageId\\s*(?:[!=]==?|,)\\s*${ID}`),
    new RegExp(`${ID}\\s*[!=]==?\\s*\\S*\\blanguageId\\b`),
    new RegExp(`\\blanguage(?:Id)?\\s*:\\s*${ID}`),
    new RegExp(`\\b(?:setTextDocumentLanguage|appendCodeblock)\\(.*,\\s*${ID}\\s*\\)`),
    new RegExp(`\\bregister\\w+Provider\\(\\s*${ID}`),
    /(?:```|\\`\\`\\`)(?:vba|xlide-vba|xml)(?![\w-])/,
];
const isBare = (line: string) => BARE_LANGUAGE.some((pattern) => pattern.test(line));

function sourceFiles(dir: string): string[] {
    return readdirSync(dir, { withFileTypes: true }).flatMap((entry) => {
        const full = path.join(dir, entry.name);
        return entry.isDirectory() ? sourceFiles(full) : entry.name.endsWith('.ts') ? [full] : [];
    });
}

describe('language ids', () => {
    it('are the ones package.json contributes', () => {
        const { languages, grammars } = JSON.parse(readFileSync(path.join(root, 'package.json'), 'utf8')).contributes;
        const byId = new Map<string, { extensions?: string[] }>(languages.map((language: { id: string }) => [language.id, language]));
        expect(byId.get(VBA_LANGUAGE_ID)?.extensions).toContain('.bas');
        expect(byId.has(XLIDE_VBA_LANGUAGE_ID)).toBe(true);
        expect(byId.get(XML_LANGUAGE_ID)?.extensions).toContain('.form');
        expect(grammars.map((grammar: { language: string }) => grammar.language))
            .toEqual(expect.arrayContaining([VBA_LANGUAGE_ID, XLIDE_VBA_LANGUAGE_ID]));
        expect(VBA_CODE_FENCE).toBe('```vba');
    });

    it('catch an id written out by hand, and leave the library alone', () => {
        expect(isBare("if (document.languageId === 'vba') {")).toBe(true);
        expect(isBare("assert.equal(renamed?.languageId, 'xml');")).toBe(true);
        expect(isBare("{ scheme: XLIDE_SCHEME, language: 'vba' },")).toBe(true);
        expect(isBare("await vscode.languages.setTextDocumentLanguage(doc, 'xml');")).toBe(true);
        expect(isBare("md.appendCodeblock(info.signature, 'vba');")).toBe(true);
        expect(isBare("'**Boolean literal**\\n\\n```vba\\nTrue\\n```',")).toBe(true);
        expect(isBare('parts.push(`\\`\\`\\`vba\\n${doc.example}\\n\\`\\`\\``);')).toBe(true);
        expect(isBare("if (tokenText(toks[0]) === 'vba' && toks[1]?.rawText === '.') {")).toBe(false);
        expect(isBare("const collection = vscode.languages.createDiagnosticCollection('vba');")).toBe(false);
        expect(isBare("await vscode.languages.setTextDocumentLanguage(document, 'plaintext');")).toBe(false);
        expect(isBare("'```json',")).toBe(false);
    });

    it('are not written out by hand anywhere under src/', () => {
        const constants = path.join(root, 'src', 'analyzer', 'languageIds.ts');
        const found = sourceFiles(path.join(root, 'src'))
            .filter((file) => file !== constants)
            .flatMap((file) => readFileSync(file, 'utf8').split('\n')
                .map((line, index) => ({ line, at: `${path.relative(root, file)}:${index + 1}` }))
                .filter(({ line }) => isBare(line))
                .map(({ line, at }) => `${at}: ${line.trim()}`));
        expect(found).toEqual([]);
    });
});
