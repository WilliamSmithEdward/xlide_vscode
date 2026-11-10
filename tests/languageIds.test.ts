import { readFileSync, readdirSync } from 'fs';
import * as path from 'path';
import { describe, expect, it } from 'vitest';
import {
    VBA_CODE_FENCE,
    VBA_LANGUAGE_ID,
    XLIDE_VBA_LANGUAGE_ID,
    XML_LANGUAGE_ID,
} from '../src/analyzer/languageIds';

// Code that means one of the languages XLIDE contributes names it through
// src/analyzer/languageIds.ts, so a bare 'vba' in the source is always the
// VBA library (`VBA.Left`). Both are common, so without this a new language
// check written as a bare string reads exactly like the library name.
//
// The patterns below read the source as text, across lines. They cover the
// ways the source states a language today and the usual variations; a
// comment that quotes one of them is reported too, and is reworded.

const root = path.join(__dirname, '..');

interface Manifest {
    contributes: {
        languages: Array<{ id: string; extensions?: string[] }>;
        grammars: Array<{ language: string }>;
    };
}

const manifest = JSON.parse(readFileSync(path.join(root, 'package.json'), 'utf8')) as Manifest;

// The ids XLIDE contributes, quoted. Other languages, such as VS Code's own
// 'plaintext', are not XLIDE's to name.
const QUOTED = `['"\`](?:vba|xlide-vba|xml)['"\`]`;

// Every place the source states a language: a document's language id either
// way round, a document filter's or a scope's key, a selector given as a
// string, setting a document's language, a code block's language,
// language-scoped settings, and a Markdown code fence written plainly or
// with escaped backticks.
const LANGUAGE_POSITIONS: RegExp[] = [
    new RegExp(`\\blanguageId\\s*[!=]==?\\s*${QUOTED}`, 'g'),
    new RegExp(`${QUOTED}\\s*[!=]==?\\s*[\\w$.?!]*\\blanguageId\\b`, 'g'),
    new RegExp(`['"]?\\blanguage(?:Id)?['"]?\\s*:\\s*${QUOTED}`, 'g'),
    new RegExp(`\\b(?:register\\w+Provider|match|createLanguageStatusItem)\\s*\\(\\s*${QUOTED}`, 'g'),
    new RegExp(`\\b(?:setTextDocumentLanguage|appendCodeblock|getConfiguration)\\s*\\([^;]*?,\\s*${QUOTED}\\s*\\)`, 'g'),
    /(?:```|\\`\\`\\`)(?:vba|xlide-vba|xml)(?![\w-])/g,
];

function sourceFiles(dir: string): string[] {
    return readdirSync(dir, { withFileTypes: true }).flatMap((entry) => {
        const full = path.join(dir, entry.name);
        if (entry.isDirectory()) {
            return sourceFiles(full);
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

    it('recognise each way the source can state a language', () => {
        const bare = [
            "if (document.languageId === 'vba') {}",
            "if ('vba' !== editor.document.languageId) {}",
            "const selector = [{ scheme: 'file', 'language' : 'vba' }];",
            "vscode.languages.registerHoverProvider('vba', provider);",
            "vscode.workspace.getConfiguration('editor', 'xlide-vba');",
            "md.appendCodeblock(signatureOf(a, b),\n    'vba');",
            "await vscode.languages.setTextDocumentLanguage(doc, 'xml');",
            "const doc = '```vba\\nx\\n```';",
            "parts.push(`\\`\\`\\`vba\\n${x}`);",
        ];
        for (const source of bare) {
            expect(LANGUAGE_POSITIONS.some((pattern) => source.match(pattern) !== null), source).toBe(true);
        }
        const library = [
            "if (tokenText(toks[0]) === 'vba') {}",
            "const libraries = ['vba', 'office'];",
            "vscode.languages.createDiagnosticCollection('vba');",
            "vscode.workspace.getConfiguration('xlide', document.uri);",
            "const fence = '```vba-sample';",
        ];
        for (const source of library) {
            expect(LANGUAGE_POSITIONS.some((pattern) => source.match(pattern) !== null), source).toBe(false);
        }
    });

    it('are never written as bare strings where the source means a language', () => {
        const constants = path.join(root, 'src', 'analyzer', 'languageIds.ts');
        const found: string[] = [];
        for (const file of sourceFiles(path.join(root, 'src'))) {
            if (file === constants) {
                continue;
            }
            const source = readFileSync(file, 'utf8');
            for (const pattern of LANGUAGE_POSITIONS) {
                for (const match of source.matchAll(pattern)) {
                    const line = source.slice(0, match.index).split('\n').length;
                    found.push(`${path.relative(root, file)}:${line}: ${match[0].replace(/\s+/g, ' ')}`);
                }
            }
        }
        expect(found).toEqual([]);
    });
});
