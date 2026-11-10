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
// src/analyzer/languageIds.ts, so a bare 'vba' in the source is the VBA
// library (`VBA.Left`), or the diagnostic collection's name. Both meanings are
// common, so without this a new language check written as a bare string reads
// exactly like the library name.
//
// The scan reads the source as text, across lines, with comments and the
// bodies of regular expressions blanked out first.

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
const QUOTED_ONLY = new RegExp(`^${QUOTED}$`);

// A document's language id compared or asserted either way round, a document
// filter's or a scope's key, and a Markdown code fence written plainly or
// with escaped backticks.
const LANGUAGE_PATTERNS: RegExp[] = [
    new RegExp(`\\blanguageId\\s*(?:[!=]==?|\\)?\\s*,)\\s*${QUOTED}`, 'g'),
    new RegExp(`${QUOTED}\\s*(?:[!=]==?|,)\\s*[\\w$.?!]*\\blanguageId\\b`, 'g'),
    new RegExp(`['"]?\\blanguage(?:Id)?['"]?\\s*:\\s*${QUOTED}`, 'g'),
    /(?:```|\\`\\`\\`)(?:vba|xlide-vba|xml)(?![\w-])/g,
];

// Calls whose argument at the given index is a language, or a selector that
// can hold one (a string, an array, a filter object).
const LANGUAGE_ARGUMENTS: Array<{ name: RegExp; index: number; selector: boolean }> = [
    { name: /\bsetTextDocumentLanguage$/, index: 1, selector: false },
    { name: /\bappendCodeblock$/, index: 1, selector: false },
    { name: /\bgetConfiguration$/, index: 1, selector: false },
    { name: /\bregister\w+Provider$/, index: 0, selector: true },
    { name: /\blanguages\.match$/, index: 0, selector: true },
    { name: /\bcreateLanguageStatusItem$/, index: 1, selector: true },
];

/**
 * The source with comments and the bodies of regular expressions blanked to
 * spaces, every position kept, so neither can hide or imitate a language.
 * Strings stay as they are: fences and ids live in them.
 */
function blankCommentsAndRegexes(source: string): string {
    const out = source.split('');
    const blank = (from: number, to: number) => {
        for (let k = from; k < to; k++) {
            if (out[k] !== '\n') {
                out[k] = ' ';
            }
        }
    };
    let previous = '';
    for (let i = 0; i < source.length; i++) {
        const c = source[i];
        if (c === '/' && source[i + 1] === '/') {
            const end = source.indexOf('\n', i);
            blank(i, end === -1 ? source.length : end);
            i = end === -1 ? source.length : end - 1;
        } else if (c === '/' && source[i + 1] === '*') {
            const end = source.indexOf('*/', i + 2);
            blank(i, end === -1 ? source.length : end + 2);
            i = end === -1 ? source.length : end + 1;
        } else if (c === "'" || c === '"' || c === '`') {
            for (i++; i < source.length && source[i] !== c; i++) {
                if (source[i] === '\\') {
                    i++;
                }
            }
            previous = c;
        } else if (c === '/' && (previous === '' || '(,=:[!&|?{};'.includes(previous))) {
            let j = i + 1;
            for (let inClass = false; j < source.length && source[j] !== '\n'; j++) {
                if (source[j] === '\\') {
                    j++;
                } else if (source[j] === '[') {
                    inClass = true;
                } else if (source[j] === ']') {
                    inClass = false;
                } else if (source[j] === '/' && !inClass) {
                    break;
                }
            }
            blank(i + 1, j);
            i = j;
            previous = ')';
        } else if (!/\s/.test(c)) {
            previous = c;
        }
    }
    return out.join('');
}

/**
 * The top-level items of the call or array whose `(` or `[` is at `open`, in
 * source that has been through blankCommentsAndRegexes.
 */
function topLevelItems(source: string, open: number): string[] {
    const items: string[] = [];
    let start = open + 1;
    let depth = 0;
    for (let i = open; i < source.length; i++) {
        const c = source[i];
        if (c === "'" || c === '"' || c === '`') {
            for (i++; i < source.length && source[i] !== c; i++) {
                if (source[i] === '\\') {
                    i++;
                }
            }
        } else if (c === '(' || c === '[' || c === '{') {
            depth++;
        } else if (c === ')' || c === ']' || c === '}') {
            depth--;
            if (depth === 0) {
                items.push(source.slice(start, i).trim());
                return items.filter((item, index, all) => item !== '' || index < all.length - 1);
            }
        } else if (c === ',' && depth === 1) {
            items.push(source.slice(start, i).trim());
            start = i + 1;
        }
    }
    return items;
}

/**
 * Whether a selector argument names one of XLIDE's languages as a bare
 * string: the selector itself, or an item of a selector array. A filter's
 * `language:` key is caught by LANGUAGE_PATTERNS; any other string in a
 * filter, such as a `pattern`, is not a language.
 */
function selectorNamesLanguage(arg: string): boolean {
    if (QUOTED_ONLY.test(arg)) {
        return true;
    }
    return arg.startsWith('[') && topLevelItems(arg, 0).some((item) => QUOTED_ONLY.test(item));
}

/** Each place in `original` that states one of XLIDE's languages as a bare string. */
function bareLanguages(original: string): Array<{ index: number; text: string }> {
    const source = blankCommentsAndRegexes(original);
    const found: Array<{ index: number; text: string }> = [];
    for (const pattern of LANGUAGE_PATTERNS) {
        for (const match of source.matchAll(pattern)) {
            found.push({ index: match.index ?? 0, text: match[0] });
        }
    }
    for (const call of source.matchAll(/([\w$.]+)\s*\(/g)) {
        const rule = LANGUAGE_ARGUMENTS.find((candidate) => candidate.name.test(call[1]));
        if (!rule) {
            continue;
        }
        const open = (call.index ?? 0) + call[0].length - 1;
        const arg = topLevelItems(source, open)[rule.index] ?? '';
        if (rule.selector ? selectorNamesLanguage(arg) : QUOTED_ONLY.test(arg)) {
            found.push({ index: call.index ?? 0, text: `${call[1]}(… ${arg} …)` });
        }
    }
    return found;
}

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

    it('are found in each way the source can state a language', () => {
        const bare = [
            "if (document.languageId === 'vba') {}",
            "if ('vba' !== editor.document.languageId) {}",
            "assert.equal(document.languageId, 'xml');",
            "assert.equal('xml', renamed?.languageId);",
            "const selector = [{ scheme: 'file', 'language' : 'vba' }];",
            "vscode.languages.registerHoverProvider('vba', provider);",
            "vscode.languages.registerHoverProvider(['xml', 'vba'], provider);",
            "vscode.languages.createLanguageStatusItem('status', 'xlide-vba');",
            "vscode.languages.match('vba', document);",
            "vscode.workspace.getConfiguration('editor', 'xlide-vba');",
            "md.appendCodeblock(signatureOf(a, b),\n    'vba');",
            "md.appendCodeblock(signature, 'vba',);",
            "md.appendCodeblock(signatureOf('a;b'), 'vba');",
            "md.appendCodeblock(signature, /* language */ 'vba');",
            "md.appendCodeblock(signature /* ), */, // (\n    'vba');",
            "registerHoverProvider([{ pattern: 'https://example/**' }, 'vba'], provider);",
            "registerHoverProvider({ language: /* VBA */ 'vba' }, provider);",
            "if (document.languageId === /* VBA */ 'vba') {}",
            "md.appendCodeblock(signature.replace(/\\)/g, ''), 'vba');",
            "await vscode.languages.setTextDocumentLanguage(doc, 'xml');",
            "const doc = '```vba\\nx\\n```';",
            "parts.push(`\\`\\`\\`vba\\n${x}`);",
        ];
        for (const source of bare) {
            expect(bareLanguages(source), source).toHaveLength(1);
        }
        const notLanguages = [
            "if (tokenText(toks[0]) === 'vba') {}",
            "const libraries = ['vba', 'office'];",
            "vscode.languages.createDiagnosticCollection('vba');",
            "vscode.workspace.getConfiguration('xlide', document.uri);",
            "md.appendCodeblock(formatLibrary(name, 'vba'));",
            "vscode.languages.createLanguageStatusItem('vba', VBA_SELECTOR);",
            "vscode.languages.registerHoverProvider({ language: VBA_LANGUAGE_ID, pattern: 'vba' }, provider);",
            "vscode.languages.registerHoverProvider(selectorForLibrary('vba'), provider);",
            "vscode.languages.registerHoverProvider([selectorForLibrary('vba')], provider);",
            "await vscode.languages.setTextDocumentLanguage(doc, 'plaintext');",
            "if (document.languageId === 'plaintext') {}",
            "const fence = '```vba-sample';",
            "const fence = '```json';",
            "// document.languageId === 'vba' in a comment",
            "const pattern = /languageId === 'vba'/;",
        ];
        for (const source of notLanguages) {
            expect(bareLanguages(source), source).toEqual([]);
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
            for (const { index, text } of bareLanguages(source)) {
                const line = source.slice(0, index).split('\n').length;
                found.push(`${path.relative(root, file)}:${line}: ${text.replace(/\s+/g, ' ')}`);
            }
        }
        expect(found).toEqual([]);
    });
});
