import { describe, expect, it } from 'vitest';
import { tokenize, tokenizeCached } from '../src/analyzer/lexer/tokenize';
import { parseModule, parseModuleFreshForTests } from '../src/analyzer/parser/parseModule';

// Issue #1442: the reference is an independent fresh parse, never the cache.
// Short sources live in a cache of their own, so parsing them evicted nothing
// and the old 'cold' parse could hand back the very AST under test.

/** Puts `source` in the editor parse cache, as the previous snapshot an edit starts from. */
function seed(source: string) {
    return parseModule(source);
}

/** A fresh parse of `source`, sharing no object with any cached AST, from tokens checked against direct lexing. */
function reference(source: string, incremental?: object) {
    expect(tokenizeCached(source)).toEqual(tokenize(source));
    const fresh = parseModuleFreshForTests(source);
    if (incremental) { expect(fresh).not.toBe(incremental); }
    return fresh;
}
const prefix = Array.from({ length: 140 }, (_, i) => 'Public Sub Pad' + i + '()\r\nDebug.Print 1\r\nEnd Sub\r\n').join('');
const tail = [
    'Private value As Object',
    'Private Sub Probe(Optional n As Long = 1)',
    ' Dim text As String: text = "a ""quoted"" string"',
    ' With ThisWorkbook.Sheets(1)',
    '  .Name = text',
    ' End With',
    ' If n Then',
    '  result = n + _',
    '   1',
    ' End If',
    'End Sub',
    '   ',
].join('\r\n');

describe('unchanged parser prefix reuse', () => {
    it('matches a cold full parse after insertions and deletions at every fixture offset', () => {
        const original = prefix + tail;
        for (let offset = prefix.length; offset <= original.length; offset++) {
            for (const inserted of ['x', "'", '"', ':', '_', '\r', '\n', 'ก้']) {
                const before = seed(original);
                const snapshot = structuredClone(before);
                const changed = original.slice(0, offset) + inserted + original.slice(offset);
                const result = parseModule(changed);
                expect(result, 'insert ' + JSON.stringify(inserted) + ' at ' + offset).toEqual(reference(changed, result));
                expect(before).toEqual(snapshot);
            }
            if (offset < original.length) {
                seed(original);
                const changed = original.slice(0, offset) + original.slice(offset + 1);
                const deleted = parseModule(changed);
                expect(deleted, 'delete at ' + offset).toEqual(reference(changed, deleted));
            }
        }
    }, 60000);

    it('shares unchanged members but reparses the edited procedure with new absolute spans', () => {
        const original = prefix + tail;
        const before = seed(original);
        const snapshot = structuredClone(before);
        const changed = original.replace('.Name = text', '.Name = text & "more"');
        const result = parseModule(changed);
        expect(result.members.filter((member, i) => member === before.members[i]).length).toBeGreaterThan(130);
        expect(result.members[result.members.length - 1]).not.toBe(before.members[before.members.length - 1]);
        expect(result).toEqual(reference(changed, result));
        expect(before).toEqual(snapshot);
    });

    it('preserves module kind, declarations and diagnostics when a tail gains an error', () => {
        const original = 'Attribute VB_PredeclaredId = True\r\n' + prefix + tail;
        const before = seed(original);
        const changed = original.replace('End With', 'End If');
        const result = parseModule(changed);
        expect(result.moduleKind).toBe('class');
        expect(result.diagnostics.length).toBeGreaterThan(0);
        expect(result.members[0]).toBe(before.members[0]);
        expect(result).toEqual(reference(changed, result));
        const next = changed.replace('result = n', 'result = x');
        parseModule(changed);
        const reparsed = parseModule(next);
        expect(reparsed).toEqual(reference(next, reparsed));
    });

    it('retains a full parse for conditional-directive state and large replacements', () => {
        for (const original of [
            prefix + '#If VBA7 Then\r\n' + tail + '\r\n#End If\r\n',
            prefix + tail.replace(' Dim text', '#If VBA7 Then\r\n#End If\r\n Dim text'),
        ]) {
            const before = seed(original);
            const changed = original.replace('.Name = text', '.Name = n');
            const result = parseModule(changed);
            expect(result.members[0]).not.toBe(before.members[0]);
            expect(result).toEqual(reference(changed, result));
        }
        const original = prefix + tail;
        const before = seed(original);
        const changed = original + "' " + 'x'.repeat(200);
        const result = parseModule(changed);
        expect(result.members[0]).not.toBe(before.members[0]);
        expect(result).toEqual(reference(changed, result));
    });

    it('does not reuse a closing word that is extended, a partial colon line, or a CRLF split', () => {
        for (const [original, changed] of [
            [prefix, prefix.trimEnd() + 'Extra\r\n'],
            [prefix + 'Sub A():End Sub:Sub B():End Sub\r\n', prefix + 'Sub A():End Sub:Sub B():End SubExtra\r\n'],
            [prefix + 'Sub A()\r\nEnd Sub\r\n', prefix + 'Sub A()\rx\nEnd Sub\r\n'],
            [prefix + 'If n Then a: Dim b As Long\r\n', prefix + 'If n Then a: Dim b As Object\r\n'],
        ]) {
            seed(original);
            const reparsed = parseModule(changed);
            expect(reparsed).toEqual(reference(changed, reparsed));
        }
    });

    it('uses safe earlier boundaries for long colon lines and prior diagnostics', () => {
        const line = Array.from({ length: 1000 }, (_, i) => 'Dim x' + i + ' As Long').join(':') + '\r\n';
        const original = prefix + line;
        const before = seed(original);
        const changed = original.replace('Dim x999 As Long', 'Dim x999 As Object');
        const result = parseModule(changed);
        expect(result.members[0]).toBe(before.members[0]);
        expect(result).toEqual(reference(changed, result));
        const broken = 'Sub MissingEnd()\r\n' + prefix + tail;
        seed(broken);
        const fixed = broken.replace('.Name = text', '.Name = n');
        const reparsed = parseModule(fixed);
        expect(reparsed).toEqual(reference(fixed, reparsed));
    });
});

describe('the reference comparison itself (#1442)', () => {
    it('is not the cache: short parses never evicted a large snapshot', () => {
        const source = prefix + tail;
        const cached = seed(source);
        for (let i = 0; i < 12; i++) { parseModule("' cache eviction " + i); }
        // What the old helper compared against: the same object, so it proved nothing.
        expect(parseModule(source)).toBe(cached);
        expect(reference(source, cached)).toEqual(cached);
    });

    it('fails for a reference of other source, and for a corrupted incremental AST', () => {
        const original = prefix + tail;
        seed(original);
        const changed = original.replace('.Name = text', '.Name = text & "more"');
        const result = parseModule(changed);
        expect(result).toEqual(reference(changed, result));
        expect(result).not.toEqual(reference(changed.replace('"more"', '"less"')));
        const corrupted = structuredClone(result);
        corrupted.members[corrupted.members.length - 1].span.start += 1;
        expect(corrupted).not.toEqual(reference(changed));
        const renamed = structuredClone(result);
        (renamed.members[0] as { name?: string }).name = 'Elsewhere';
        expect(renamed).not.toEqual(reference(changed));
    });
});
