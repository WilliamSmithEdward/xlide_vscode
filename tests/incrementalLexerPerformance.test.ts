import { afterEach, describe, expect, it } from 'vitest';
import { tokenize, tokenizeCached, startTokenizeMissLogForTests, stopTokenizeMissLogForTests } from '../src/analyzer/lexer/tokenize';

afterEach(() => { stopTokenizeMissLogForTests(); });

const padding = "' immutable padding for large editor source\r\n".repeat(110);
const fixture = [
    'Option Compare Text',
    'Public Sub Пример(Optional value As Long = 1)',
    ' Dim output As String: output = "a ""quoted"" string"',
    ' If value <= 42 Then',
    '  result = value + _',
    '    offset(1, 2)',
    '  ThisWorkbook.Sheets(1).Name = [odd name]',
    "  ' continued comment _",
    '    continuation comment',
    '  Rem another comment',
    '  Open "x" For Binary As #1',
    '  If ready Then Rem conditional comment',
    ' End If',
    'End Sub',
    '   ',
].join('\r\n');

describe('small-edit lexer reuse', () => {
    it('matches a full lex including coordinates, canonical keywords and trivia for every fixture offset', () => {
        const original = padding + fixture;
        const before = tokenizeCached(original);
        const snapshot = structuredClone(before);
        for (let offset = padding.length; offset <= original.length; offset++) {
            for (const inserted of ['x', "'", '"', '_', ':', 'ก้', '\t']) {
                tokenizeCached(original);
                const changed = original.slice(0, offset) + inserted + original.slice(offset);
                expect(tokenizeCached(changed), 'insert ' + JSON.stringify(inserted) + ' at ' + offset)
                    .toStrictEqual(tokenize(changed));
            }
            if (offset < original.length) {
                tokenizeCached(original);
                const changed = original.slice(0, offset) + original.slice(offset + 1);
                expect(tokenizeCached(changed), 'delete at ' + offset).toStrictEqual(tokenize(changed));
            }
        }
        expect(before).toStrictEqual(snapshot);
    }, 20000);

    it('re-lexes only a short window after a late edit and preserves previous snapshots', () => {
        const original = 'Sub Pad()\r\nDebug.Print 1\r\nEnd Sub\r\n'.repeat(1200) +
            'Sub Probe()\r\nThisWorkbook.Sheets(1).Na\r\nEnd Sub\r\n';
        const before = tokenizeCached(original);
        const snapshot = structuredClone(before);
        const offset = original.lastIndexOf('.Na') + 3;
        const changed = original.slice(0, offset) + 'm' + original.slice(offset);
        startTokenizeMissLogForTests();
        const result = tokenizeCached(changed);
        const lengths = stopTokenizeMissLogForTests();
        expect(lengths).toHaveLength(1);
        expect(lengths[0]).toBeLessThan(100);
        expect(result).toStrictEqual(tokenize(changed));
        expect(result[0]).toBe(before[0]);
        expect(before).toStrictEqual(snapshot);
    });

    it('shifts suffix token/trivia offsets without changing physical coordinates or mutating old tokens', () => {
        const original = padding + fixture;
        const before = tokenizeCached(original);
        const snapshot = structuredClone(before);
        const offset = original.indexOf('value <=') + 5;
        const changed = original.slice(0, offset) + 'More' + original.slice(offset);
        expect(tokenizeCached(changed)).toStrictEqual(tokenize(changed));
        expect(before).toStrictEqual(snapshot);
    });

    it('matches full lexes through sequential edits with CR, LF and CRLF boundaries', () => {
        for (const terminator of ['\r', '\n', '\r\n']) {
            let source = fixture.replace(/\r\n/g, terminator) + padding.replace(/\r\n/g, terminator);
            let random = 982451653;
            const next = () => { random = (Math.imul(random, 1664525) + 1013904223) >>> 0; return random; };
            const inserts = ['a', "'", '"', '_', ':', ' ', '\t', 'ก้', '\r', '\n', '😀'];
            for (let i = 0; i < 180; i++) {
                const before = tokenizeCached(source);
                const snapshot = structuredClone(before);
                const offset = i < 3 ? i : next() % (source.length + 1);
                const inserted = inserts[next() % inserts.length];
                const removed = next() % 3;
                source = source.slice(0, offset) + inserted + source.slice(offset + removed);
                expect(tokenizeCached(source), terminator + ' edit ' + i).toStrictEqual(tokenize(source));
                expect(before).toStrictEqual(snapshot);
            }
        }
    }, 20000);

    it('retains module reuse when expression lookups are newer cache entries', () => {
        const original = padding + fixture;
        tokenizeCached(original);
        tokenizeCached('ThisWorkbook.Sheets(1)');
        const changed = original.replace('value <=', 'valueX <=');
        startTokenizeMissLogForTests();
        const result = tokenizeCached(changed);
        expect(stopTokenizeMissLogForTests()).toEqual([expect.any(Number)]);
        expect(result).toStrictEqual(tokenize(changed));
        tokenizeCached(changed);
        tokenizeCached('ThisWorkbook.Sheets(2)');
        startTokenizeMissLogForTests();
        const next = changed.replace('valueX <=', 'valueXY <=');
        expect(tokenizeCached(next)).toStrictEqual(tokenize(next));
        expect(stopTokenizeMissLogForTests()[0]).toBeLessThan(100);
    });

    it('falls back for newline edits, large replacements and lost logical-line boundaries', () => {
        for (const [oldTail, newTail] of [
            ['x = 1\r\ny = 2\r\n', 'x = 1\r\nz = 3\r\ny = 2\r\n'],
            ['x = 1\r\ny = 2\r\n', 'x = 1 _\r\ny = 2\r\n'],
            ['x = 1\r\ny = 2\r\n', 'x = ' + '2'.repeat(200) + '\r\ny = 2\r\n'],
        ]) {
            const original = padding + oldTail;
            const changed = padding + newTail;
            tokenizeCached(original);
            startTokenizeMissLogForTests();
            const result = tokenizeCached(changed);
            const lengths = stopTokenizeMissLogForTests();
            expect(result).toStrictEqual(tokenize(changed));
            expect(lengths).toContain(changed.length);
        }
    });
});
