import { afterEach, describe, expect, it } from 'vitest';
import { tokenize, tokenizeCached, startTokenizeMissLogForTests, stopTokenizeMissLogForTests } from '../src/analyzer/lexer/tokenize';

afterEach(() => { stopTokenizeMissLogForTests(); });

function fixture(eol: string, body: string) {
    return ("' Stable nearby-edit prefix" + eol).repeat(240) + body +
        ("' Stable nearby-edit suffix" + eol).repeat(40);
}

function pressure(revision: number) {
    for (let index = 0; index < 20; index++) { tokenizeCached(`receiver${revision}_${index}.Members(${index}).Name`); }
}

describe('nearby incremental edits across physical lines', () => {
    it.each(['\n', '\r\n', '\r'])('bounds lexing for coupled nonce and member edits with %j line breaks', eol => {
        const body = [
            'Public Sub Probe()', '    Dim value As Long',
            "    value = 1 ' revision 000", '    ThisWorkbook.Sheets(1).cez',
            'End Sub', '',
        ].join(eol);
        let source = fixture(eol, body);
        for (let revision = 1; revision <= 8; revision++) {
            const before = tokenizeCached(source);
            const snapshot = structuredClone(before);
            const changed = source.replace(/revision \d{3}/, 'revision ' + String(revision).padStart(3, '0'))
                .replace(/\.cez?/, revision % 2 ? '.ce' : '.cez');
            pressure(revision);
            startTokenizeMissLogForTests();
            const after = tokenizeCached(changed);
            const misses = stopTokenizeMissLogForTests();
            expect(misses.length).toBeGreaterThan(0);
            expect(Math.max(...misses)).toBeLessThan(160);
            expect(after).toStrictEqual(tokenize(changed));
            expect(before).toStrictEqual(snapshot);
            expect(after[0] === before[0]).toBe(true);
            source = changed;
        }
    });

    it.each(['\n', '\r\n', '\r'])('keeps lexer metadata exact through deterministic nearby edit histories (%j)', eol => {
        const body = [
            '#Const Enabled = True', '#If Enabled Then', 'Public Sub Mixed()',
            '    Dim [Unicode α] As Collection', '    Set [Unicode α] = New Collection',
            '    value = Left( _', '        "quoted ""text""", 2): Rem comment',
            '    Open "file" For Output As #1', '    Debug.Print #1/1/2000#',
            'End Sub', '#End If', '',
        ].join(eol);
        const prefix = ("' Stable nearby-edit prefix" + eol).repeat(240);
        let source = prefix + body + ("' trailing stable text" + eol).repeat(30);
        let state = 12345;
        const next = () => { state = (Math.imul(state, 1664525) + 1013904223) >>> 0; return state; };
        const inserts = ['', 'x', ' ', '"', "'", '_', ':', '#', 'α', '!', '.', '\n', '\r', '\r\n'];
        for (let revision = 0; revision < 120; revision++) {
            const before = tokenizeCached(source);
            const snapshot = structuredClone(before);
            const first = prefix.length + next() % body.length;
            const second = Math.min(source.length - 1, first + 1 + next() % 60);
            const insertOne = inserts[next() % inserts.length];
            const insertTwo = inserts[next() % inserts.length];
            const changed = source.slice(0, first) + insertOne + source.slice(first + 1, second) + insertTwo + source.slice(second + 1);
            pressure(revision);
            expect(tokenizeCached(changed)).toStrictEqual(tokenize(changed));
            expect(before).toStrictEqual(snapshot);
            source = changed;
        }
    });

    it.each(['\n', '\r\n', '\r'])('uses full lexing for added, removed or converted line breaks (%j)', eol => {
        const original = fixture(eol, 'Public Sub Breaks()' + eol + '    Debug.Print 1' + eol + 'End Sub' + eol);
        const variants = [
            original.replace('Debug.Print 1', 'Debug.Print' + eol + ' 1'),
            original.replace(' 1' + eol + 'End Sub', ' 1End Sub'),
            original.replace('Breaks()' + eol, 'Breaks()' + (eol === '\n' ? '\r\n' : '\n')),
        ];
        for (const changed of variants) {
            tokenizeCached(original);
            startTokenizeMissLogForTests();
            expect(tokenizeCached(changed)).toStrictEqual(tokenize(changed));
            expect(stopTokenizeMissLogForTests()).toContain(changed.length);
        }
    });

    it.each([
        ['\rX\n', '\r\n'], ['\r\n', '\rX\n'],
        ['\r\nX', '\rX\n'], ['X\r\n', '\rX\n'],
        ['\rX\n', 'X\r\n'], ['\rX\n', '\r\nX'],
    ].map(([before, after], index) => [before, after, index] as const))('preserves full-lexer fallback when a boundary edit changes physical breaks (%j to %j)', (before, after, index) => {
        const original = fixture('\n', `Public Sub Boundary${index}()\n` + before + 'Debug.Print 1\nEnd Sub\n');
        const changed = original.replace(before + 'Debug.Print', after + 'Debug.Print');
        tokenizeCached(original);
        startTokenizeMissLogForTests();
        expect(tokenizeCached(changed)).toStrictEqual(tokenize(changed));
        expect(stopTokenizeMissLogForTests()).toContain(changed.length);
    });
});
