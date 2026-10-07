import { afterEach, describe, expect, it } from 'vitest';
import {
    tokenize, tokenizeCached, startTokenizeMissLogForTests, stopTokenizeMissLogForTests,
} from '../src/analyzer/lexer/tokenize';
import {
    resolveTypeSemanticTokens, collectHostGlobalTokens,
    collectImplicitMemberMethodTokens, collectHostMemberMethodTokens,
} from '../src/analyzer/semantic/typeSemanticTokens';

afterEach(() => { stopTokenizeMissLogForTests(); });

function shortLookups(label: string) {
    for (let index = 0; index < 20; index++) {
        tokenizeCached(`receiver${label}${index}.Members(${index}).Name`);
    }
}

function source(eol: string, label: string, revision = 0) {
    const padding = Array.from({ length: 500 }, (_, index) => [
        `Public Sub Padding${label}${index}()`, '    Dim local As Long', 'End Sub', '',
    ].join(eol)).join('');
    return padding + [
        'Public Sub Probe()', '    Dim item As Collection',
        `    Set item = New Collection ' revision ${String(revision).padStart(3, '0')}`,
        '    Debug.Print ThisWorkbook.Name', 'End Sub', '',
    ].join(eol);
}

function semantic(source: string) {
    return [
        resolveTypeSemanticTokens(source), collectHostGlobalTokens(source),
        collectImplicitMemberMethodTokens(source), collectHostMemberMethodTokens(source),
    ];
}

describe('module lexer cache under short lookup pressure', () => {
    it.each(['\n', '\r\n', '\r'])('keeps exact module tokens and bounded edit reuse after short lookups (%j)', eol => {
        const original = source(eol, 'Identity');
        const before = tokenizeCached(original);
        const snapshot = structuredClone(before);
        shortLookups('Identity');
        startTokenizeMissLogForTests();
        expect(tokenizeCached(original) === before).toBe(true);
        expect(stopTokenizeMissLogForTests()).toEqual([]);

        const changed = original.replace('revision 000', 'revision 001');
        shortLookups('EditedIdentity');
        startTokenizeMissLogForTests();
        const after = tokenizeCached(changed);
        const misses = stopTokenizeMissLogForTests();
        expect(misses.length).toBeGreaterThan(0);
        expect(Math.max(...misses)).toBeLessThan(160);
        expect(after).toStrictEqual(tokenize(changed));
        expect(before).toStrictEqual(snapshot);
        expect(after[0]).toBe(before[0]);
    });

    it('bounds lexing through the complete semantic pipeline on six fresh revisions', () => {
        const original = source('\r\n', 'Semantic');
        const expected = semantic(original);
        for (let revision = 1; revision <= 6; revision++) {
            const changed = source('\r\n', 'Semantic', revision);
            shortLookups('Semantic' + revision);
            startTokenizeMissLogForTests();
            expect(semantic(changed)).toEqual(expected);
            const misses = stopTokenizeMissLogForTests();
            expect(misses.length).toBeGreaterThan(0);
            expect(Math.max(...misses)).toBeLessThan(160);
        }
    });

    it.each(['\n', '\r\n', '\r'])('matches uncached lexing through mixed cache histories (%j)', eol => {
        const grammar = [
            '#Const Enable = True', '#If Enable Then', 'Public Sub Mixed()',
            '    Dim [Unicode α] As Collection', '    Set [Unicode α] = New Collection',
            '    If TypeOf [Unicode α] Is Collection Then Debug.Print #1/1/2000#',
            '    Open "file" For Output As #1', '    Debug.Print Left( _',
            '        "quoted ""text""", 2): Rem trailing comment',
            'End Sub', '#End If', '',
        ].join(eol);
        const prefix = ("' Stable mixed-cache prefix" + eol).repeat(200);
        let changed = prefix + grammar;
        let state = 12345;
        const next = () => { state = (Math.imul(state, 1664525) + 1013904223) >>> 0; return state; };
        for (let revision = 0; revision < 36; revision++) {
            const before = tokenizeCached(changed);
            const snapshot = structuredClone(before);
            shortLookups('Mixed' + revision);
            if (revision % 3 === 0) { tokenizeCached(prefix + 'Public Sub Sibling' + revision + '()' + eol); }
            const offset = prefix.length + next() % (changed.length - prefix.length);
            const insert = ['', 'x', ' ', '"', "'", '_', ':', '#', 'α', '!', '.', eol][next() % 12];
            const remove = next() % 2;
            changed = changed.slice(0, offset) + insert + changed.slice(offset + remove);
            expect(tokenizeCached(changed)).toStrictEqual(tokenize(changed));
            expect(before).toStrictEqual(snapshot);
        }
    });

    it('preserves exact tokens when edits cross the short/module threshold in both directions', () => {
        const suffix = '\nPublic Sub Threshold()\nDebug.Print "a"\nEnd Sub\n';
        const small = "'" + 'p'.repeat(4095 - suffix.length - 1) + suffix;
        const large = small.replace('"a"', '"ab"');
        expect(small.length).toBe(4095);
        expect(large.length).toBe(4096);
        const before = tokenizeCached(small);
        const snapshot = structuredClone(before);
        shortLookups('Threshold');
        expect(tokenizeCached(large)).toStrictEqual(tokenize(large));
        expect(tokenizeCached(small)).toStrictEqual(tokenize(small));
        expect(before).toStrictEqual(snapshot);
    });

    it('retains full-lexer fallback for newline edits after short lookup pressure', () => {
        const original = source('\r\n', 'Newline');
        tokenizeCached(original);
        shortLookups('Newline');
        const changed = original.replace('Set item = New Collection', 'Set item = New Collection\r\n    Debug.Print item.Count');
        startTokenizeMissLogForTests();
        expect(tokenizeCached(changed)).toStrictEqual(tokenize(changed));
        expect(stopTokenizeMissLogForTests()).toContain(changed.length);
    });

    it('keeps a short lookup cached while distinct large modules fill their own cache', () => {
        const short = 'ReciprocalReceiver.Member';
        const before = tokenizeCached(short);
        for (let index = 0; index < 9; index++) { tokenizeCached(source('\n', 'Reciprocal' + index)); }
        expect(tokenizeCached(short) === before).toBe(true);
    });

    it('distinguishes equal-length module snapshots with identical ends', () => {
        const original = source('\n', 'SameEnds');
        const before = tokenizeCached(original);
        const changed = original.replace('PaddingSameEnds0', 'PaddingSameEndsX');
        expect(changed.length).toBe(original.length);
        expect(changed.slice(-128)).toBe(original.slice(-128));
        shortLookups('SameEnds');
        expect(tokenizeCached(changed)).toStrictEqual(tokenize(changed));
        expect(tokenizeCached(changed) === before).toBe(false);
        const equalCopy = ('prefix' + changed).slice(6);
        expect(tokenizeCached(equalCopy) === tokenizeCached(changed)).toBe(true);
    });

    it('retains bounded capacity for both short strings and module snapshots', () => {
        const original = source('\n', 'Capacity');
        const before = tokenizeCached(original);
        for (let index = 0; index < 9; index++) {
            tokenizeCached(original + `' separate module snapshot ${index}\n`);
        }
        expect(tokenizeCached(original)).not.toBe(before);

        const short = 'CapacityReceiver.Member';
        const fragment = tokenizeCached(short);
        for (let index = 0; index < 9; index++) { tokenizeCached('CapacityReceiver' + index + '.Member'); }
        expect(tokenizeCached(short)).not.toBe(fragment);
        expect(tokenizeCached(short)).toStrictEqual(tokenize(short));
    });
});
