import { describe, expect, it } from 'vitest';
import { applyModuleEdits, readModuleParts, splitModuleLines } from '../src/moduleParts';

// Line numbers, 1-based, as xlide_readModule shows them:
//  1 '@Folder("X")          10     Value = m
//  2 Option Explicit        11 End Property
//  3                        12
//  4 ' Adds one             13 Property Let Value(v As Long)
//  5 Public Function Inc(   14     m = v
//  6     Inc = n + 1        15 End Property
//  7 End Function           16
//  8                        17 Sub Last()
//  9 Property Get Value()   18 End Sub
const MODULE = [
    '\'@Folder("X")',
    'Option Explicit',
    '',
    '\' Adds one',
    'Public Function Inc(n As Long) As Long',
    '    Inc = n + 1',
    'End Function',
    '',
    'Property Get Value() As Long',
    '    Value = m',
    'End Property',
    '',
    'Property Let Value(v As Long)',
    '    m = v',
    'End Property',
    '',
    'Sub Last()',
    'End Sub',
    '',
].join('\r\n');

describe('the lines of a module', () => {
    it('are numbered as the read tool numbers them, with nothing after the last line break', () => {
        expect(splitModuleLines(MODULE)).toHaveLength(18);
        expect(splitModuleLines('a\nb')).toEqual(['a', 'b']);
        expect(splitModuleLines('')).toEqual(['']);
    });
});

describe('reading parts of a module', () => {
    it('gives each range and procedure asked for, in that order, with its lines', () => {
        const result = readModuleParts(MODULE, { ranges: [{ startLine: 1, endLine: 2 }], procedures: ['Inc', 'Property Let Value'] });

        expect(result).toEqual({
            ok: true,
            parts: [
                { label: 'lines 1-2', startLine: 1, endLine: 2, text: '\'@Folder("X")\nOption Explicit' },
                // The comment above the header belongs to the procedure; the blank line above that does not.
                { label: 'Function Inc', startLine: 4, endLine: 7, text: '\' Adds one\nPublic Function Inc(n As Long) As Long\n    Inc = n + 1\nEnd Function' },
                { label: 'Property Let Value', startLine: 13, endLine: 15, text: 'Property Let Value(v As Long)\n    m = v\nEnd Property' },
            ],
        });
    });

    it('gives the last procedure without the module s trailing blank lines', () => {
        const result = readModuleParts(`${MODULE}\r\n\r\n`, { procedures: ['Last'] });

        expect(result).toEqual({ ok: true, parts: [{ label: 'Sub Last', startLine: 17, endLine: 18, text: 'Sub Last()\nEnd Sub' }] });
    });

    it('asks for the kind when a name is shared, and says when there is no such procedure', () => {
        expect(readModuleParts(MODULE, { procedures: ['Value'] })).toEqual({
            ok: false,
            message: '"Value" names 2 procedures: Property Get Value, Property Let Value. Give the kind too.',
        });
        expect(readModuleParts(MODULE, { procedures: ['value'] })).toMatchObject({ ok: false });
        expect(readModuleParts(MODULE, { procedures: ['get value'] })).toMatchObject({ ok: true, parts: [{ label: 'Property Get Value' }] });
        expect(readModuleParts(MODULE, { procedures: ['Nope'] })).toEqual({
            ok: false,
            message: 'The module has no procedure "Nope". xlide_listSubs lists its procedures.',
        });
        expect(readModuleParts(MODULE, { procedures: ['Sub Inc'] })).toMatchObject({ ok: false });
    });

    it('refuses a range that is not one, or that runs past the module, rather than cutting it short', () => {
        expect(readModuleParts(MODULE, { ranges: [{ startLine: 17, endLine: 19 }] })).toEqual({
            ok: false,
            message: 'lines 17-19 run past the end of the module, which has 18 lines.',
        });
        expect(readModuleParts(MODULE, { ranges: [{ startLine: 5, endLine: 4 }] })).toMatchObject({ ok: false });
        expect(readModuleParts(MODULE, { ranges: [{ startLine: 0, endLine: 4 }] })).toMatchObject({ ok: false });
        expect(readModuleParts(MODULE, { ranges: [{ startLine: 1.5, endLine: 4 }] })).toMatchObject({ ok: false });
    });
});

describe('editing parts of a module', () => {
    it('applies a procedure, a range and an insertion in one pass, every line number meaning the module as read', () => {
        const result = applyModuleEdits(MODULE, [
            { procedure: 'Inc', text: '\' Adds two\r\nPublic Function Inc(n As Long) As Long\r\n    Inc = n + 2\r\nEnd Function' },
            { insertAfterLine: 18, text: '\r\nSub Added()\r\nEnd Sub' },
            { startLine: 10, endLine: 10, text: '    Value = m * 2' },
        ]);

        expect(result).toEqual({
            ok: true,
            source: [
                '\'@Folder("X")',
                'Option Explicit',
                '',
                '\' Adds two',
                'Public Function Inc(n As Long) As Long',
                '    Inc = n + 2',
                'End Function',
                '',
                'Property Get Value() As Long',
                '    Value = m * 2',
                'End Property',
                '',
                'Property Let Value(v As Long)',
                '    m = v',
                'End Property',
                '',
                'Sub Last()',
                'End Sub',
                '',
                'Sub Added()',
                'End Sub',
                '',
            ].join('\r\n'),
            applied: [
                { label: 'Function Inc', newStartLine: 4, newEndLine: 7 },
                { label: 'after line 18', newStartLine: 19, newEndLine: 21 },
                { label: 'lines 10-10', newStartLine: 10, newEndLine: 10 },
            ],
        });
    });

    it('says where each edit landed once the ones above it moved the lines', () => {
        const result = applyModuleEdits(MODULE, [
            { startLine: 17, endLine: 18, text: 'Sub Last()\r\n    Debug.Print 1\r\nEnd Sub' },
            { insertAfterLine: 2, text: 'Private m As Long' },
            { startLine: 4, endLine: 4, text: '' },
        ]);

        expect(result).toMatchObject({
            ok: true,
            applied: [
                { label: 'lines 17-18', newStartLine: 17, newEndLine: 19 },
                { label: 'after line 2', newStartLine: 3, newEndLine: 3 },
                { label: 'lines 4-4' },
            ],
        });
        const lines = splitModuleLines((result as { source: string }).source);
        expect(lines[2]).toBe('Private m As Long');
        expect(lines[3]).toBe('');
        expect(lines[4]).toBe('Public Function Inc(n As Long) As Long');
        expect(lines.slice(16, 19)).toEqual(['Sub Last()', '    Debug.Print 1', 'End Sub']);
    });

    it('keeps the module s own line breaks, and reads a trailing line break in the text as the end of a line', () => {
        expect(applyModuleEdits('a\nb\n', [{ insertAfterLine: 0, text: 'z' }])).toMatchObject({ ok: true, source: 'z\na\nb\n' });
        expect(applyModuleEdits('a\r\n', [{ insertAfterLine: 0, text: 'z\r\n' }])).toMatchObject({ ok: true, source: 'z\r\na\r\n' });
        expect(applyModuleEdits('a\r\n', [{ insertAfterLine: 1, text: 'z\n\n' }])).toMatchObject({ ok: true, source: 'a\r\nz\r\n\r\n' });
        expect(applyModuleEdits('a\r\n', [{ startLine: 1, endLine: 1, text: '' }])).toMatchObject({ ok: true, source: '' });
    });

    it('refuses edits that touch the same lines, and two insertions at one place', () => {
        expect(applyModuleEdits(MODULE, [
            { startLine: 5, endLine: 7, text: 'x' },
            { procedure: 'Inc', text: 'y' },
        ])).toEqual({
            ok: false,
            message: 'edits[1] (Function Inc) and edits[0] (lines 5-7) touch the same lines. Make them one edit.',
        });
        expect(applyModuleEdits(MODULE, [
            { insertAfterLine: 3, text: 'x' },
            { insertAfterLine: 3, text: 'y' },
        ])).toMatchObject({ ok: false });
        // An insertion at the line a range ends on comes after it: no overlap.
        expect(applyModuleEdits(MODULE, [
            { startLine: 1, endLine: 2, text: 'x' },
            { insertAfterLine: 2, text: 'y' },
        ])).toMatchObject({ ok: true, source: expect.stringMatching(/^x\r\ny\r\n/) });
    });

    it('refuses what it cannot place, naming the edit', () => {
        expect(applyModuleEdits(MODULE, [{ insertAfterLine: 19, text: 'x' }])).toEqual({
            ok: false,
            message: 'edits[0]: insertAfterLine 19 is not a line of the module: 0 inserts at the top, and 18 at the end.',
        });
        expect(applyModuleEdits(MODULE, [{ procedure: 'Value', text: 'x' }])).toMatchObject({ ok: false, message: expect.stringContaining('edits[0]: "Value" names 2 procedures') });
        expect(applyModuleEdits(MODULE, [{ startLine: 2, endLine: 40, text: 'x' }])).toMatchObject({ ok: false, message: expect.stringContaining('run past the end') });
        expect(applyModuleEdits(MODULE, [])).toEqual({ ok: false, message: 'edits is empty: nothing to change.' });
        // The schema names text as required; a call that leaves it out is refused, not crashed on.
        expect(applyModuleEdits(MODULE, [{ procedure: 'Inc' } as never])).toEqual({
            ok: false,
            message: 'edits[0]: an edit is an object with a text field, plus startLine and endLine, insertAfterLine, or procedure.',
        });
    });
});
