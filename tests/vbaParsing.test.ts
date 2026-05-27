import { vi, describe, it, expect } from 'vitest';

vi.mock('vscode', () => ({}));
vi.mock('../src/pythonBridge', () => ({ PythonBridge: class PythonBridge {} }));

import { parseVbaModule } from '../src/vbaSymbolIndex';

describe('parseVbaModule', () => {
    it('returns empty array for empty source', () => {
        expect(parseVbaModule('')).toEqual([]);
    });

    it('parses a simple Public Sub', () => {
        const src = 'Sub Hello()\nEnd Sub\n';
        const syms = parseVbaModule(src);
        expect(syms).toHaveLength(1);
        expect(syms[0].name).toBe('Hello');
        expect(syms[0].kind).toBe('Sub');
        expect(syms[0].isPublic).toBe(true);
    });

    it('parses a Private Function', () => {
        const src = 'Private Function Compute() As Long\n    Compute = 42\nEnd Function\n';
        const syms = parseVbaModule(src);
        expect(syms).toHaveLength(1);
        expect(syms[0].name).toBe('Compute');
        expect(syms[0].kind).toBe('Function');
        expect(syms[0].isPublic).toBe(false);
    });

    it('parses Property Get / Let pair', () => {
        const src = [
            'Public Property Get Name() As String',
            '    Name = mName',
            'End Property',
            'Public Property Let Name(ByVal v As String)',
            '    mName = v',
            'End Property',
        ].join('\n');
        const syms = parseVbaModule(src);
        expect(syms).toHaveLength(2);
        expect(syms[0].kind).toBe('PropertyGet');
        expect(syms[1].kind).toBe('PropertyLet');
    });

    it('parses a Const declaration (point symbol)', () => {
        const src = 'Public Const MAX_SIZE As Long = 100\n';
        const syms = parseVbaModule(src);
        expect(syms).toHaveLength(1);
        expect(syms[0].name).toBe('MAX_SIZE');
        expect(syms[0].kind).toBe('Const');
        // Const is a single-line symbol — startLine equals endLine
        expect(syms[0].startLine).toBe(syms[0].endLine);
    });

    it('parses an Enum declaration (point symbol)', () => {
        const src = 'Public Enum Status\n    Running\n    Stopped\nEnd Enum\n';
        const syms = parseVbaModule(src);
        expect(syms).toHaveLength(1);
        expect(syms[0].name).toBe('Status');
        expect(syms[0].kind).toBe('Enum');
    });

    it('parses a Type declaration (point symbol)', () => {
        const src = 'Private Type MyRecord\n    Name As String\nEnd Type\n';
        const syms = parseVbaModule(src);
        expect(syms).toHaveLength(1);
        expect(syms[0].name).toBe('MyRecord');
        expect(syms[0].kind).toBe('Type');
    });

    it('records correct line spans for multiple Subs', () => {
        const src = [
            'Sub Alpha()',   // 0
            'End Sub',       // 1
            '',              // 2
            'Sub Beta()',    // 3
            'End Sub',       // 4
        ].join('\n');
        const syms = parseVbaModule(src);
        expect(syms).toHaveLength(2);
        expect(syms[0].name).toBe('Alpha');
        expect(syms[0].startLine).toBe(0);
        expect(syms[0].endLine).toBe(1);
        expect(syms[1].name).toBe('Beta');
        expect(syms[1].startLine).toBe(3);
        expect(syms[1].endLine).toBe(4);
    });

    it('column points at the identifier, not the keyword', () => {
        const src = 'Public Sub HelloWorld()\nEnd Sub\n';
        const syms = parseVbaModule(src);
        expect(syms[0].column).toBe(src.indexOf('HelloWorld'));
    });

    it('ignores Const inside a procedure body', () => {
        const src = [
            'Sub Calc()',
            '    Const PI As Double = 3.14',
            'End Sub',
        ].join('\n');
        // The Const is inside a Sub — only the Sub should be reported
        const syms = parseVbaModule(src);
        const kinds = syms.map((s) => s.kind);
        expect(kinds).not.toContain('Const');
        expect(syms.some((s) => s.name === 'Calc')).toBe(true);
    });
});
