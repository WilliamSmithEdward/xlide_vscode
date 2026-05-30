import { describe, it, expect } from 'vitest';
import {
    lintVbaSource,
    stripVba,
    detectProcOpener,
    isProcClosedAhead,
} from '../src/vbaLinter';

describe('lintVbaSource', () => {
    it('reports no problems for a balanced Sub', () => {
        const src = 'Sub Foo()\n    MsgBox 1\nEnd Sub\n';
        expect(lintVbaSource(src)).toEqual([]);
    });

    it('flags a Sub missing End Sub', () => {
        const src = 'Sub Foo()\n    MsgBox 1\n';
        const problems = lintVbaSource(src);
        expect(problems).toHaveLength(1);
        expect(problems[0].line).toBe(0);
        expect(problems[0].message).toContain("Missing 'End Sub'");
        expect(problems[0].message).toContain('Sub Foo');
        expect(problems[0].severity).toBe('error');
    });

    it('flags a Function missing End Function', () => {
        const src = 'Function Bar() As Long\n    Bar = 2\n';
        const problems = lintVbaSource(src);
        expect(problems).toHaveLength(1);
        expect(problems[0].message).toContain("Missing 'End Function'");
    });

    it('flags a stray End If', () => {
        const src = 'Sub Foo()\n    End If\nEnd Sub\n';
        const problems = lintVbaSource(src);
        expect(problems).toHaveLength(1);
        expect(problems[0].line).toBe(1);
        expect(problems[0].message).toContain("'End If' has no matching 'If'");
    });

    it('accepts a balanced multiline If', () => {
        const src = 'Sub Foo()\n    If x Then\n        y = 1\n    End If\nEnd Sub\n';
        expect(lintVbaSource(src)).toEqual([]);
    });

    it('does not treat a single-line If as a block', () => {
        const src = 'Sub Foo()\n    If x Then y = 1\nEnd Sub\n';
        expect(lintVbaSource(src)).toEqual([]);
    });

    it('accepts For/Next, Do/Loop, While/Wend, With, Select Case', () => {
        const src = [
            'Sub Foo()',
            '    For i = 1 To 3',
            '        Do',
            '            With obj',
            '                Select Case i',
            '                    Case 1',
            '                        While j < 2',
            '                            j = j + 1',
            '                        Wend',
            '                End Select',
            '            End With',
            '        Loop',
            '    Next i',
            'End Sub',
            '',
        ].join('\n');
        expect(lintVbaSource(src)).toEqual([]);
    });

    it('flags an inner block left unclosed', () => {
        const src = 'Sub Foo()\n    If x Then\n        y = 1\nEnd Sub\n';
        const problems = lintVbaSource(src);
        // The If is unclosed; End Sub closes the Sub leaving the If reported.
        expect(problems.some((p) => p.message.includes("Missing 'End If'"))).toBe(true);
    });

    it('ignores block keywords inside strings and comments', () => {
        const src = 'Sub Foo()\n    s = "End Sub"  \' If Then For\nEnd Sub\n';
        expect(lintVbaSource(src)).toEqual([]);
    });

    it('handles line continuations in an If', () => {
        const src = 'Sub Foo()\n    If x = 1 _\n        Then\n        y = 1\n    End If\nEnd Sub\n';
        expect(lintVbaSource(src)).toEqual([]);
    });

    it('closes multiple For loops with Next i, j', () => {
        const src = 'Sub Foo()\n    For i = 1 To 2\n        For j = 1 To 2\n        Next i, j\nEnd Sub\n';
        expect(lintVbaSource(src)).toEqual([]);
    });

    it('does not treat Declare Sub as a block', () => {
        const src = 'Declare Sub Sleep Lib "kernel32" (ByVal ms As Long)\n';
        expect(lintVbaSource(src)).toEqual([]);
    });

    it('balances Type and Enum blocks', () => {
        const src = 'Public Type TPoint\n    X As Long\nEnd Type\nEnum Color\n    Red\nEnd Enum\n';
        expect(lintVbaSource(src)).toEqual([]);
    });
});

describe('stripVba', () => {
    it('blanks string contents but keeps columns', () => {
        const out = stripVba('x = "hello"');
        expect(out).toHaveLength('x = "hello"'.length);
        expect(out).not.toContain('hello');
    });

    it('blanks a Rem comment', () => {
        const out = stripVba('    Rem this is a note');
        expect(out.trim()).toBe('');
    });
});

describe('detectProcOpener', () => {
    it('detects Sub', () => {
        expect(detectProcOpener('Sub Foo()')).toEqual({ endKeyword: 'End Sub' });
    });
    it('detects Public Function', () => {
        expect(detectProcOpener('Public Function Bar() As Long')).toEqual({ endKeyword: 'End Function' });
    });
    it('detects Property Get', () => {
        expect(detectProcOpener('Property Get Name() As String')).toEqual({ endKeyword: 'End Property' });
    });
    it('ignores Declare Sub', () => {
        expect(detectProcOpener('Declare Sub Sleep Lib "k" ()')).toBeUndefined();
    });
    it('ignores non-procedures', () => {
        expect(detectProcOpener('Dim x As Long')).toBeUndefined();
    });
});

describe('isProcClosedAhead', () => {
    it('returns true when End Sub follows', () => {
        const lines = ['Sub Foo()', '    x = 1', 'End Sub'];
        expect(isProcClosedAhead(lines, 0, 'End Sub')).toBe(true);
    });
    it('returns false when no End before next proc', () => {
        const lines = ['Sub Foo()', '    x = 1', 'Sub Bar()', 'End Sub'];
        expect(isProcClosedAhead(lines, 0, 'End Sub')).toBe(false);
    });
    it('returns false at end of file', () => {
        const lines = ['Sub Foo()', '    x = 1'];
        expect(isProcClosedAhead(lines, 0, 'End Sub')).toBe(false);
    });
});
