import { describe, it, expect } from 'vitest';
import {
    lintVbaSource,
    stripVba,
    detectSmartBlockOpener,
    isSmartBlockClosedAhead,
    openSmartBlockClosersBefore,
    resolveLoopIteratorSyncEdit,
    smartBlockBodyIndent,
    smartBlockBodyText,
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
        expect(problems[0].code).toBe('missing-block-closer');
        expect(problems[0].expectedClose).toBe('End Sub');
        expect(problems[0].insertLine).toBe(3);
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
        expect(problems[0].code).toBe('unmatched-block-closer');
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
        const missingIf = problems.find((p) => p.expectedClose === 'End If');
        expect(missingIf?.code).toBe('missing-block-closer');
        expect(missingIf?.insertLine).toBe(3);
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

    it('balances conditional compilation #If blocks', () => {
        const src = [
            '#If VBA7 Then',
            'Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As LongPtr)',
            '#Else',
            'Declare Sub Sleep Lib "kernel32" (ByVal ms As Long)',
            '#End If',
            '',
        ].join('\n');
        expect(lintVbaSource(src)).toEqual([]);
    });

    it('flags a conditional compilation block missing #End If', () => {
        const src = '#If VBA7 Then\nDeclare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As LongPtr)\n';
        const problems = lintVbaSource(src);
        expect(problems).toHaveLength(1);
        expect(problems[0].message).toContain("Missing '#End If'");
    });

    it('flags stray conditional compilation branch and closer directives', () => {
        const problems = lintVbaSource('#Else\n#End If\n');
        expect(problems).toHaveLength(2);
        expect(problems[0].message).toContain("has no matching '#If'");
        expect(problems[1].message).toContain("has no matching '#If'");
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

describe('detectSmartBlockOpener', () => {
    it('detects procedure headers', () => {
        expect(detectSmartBlockOpener('Public Function Bar() As Long')).toEqual({
            endKeyword: 'End Function',
        });
    });

    it('detects With blocks and seeds leading-dot body editing', () => {
        expect(detectSmartBlockOpener('With Range("A1")')).toEqual({
            endKeyword: 'End With',
            bodyPrefix: '.',
        });
    });

    it('detects multiline If blocks but ignores single-line If statements', () => {
        expect(detectSmartBlockOpener('If ready Then')).toEqual({ endKeyword: 'End If' });
        expect(detectSmartBlockOpener('If ready Then value = 1')).toBeUndefined();
    });

    it('detects For and For Each iterators for matching Next statements', () => {
        expect(detectSmartBlockOpener('For i = 1 To 10')).toEqual({ endKeyword: 'Next i' });
        expect(detectSmartBlockOpener('For Each cell In Selection')).toEqual({ endKeyword: 'Next cell' });
    });

    it('detects common structured block openers', () => {
        expect(detectSmartBlockOpener('Select Case value')).toEqual({ endKeyword: 'End Select' });
        expect(detectSmartBlockOpener('Do')).toEqual({ endKeyword: 'Loop' });
        expect(detectSmartBlockOpener('Do While ready')).toEqual({ endKeyword: 'Loop' });
        expect(detectSmartBlockOpener('Do Until ready')).toEqual({ endKeyword: 'Loop' });
        expect(detectSmartBlockOpener('While ready')).toEqual({ endKeyword: 'Wend' });
        expect(detectSmartBlockOpener('Private Type TPoint')).toEqual({ endKeyword: 'End Type' });
        expect(detectSmartBlockOpener('Enum Color')).toEqual({ endKeyword: 'End Enum' });
        expect(detectSmartBlockOpener('#If VBA7 Then')).toEqual({ endKeyword: '#End If' });
    });

    it('does not auto-block one-line colon statements', () => {
        expect(detectSmartBlockOpener('If ready Then: value = 1')).toBeUndefined();
        expect(detectSmartBlockOpener('For i = 1 To 3: Next i')).toBeUndefined();
    });

    it('ignores Declare Sub', () => {
        expect(detectSmartBlockOpener('Declare Sub Sleep Lib "k" ()')).toBeUndefined();
    });

    it('waits for complete-looking non-procedure openers', () => {
        expect(detectSmartBlockOpener('With')).toBeUndefined();
        expect(detectSmartBlockOpener('For')).toBeUndefined();
        expect(detectSmartBlockOpener('For i = 1')).toBeUndefined();
        expect(detectSmartBlockOpener('Do While')).toBeUndefined();
        expect(detectSmartBlockOpener('Do Until')).toBeUndefined();
        expect(detectSmartBlockOpener('Select Case')).toBeUndefined();
        expect(detectSmartBlockOpener('While')).toBeUndefined();
        expect(detectSmartBlockOpener('#If VBA7')).toBeUndefined();
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

describe('isSmartBlockClosedAhead', () => {
    it('returns true when a matching With closer follows', () => {
        const lines = ['With Range("A1")', '    .Value = 1', 'End With'];
        expect(isSmartBlockClosedAhead(lines, 0, { endKeyword: 'End With', bodyPrefix: '.' })).toBe(true);
    });

    it('treats any Next statement as a For closer', () => {
        const lines = ['For i = 1 To 3', '    Debug.Print i', 'Next'];
        expect(isSmartBlockClosedAhead(lines, 0, { endKeyword: 'Next i' })).toBe(true);
    });

    it('does not scan past another procedure opener', () => {
        const lines = ['With Range("A1")', 'Sub Other()', 'End With'];
        expect(isSmartBlockClosedAhead(lines, 0, { endKeyword: 'End With', bodyPrefix: '.' })).toBe(false);
    });
});

describe('openSmartBlockClosersBefore', () => {
    it('returns open closers from outermost to innermost', () => {
        const src = 'Sub T()\n    With rng\n        If ready Then\n            \n';
        expect(openSmartBlockClosersBefore(src, src.length)).toEqual([
            'End Sub',
            'End With',
            'End If',
        ]);
    });

    it('tracks For and For Each iterator names for close suggestions', () => {
        const src = 'Sub T()\n    For i = 1 To 3\n        For Each cell In Selection\n';
        expect(openSmartBlockClosersBefore(src, src.length).slice(-2)).toEqual([
            'Next i',
            'Next cell',
        ]);
    });

    it('closes matched blocks before reporting the active closer', () => {
        const src = 'Sub T()\n    With rng\n    End With\n    \n';
        expect(openSmartBlockClosersBefore(src, src.length)).toEqual(['End Sub']);
    });

    it('does not treat one-line statements as open smart blocks', () => {
        const src = 'Sub T()\n    If ready Then value = 1\n    \n';
        expect(openSmartBlockClosersBefore(src, src.length)).toEqual(['End Sub']);
    });
});

describe('resolveLoopIteratorSyncEdit', () => {
    const applySync = (source: string, offset: number): string | undefined => {
        const edit = resolveLoopIteratorSyncEdit(source, offset);
        return edit
            ? source.slice(0, edit.span.start) + edit.newText + source.slice(edit.span.end)
            : undefined;
    };

    it('updates a matching Next variable when the For iterator changes', () => {
        const src = [
            'Sub T()',
            '    For idx = 1 To 10',
            '        Debug.Print idx',
            '    Next i',
            'End Sub',
            '',
        ].join('\n');
        const out = applySync(src, src.indexOf('idx =') + 'idx'.length);
        expect(out).toContain('    Next idx');
    });

    it('updates a matching For iterator when the Next variable changes', () => {
        const src = [
            'Sub T()',
            '    For i = 1 To 10',
            '        Debug.Print i',
            '    Next idx',
            'End Sub',
            '',
        ].join('\n');
        const out = applySync(src, src.indexOf('idx') + 'idx'.length);
        expect(out).toContain('    For idx = 1 To 10');
    });

    it('supports For Each iterator pairs', () => {
        const src = [
            'Sub T()',
            '    For Each cell In Selection',
            '        Debug.Print cell.Value',
            '    Next item',
            'End Sub',
            '',
        ].join('\n');
        const out = applySync(src, src.indexOf('cell In') + 'cell'.length);
        expect(out).toContain('    Next cell');
    });

    it('matches nested loop pairs without touching the outer loop', () => {
        const src = [
            'Sub T()',
            '    For i = 1 To 10',
            '        For innerIndex = 1 To 3',
            '            Debug.Print innerIndex',
            '        Next j',
            '    Next i',
            'End Sub',
            '',
        ].join('\n');
        const out = applySync(src, src.indexOf('innerIndex =') + 'innerIndex'.length);
        expect(out).toContain('        Next innerIndex');
        expect(out).toContain('    Next i');
    });

    it('skips bare and multi-variable Next statements', () => {
        const bare = 'Sub T()\n    For idx = 1 To 3\n    Next\nEnd Sub\n';
        expect(resolveLoopIteratorSyncEdit(bare, bare.indexOf('idx =') + 'idx'.length)).toBeUndefined();

        const multi = [
            'Sub T()',
            '    For idx = 1 To 3',
            '        For j = 1 To 3',
            '        Next j, i',
            'End Sub',
            '',
        ].join('\n');
        expect(resolveLoopIteratorSyncEdit(multi, multi.indexOf('idx =') + 'idx'.length)).toBeUndefined();
    });

    it('does not edit when the pair already matches or the edit is outside the iterator token', () => {
        const matched = 'Sub T()\n    For i = 1 To 3\n    Next i\nEnd Sub\n';
        expect(resolveLoopIteratorSyncEdit(matched, matched.indexOf('i =') + 1)).toBeUndefined();

        const outside = 'Sub T()\n    For Each item In collection\n    Next item\nEnd Sub\n';
        expect(resolveLoopIteratorSyncEdit(outside, outside.indexOf('collection') + 3)).toBeUndefined();
    });
});

describe('smartBlockBodyIndent', () => {
    it('indents the body one real tab deeper than the opener when VS Code did not', () => {
        expect(smartBlockBodyIndent('    With ActiveSheet', '    ')).toBe('    \t');
    });

    it('indents If block bodies with the same shared block unit', () => {
        expect(smartBlockBodyIndent('    If ready Then', '    ')).toBe('    \t');
    });

    it('keeps an already deeper body indent when it uses the shared block unit', () => {
        expect(smartBlockBodyIndent('    With ActiveSheet', '    \t\t')).toBe('    \t\t');
    });

    it('normalizes VS Code space auto-indent to the shared block unit', () => {
        expect(smartBlockBodyIndent('\tWhile True', '    ')).toBe('\t\t');
        expect(smartBlockBodyIndent('    While True', '        ')).toBe('    \t');
    });

    it('uses the configured indent unit', () => {
        expect(smartBlockBodyIndent('\tWith ActiveSheet', '\t', '\t')).toBe('\t\t');
        expect(smartBlockBodyIndent('    With ActiveSheet', '    ', '    ')).toBe('        ');
    });
});

describe('smartBlockBodyText', () => {
    it('keeps For Each body lines indented when the matching Next already exists', () => {
        expect(smartBlockBodyText('    For Each item In collection', '    ', {})).toBe('    \t');
    });

    it('keeps If body lines indented when End If already exists', () => {
        expect(smartBlockBodyText('    If ready Then', '    ', {})).toBe('    \t');
    });

    it('adds the With leading-dot after the body indent', () => {
        expect(smartBlockBodyText('    With ActiveSheet', '    ', { bodyPrefix: '.' })).toBe('    \t.');
    });
});
