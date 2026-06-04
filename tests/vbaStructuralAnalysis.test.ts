import { describe, it, expect } from 'vitest';
import {
    analyzeVbaStructure,
    stripVba,
    detectSmartBlockOpener,
    findIdentifierOccurrences,
    isSmartBlockClosedAhead,
    lineStartOffsets,
    leadingWhitespace,
    openSmartBlockClosersBefore,
    resolveLoopIteratorSyncEdit,
    smartBlockBodyIndent,
    smartBlockBodyText,
    smartBlockInsertion,
    withMemberContinuationText,
    normalizeSmartBlockLayout,
} from '../src/vbaStructuralAnalysis';

describe('analyzeVbaStructure', () => {
    it('reports no problems for a balanced Sub', () => {
        const src = 'Sub Foo()\n    MsgBox 1\nEnd Sub\n';
        expect(analyzeVbaStructure(src)).toEqual([]);
    });

    it('flags a Sub missing End Sub', () => {
        const src = 'Sub Foo()\n    MsgBox 1\n';
        const problems = analyzeVbaStructure(src);
        expect(problems).toHaveLength(1);
        expect(problems[0].line).toBe(0);
        expect(problems[0].message).toContain("Missing 'End Sub'");
        expect(problems[0].message).toContain('Sub Foo');
        expect(problems[0].severity).toBe('error');
        expect(problems[0].code).toBe('missing-block-closer');
        expect(problems[0].expectedClose).toBe('End Sub');
        expect(problems[0].insertLine).toBe(3);
        expect(problems[0].startCol).toBe(0);
        expect(problems[0].endCol).toBe(3);
    });

    it('flags a Function missing End Function', () => {
        const src = 'Function Bar() As Long\n    Bar = 2\n';
        const problems = analyzeVbaStructure(src);
        expect(problems).toHaveLength(1);
        expect(problems[0].message).toContain("Missing 'End Function'");
        expect(problems[0].startCol).toBe(0);
        expect(problems[0].endCol).toBe('Function'.length);
    });

    it('records an in-place replacement for mismatched procedure closers', () => {
        const src =
            'Public Property Get Measurement() As Double\n' +
            '    Measurement = 1\n' +
            'End Function\n';
        const problem = analyzeVbaStructure(src).find((p) => p.expectedClose === 'End Property');

        expect(problem).toMatchObject({
            code: 'missing-block-closer',
            expectedClose: 'End Property',
            expectedCloseReplacement: {
                line: 2,
                startCol: 0,
                endCol: 'End Function'.length,
                text: 'End Property',
            },
        });
    });

    it('flags a stray End If', () => {
        const src = 'Sub Foo()\n    End If\nEnd Sub\n';
        const problems = analyzeVbaStructure(src);
        expect(problems).toHaveLength(1);
        expect(problems[0].line).toBe(1);
        expect(problems[0].message).toContain("'End If' has no matching 'If'");
        expect(problems[0].code).toBe('unmatched-block-closer');
        expect(problems[0].startCol).toBe(4);
        expect(problems[0].endCol).toBe(10);
    });

    it('accepts a balanced multiline If', () => {
        const src = 'Sub Foo()\n    If x Then\n        y = 1\n    End If\nEnd Sub\n';
        expect(analyzeVbaStructure(src)).toEqual([]);
    });

    it('does not treat a single-line If as a block', () => {
        const src = 'Sub Foo()\n    If x Then y = 1\nEnd Sub\n';
        expect(analyzeVbaStructure(src)).toEqual([]);
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
        expect(analyzeVbaStructure(src)).toEqual([]);
    });

    it('flags an inner block left unclosed', () => {
        const src = 'Sub Foo()\n    If x Then\n        y = 1\nEnd Sub\n';
        const problems = analyzeVbaStructure(src);
        // The If is unclosed; End Sub closes the Sub leaving the If reported.
        expect(problems.some((p) => p.message.includes("Missing 'End If'"))).toBe(true);
        const missingIf = problems.find((p) => p.expectedClose === 'End If');
        expect(missingIf?.code).toBe('missing-block-closer');
        expect(missingIf?.insertLine).toBe(3);
        expect(missingIf?.startCol).toBe(4);
        expect(missingIf?.endCol).toBe(6);
    });

    it('ignores block keywords inside strings and comments', () => {
        const src = 'Sub Foo()\n    s = "End Sub"  \' If Then For\nEnd Sub\n';
        expect(analyzeVbaStructure(src)).toEqual([]);
    });

    it('handles line continuations in an If', () => {
        const src = 'Sub Foo()\n    If x = 1 _\n        Then\n        y = 1\n    End If\nEnd Sub\n';
        expect(analyzeVbaStructure(src)).toEqual([]);
    });

    it('closes multiple For loops with Next i, j', () => {
        const src = 'Sub Foo()\n    For i = 1 To 2\n        For j = 1 To 2\n        Next i, j\nEnd Sub\n';
        expect(analyzeVbaStructure(src)).toEqual([]);
    });

    it('does not treat Declare Sub as a block', () => {
        const src = 'Declare Sub Sleep Lib "kernel32" (ByVal ms As Long)\n';
        expect(analyzeVbaStructure(src)).toEqual([]);
    });

    it('balances Type and Enum blocks', () => {
        const src = 'Public Type TPoint\n    X As Long\nEnd Type\nEnum Color\n    Red\nEnd Enum\n';
        expect(analyzeVbaStructure(src)).toEqual([]);
    });

    it('flags indented module declarations inside procedures', () => {
        const src = [
            'Sub Foo()',
            '    Type TPoint',
            '        X As Long',
            '    End Type',
            '    Enum Color',
            '        Red',
            '    End Enum',
            '    Declare Sub Sleep Lib "kernel32" (ByVal ms As Long)',
            '    Private Declare Function GetTickCount Lib "kernel32" () As Long',
            '    Sub Inner()',
            '    End Sub',
            '    Function Nested() As Long',
            '    End Function',
            '    Property Get Value() As Long',
            '    End Property',
            'End Sub',
            '',
        ].join('\n');

        const hits = analyzeVbaStructure(src).filter((problem) =>
            problem.code === 'module-declaration-in-procedure');

        expect(hits.map((hit) => ({
            line: hit.line,
            text: src.split('\n')[hit.line].slice(hit.startCol, hit.endCol),
            message: hit.message,
        }))).toEqual([
            {
                line: 1,
                text: 'Type',
                message: 'Type declarations must appear in the module declarations section, not inside a procedure.',
            },
            {
                line: 4,
                text: 'Enum',
                message: 'Enum declarations must appear in the module declarations section, not inside a procedure.',
            },
            {
                line: 7,
                text: 'Declare',
                message: 'Declare statements must appear in the module declarations section, not inside a procedure.',
            },
            {
                line: 8,
                text: 'Declare',
                message: 'Declare statements must appear in the module declarations section, not inside a procedure.',
            },
            {
                line: 9,
                text: 'Sub',
                message: 'Procedure declarations must appear in the module declarations section, not inside a procedure.',
            },
            {
                line: 11,
                text: 'Function',
                message: 'Procedure declarations must appear in the module declarations section, not inside a procedure.',
            },
            {
                line: 13,
                text: 'Property Get',
                message: 'Procedure declarations must appear in the module declarations section, not inside a procedure.',
            },
        ]);
    });

    it('keeps same-indent next members as missing-closer recovery instead of nested declarations', () => {
        const src = [
            'Sub First()',
            '    value = 1',
            'Sub Second()',
            'End Sub',
            '',
        ].join('\n');
        const problems = analyzeVbaStructure(src);

        expect(problems.map((problem) => problem.code)).toEqual(['missing-block-closer']);
        expect(problems[0].message).toContain("Missing 'End Sub'");
        expect(problems[0].message).toContain('Sub First');
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
        expect(analyzeVbaStructure(src)).toEqual([]);
    });

    it('flags a conditional compilation block missing #End If', () => {
        const src = '#If VBA7 Then\nDeclare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As LongPtr)\n';
        const problems = analyzeVbaStructure(src);
        expect(problems).toHaveLength(1);
        expect(problems[0].message).toContain("Missing '#End If'");
        expect(problems[0].startCol).toBe(0);
        expect(problems[0].endCol).toBe(3);
    });

    it('flags stray conditional compilation branch and closer directives', () => {
        const problems = analyzeVbaStructure('#Else\n#End If\n');
        expect(problems).toHaveLength(2);
        expect(problems[0].message).toContain("has no matching '#If'");
        expect(problems[1].message).toContain("has no matching '#If'");
        expect(problems[0].startCol).toBe(0);
        expect(problems[0].endCol).toBe(5);
        expect(problems[1].startCol).toBe(0);
        expect(problems[1].endCol).toBe(7);
    });

    it('pins structural ranges to block syntax phrases instead of full lines', () => {
        const src = [
            'Sub Foo()',
            '    With ActiveSheet',
            '        For Each cell In Selection',
            'End Sub',
            '',
        ].join('\n');
        const problems = analyzeVbaStructure(src);

        expect(problems.find((p) => p.expectedClose === 'End With')).toMatchObject({
            line: 1,
            startCol: 4,
            endCol: 8,
        });
        expect(problems.find((p) => p.expectedClose === 'Next')).toMatchObject({
            line: 2,
            startCol: 8,
            endCol: 16,
        });
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

describe('shared VBA source text helpers', () => {
    it('computes physical line starts for LF and CRLF sources', () => {
        expect(lineStartOffsets('a\nbb\nccc')).toEqual([0, 2, 5]);
        expect(lineStartOffsets('a\r\nbb\r\nccc')).toEqual([0, 3, 7]);
    });

    it('extracts leading spaces and tabs through one shared rule', () => {
        expect(leadingWhitespace('  \tValue = 1')).toBe('  \t');
        expect(leadingWhitespace('Value = 1')).toBe('');
    });

    it('finds identifier occurrences outside comments and strings with absolute offsets', () => {
        const src = [
            'Sub T()',
            '    Dim value As String',
            '    Debug.Print value, "value"',
            "    ' value in comment",
            'End Sub',
            '',
        ].join('\n');

        expect(findIdentifierOccurrences(src, 'value')).toEqual([
            { line: 1, column: 8, offset: src.indexOf('value'), text: 'value' },
            {
                line: 2,
                column: 16,
                offset: src.indexOf('value,'),
                text: 'value',
            },
        ]);
    });
});

describe('detectSmartBlockOpener procedure headers', () => {
    it('detects Sub', () => {
        expect(detectSmartBlockOpener('Sub Foo()')).toEqual({ endKeyword: 'End Sub' });
    });
    it('detects Public Function', () => {
        expect(detectSmartBlockOpener('Public Function Bar() As Long')).toEqual({ endKeyword: 'End Function' });
    });
    it('detects Property Get', () => {
        expect(detectSmartBlockOpener('Property Get Name() As String')).toEqual({ endKeyword: 'End Property' });
    });
    it('ignores Declare Sub', () => {
        expect(detectSmartBlockOpener('Declare Sub Sleep Lib "k" ()')).toBeUndefined();
    });
    it('ignores non-procedures', () => {
        expect(detectSmartBlockOpener('Dim x As Long')).toBeUndefined();
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

describe('isSmartBlockClosedAhead procedure headers', () => {
    it('returns true when End Sub follows', () => {
        const lines = ['Sub Foo()', '    x = 1', 'End Sub'];
        expect(isSmartBlockClosedAhead(lines, 0, { endKeyword: 'End Sub' })).toBe(true);
    });
    it('returns false when no End before next proc', () => {
        const lines = ['Sub Foo()', '    x = 1', 'Sub Bar()', 'End Sub'];
        expect(isSmartBlockClosedAhead(lines, 0, { endKeyword: 'End Sub' })).toBe(false);
    });
    it('returns false at end of file', () => {
        const lines = ['Sub Foo()', '    x = 1'];
        expect(isSmartBlockClosedAhead(lines, 0, { endKeyword: 'End Sub' })).toBe(false);
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

describe('smartBlockInsertion', () => {
    it('builds the requested expanded CRLF body-line shape for If blocks', () => {
        const opener = detectSmartBlockOpener('    If True Then');
        expect(opener).toBeDefined();

        expect(smartBlockInsertion('    If True Then', '    ', opener!, { eol: '\r\n' })).toEqual({
            bodyText: '    \t',
            bodyLineOffset: 1,
            replacementText: '\r\n    \t\r\n\r\n    End If',
        });
    });

    it('seeds With bodies with a leading dot before the closer line', () => {
        const opener = detectSmartBlockOpener('\tWith ActiveSheet');
        expect(opener).toBeDefined();

        expect(smartBlockInsertion('\tWith ActiveSheet', '\t', opener!, { eol: '\n' })).toEqual({
            bodyText: '\t\t.',
            bodyLineOffset: 1,
            replacementText: '\n\t\t.\n\n\tEnd With',
        });
    });

    it('builds compact blocks from the same Smart Enter helper', () => {
        const opener = detectSmartBlockOpener('    If True Then');
        expect(opener).toBeDefined();

        expect(smartBlockInsertion('    If True Then', '    ', opener!, {
            eol: '\r\n',
            layout: 'compact',
        })).toEqual({
            bodyText: '    \t',
            bodyLineOffset: 0,
            replacementText: '    \t\r\n    End If',
        });
    });

    it('keeps existing-closer Enter compact while preserving body indentation', () => {
        const opener = detectSmartBlockOpener('    For Each item In collection');
        expect(opener).toBeDefined();

        expect(smartBlockInsertion('    For Each item In collection', '        ', opener!, {
            eol: '\r\n',
            insertCloser: false,
        })).toEqual({
            bodyText: '    \t',
            bodyLineOffset: 0,
            replacementText: '    \t',
        });
    });
});

describe('normalizeSmartBlockLayout', () => {
    it('defaults unknown settings to the comfy layout', () => {
        expect(normalizeSmartBlockLayout(undefined)).toBe('comfy');
        expect(normalizeSmartBlockLayout('nonsense')).toBe('comfy');
        expect(normalizeSmartBlockLayout('compact')).toBe('compact');
    });
});

describe('withMemberContinuationText', () => {
    it('continues leading-dot member lines inside an active With block', () => {
        const src = [
            'Sub T()',
            '    With ActiveSheet',
            '        .Range("A1")',
            '',
            '    End With',
            'End Sub',
        ].join('\n');

        expect(withMemberContinuationText(src, 2)).toBe('        .');
    });

    it('uses the same indentation as the previous leading-dot line', () => {
        const src = [
            'Sub T()',
            '\tWith ActiveSheet',
            '\t\t.Range("A1")',
            '',
            '\tEnd With',
            'End Sub',
        ].join('\n');

        expect(withMemberContinuationText(src, 2)).toBe('\t\t.');
    });

    it('does not seed dots outside active With blocks', () => {
        expect(withMemberContinuationText('Sub T()\n    .Range("A1")\nEnd Sub\n', 1)).toBeUndefined();
        expect(withMemberContinuationText('Sub T()\n    With x\n    End With\n    .Range("A1")\nEnd Sub\n', 3)).toBeUndefined();
    });
});
