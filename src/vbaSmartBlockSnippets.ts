// Pure, dependency-free Smart Block snippet catalog shared by keyword
// completion, the static language configuration, and the Smart Enter tests
// that keep all three aligned as one contract.

import {
    DEFAULT_VBA_SMART_BLOCK_LAYOUT,
    VBA_BLOCK_INDENT_UNIT,
    type VbaSmartBlockLayout,
} from './vbaSmartEnter';

export type VbaSmartBlockSnippetContext = 'statement' | 'modifier' | 'directive';

export interface VbaSmartBlockSnippetSpec {
    label: string;
    detail: string;
    insertText: string;
    contexts: readonly VbaSmartBlockSnippetContext[];
    matchText?: readonly string[];
    /**
     * Concrete opener used by tests to keep snippet scaffolds aligned with
     * Smart Enter and the static language configuration.
     */
    smartEnterExample?: string;
    /** Expected Smart Enter closer for `smartEnterExample`. */
    smartEnterCloser?: string;
}

export function vbaSmartBlockSnippetsFor(
    context: VbaSmartBlockSnippetContext,
    layout: VbaSmartBlockLayout = DEFAULT_VBA_SMART_BLOCK_LAYOUT,
): readonly VbaSmartBlockSnippetSpec[] {
    return smartBlockSnippets(layout).filter((spec) => spec.contexts.includes(context));
}

const B = VBA_BLOCK_INDENT_UNIT;

export const VBA_SMART_BLOCK_SNIPPETS: readonly VbaSmartBlockSnippetSpec[] = smartBlockSnippets();

function smartBlockSnippets(
    layout: VbaSmartBlockLayout = DEFAULT_VBA_SMART_BLOCK_LAYOUT,
): readonly VbaSmartBlockSnippetSpec[] {
    const block = (
        opener: string,
        bodyLines: readonly string[],
        closer: string,
    ): string => smartBlockText(opener, bodyLines, closer, layout);
    return [
        smartBlockSnippet('If', block('If ${1:condition} Then', [B + '$0'], 'End If'), 'If...Then block', {
            smartEnterExample: 'If ready Then',
            smartEnterCloser: 'End If',
        }),
        smartBlockSnippet('If Else', block('If ${1:condition} Then', [B + '$2', 'Else', B + '$0'], 'End If'), 'If...Else block', {
            matchText: ['ifelse'],
            smartEnterExample: 'If ready Then',
            smartEnterCloser: 'End If',
        }),
        smartBlockSnippet('With', block('With ${1:object}', [B + '.$0'], 'End With'), 'With...End With block', {
            smartEnterExample: 'With ActiveSheet',
            smartEnterCloser: 'End With',
        }),
        smartBlockSnippet('For', block('For ${1:i} = ${2:1} To ${3:10}', [B + '$0'], 'Next ${1/(.*)/$1/}'), 'For...Next block', {
            smartEnterExample: 'For i = 1 To 10',
            smartEnterCloser: 'Next i',
        }),
        smartBlockSnippet('For Each', block('For Each ${1:item} In ${2:collection}', [B + '$0'], 'Next ${1/(.*)/$1/}'), 'For Each...Next block', {
            smartEnterExample: 'For Each item In collection',
            smartEnterCloser: 'Next item',
        }),
        smartBlockSnippet('Do While', block('Do While ${1:condition}', [B + '$0'], 'Loop'), 'Do While...Loop block', {
            smartEnterExample: 'Do While ready',
            smartEnterCloser: 'Loop',
        }),
        smartBlockSnippet('Do Until', block('Do Until ${1:condition}', [B + '$0'], 'Loop'), 'Do Until...Loop block', {
            smartEnterExample: 'Do Until done',
            smartEnterCloser: 'Loop',
        }),
        smartBlockSnippet('Do Loop Until', block('Do', [B + '$0'], 'Loop Until ${1:condition}'), 'Do...Loop Until block', {
            matchText: ['dountil'],
            smartEnterExample: 'Do',
            smartEnterCloser: 'Loop',
        }),
        smartBlockSnippet('While', block('While ${1:condition}', [B + '$0'], 'Wend'), 'While...Wend block', {
            smartEnterExample: 'While ready',
            smartEnterCloser: 'Wend',
        }),
        smartBlockSnippet('Select Case', block('Select Case ${1:expression}', [B + 'Case ${2:value}', B + B + '$0'], 'End Select'), 'Select Case block', {
            smartEnterExample: 'Select Case value',
            smartEnterCloser: 'End Select',
        }),
        smartBlockSnippet('Sub', block('Sub ${1:Name}()', [B + '$0'], 'End Sub'), 'Procedure block', {
            contexts: ['statement', 'modifier'],
            smartEnterExample: 'Sub Foo()',
            smartEnterCloser: 'End Sub',
        }),
        smartBlockSnippet('Function', block('Function ${1:Name}() As ${2:Variant}', [B + '$0'], 'End Function'), 'Function block', {
            contexts: ['statement', 'modifier'],
            matchText: ['func'],
            smartEnterExample: 'Function Total() As Variant',
            smartEnterCloser: 'End Function',
        }),
        smartBlockSnippet('Property Get', block('Property Get ${1:Name}() As ${2:Variant}', [B + '$0'], 'End Property'), 'Property Get block', {
            contexts: ['statement', 'modifier'],
            matchText: ['propget'],
            smartEnterExample: 'Property Get Name() As Variant',
            smartEnterCloser: 'End Property',
        }),
        smartBlockSnippet('Property Let', block('Property Let ${1:Name}(ByVal ${2:value} As ${3:Variant})', [B + '$0'], 'End Property'), 'Property Let block', {
            contexts: ['statement', 'modifier'],
            matchText: ['proplet'],
            smartEnterExample: 'Property Let Name(ByVal value As Variant)',
            smartEnterCloser: 'End Property',
        }),
        smartBlockSnippet('Property Set', block('Property Set ${1:Name}(ByVal ${2:value} As ${3:Object})', [B + '$0'], 'End Property'), 'Property Set block', {
            contexts: ['statement', 'modifier'],
            matchText: ['propset'],
            smartEnterExample: 'Property Set Name(ByVal value As Object)',
            smartEnterCloser: 'End Property',
        }),
        smartBlockSnippet('Type', block('Type ${1:Name}', [B + '${2:Field} As ${3:Variant}'], 'End Type'), 'User-defined type block', {
            smartEnterExample: 'Type TPoint',
            smartEnterCloser: 'End Type',
        }),
        smartBlockSnippet('Enum', block('Enum ${1:Name}', [B + '${2:Value1} = ${3:0}'], 'End Enum'), 'Enum block', {
            smartEnterExample: 'Enum Color',
            smartEnterCloser: 'End Enum',
        }),
        smartBlockSnippet('Private Sub', block('Private Sub ${1:Name}()', [B + '$0'], 'End Sub'), 'Private procedure block', {
            smartEnterExample: 'Private Sub Foo()',
            smartEnterCloser: 'End Sub',
        }),
        smartBlockSnippet('Public Sub', block('Public Sub ${1:Name}()', [B + '$0'], 'End Sub'), 'Public procedure block', {
            smartEnterExample: 'Public Sub Foo()',
            smartEnterCloser: 'End Sub',
        }),
        smartBlockSnippet('Private Function', block('Private Function ${1:Name}() As ${2:Variant}', [B + '$0'], 'End Function'), 'Private function block', {
            smartEnterExample: 'Private Function Total() As Variant',
            smartEnterCloser: 'End Function',
        }),
        smartBlockSnippet('Public Function', block('Public Function ${1:Name}() As ${2:Variant}', [B + '$0'], 'End Function'), 'Public function block', {
            smartEnterExample: 'Public Function Total() As Variant',
            smartEnterCloser: 'End Function',
        }),
        smartBlockSnippet('#If', block('#If ${1:condition} Then', [B + '$0'], '#End If'), 'Conditional compilation block', {
            contexts: ['directive'],
            smartEnterExample: '#If VBA7 Then',
            smartEnterCloser: '#End If',
        }),
    ];
}

function smartBlockSnippet(
    label: string,
    insertText: string,
    detail: string,
    options: {
        contexts?: readonly VbaSmartBlockSnippetContext[];
        matchText?: readonly string[];
        smartEnterExample?: string;
        smartEnterCloser?: string;
    } = {},
): VbaSmartBlockSnippetSpec {
    return {
        label,
        insertText,
        detail,
        contexts: options.contexts ?? ['statement'],
        matchText: options.matchText,
        smartEnterExample: options.smartEnterExample,
        smartEnterCloser: options.smartEnterCloser,
    };
}

function blockText(...lines: string[]): string {
    return lines.join('\n');
}

function smartBlockText(
    opener: string,
    bodyLines: readonly string[],
    closer: string,
    layout: VbaSmartBlockLayout = DEFAULT_VBA_SMART_BLOCK_LAYOUT,
): string {
    if (layout === 'compact') {
        return blockText(opener, ...bodyLines, closer);
    }
    return blockText(opener, '', ...bodyLines.flatMap((line) => [line, '']), closer);
}
