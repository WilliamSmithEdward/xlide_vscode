// Pure, dependency-free VBA structural analysis used by the diagnostics
// provider and the smart auto-block editing feature. Keeping this module
// free of any `vscode` import means it can be unit-tested directly with vitest.

export interface VbaStructuralDiagnostic {
    /** 0-based physical line of the relevant token. */
    line: number;
    /** 0-based start column (inclusive). */
    startCol: number;
    /** 0-based end column (exclusive). */
    endCol: number;
    /** Stable diagnostic code for editor integrations. */
    code?: 'missing-block-closer' | 'unmatched-block-closer' | 'module-declaration-in-procedure';
    /** Closing phrase that can repair a missing-block diagnostic. */
    expectedClose?: string;
    /** 0-based physical line before which the missing closer should be inserted. */
    insertLine?: number;
    message: string;
    severity: 'error' | 'warning';
}

export interface VbaStructuralAnalysisOptions {
    /**
     * Returns true for physical source lines that belong to a proven-inactive
     * conditional-compilation branch. Preprocessor lines are still analyzed so
     * #If/#Else/#End If balance remains checked.
     */
    isInactiveLine?: (line: number) => boolean;
}

/** A logical line after string/comment stripping and continuation joining. */
interface LogicalLine {
    /** Stripped, continuation-joined text. */
    text: string;
    /** 0-based physical line where this logical line begins. */
    line: number;
}

type BlockKind =
    | 'Sub' | 'Function' | 'Property'
    | 'If' | 'With' | 'Select' | 'Type' | 'Enum'
    | 'For' | 'Do' | 'While' | 'PreprocessorIf';

export interface VbaSmartBlockOpener {
    /** Closing statement smart-enter should insert. */
    endKeyword: string;
    /** Optional text to place on the indented body line before moving the caret. */
    bodyPrefix?: string;
}

export interface VbaSmartBlockInsertion {
    /** Text that should occupy the editable body line after Enter. */
    bodyText: string;
    /** Number of lines after the editor-created line where `bodyText` lands. */
    bodyLineOffset: number;
    /** Replacement for the editor-created body line. */
    replacementText: string;
}

export type VbaSmartBlockLayout = 'comfy' | 'compact';

export const VBA_BLOCK_INDENT_UNIT = '\t';
export const DEFAULT_VBA_SMART_BLOCK_LAYOUT: VbaSmartBlockLayout = 'comfy';
export const VBA_IDENTIFIER_PATTERN = '[A-Za-z_][A-Za-z0-9_]*';
export const VBA_IDENTIFIER_RE = /[A-Za-z_][A-Za-z0-9_]*/;
export const VBA_IDENTIFIER_NAME_RE = /^[A-Za-z_][A-Za-z0-9_]*$/;

export function normalizeSmartBlockLayout(value: unknown): VbaSmartBlockLayout {
    return value === 'compact' ? 'compact' : DEFAULT_VBA_SMART_BLOCK_LAYOUT;
}

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

const VBA_IDENTIFIER_WORD_RE = /[A-Za-z_][A-Za-z0-9_]*/g;

export interface VbaLoopIteratorSyncEdit {
    /** Absolute source span to replace. */
    span: { start: number; end: number };
    /** Iterator text copied from the edited side of the loop pair. */
    newText: string;
}

export interface VbaIdentifierOccurrence {
    line: number;
    column: number;
    offset: number;
    text: string;
}

interface SmartOpenBlock {
    kind: BlockKind;
    closer: string;
}

interface OpenBlock {
    kind: BlockKind;
    /** 0-based physical line of the opener. */
    line: number;
    /** Friendly descriptor, e.g. "Sub Foo" or "If". */
    label: string;
}

interface ColumnSpan {
    startCol: number;
    endCol: number;
}

interface ModuleDeclarationInProcedureHit {
    label: string;
    span: ColumnSpan;
}

interface PhysicalLine {
    text: string;
    start: number;
    end: number;
}

interface LoopIteratorToken {
    name: string;
    span: { start: number; end: number };
}

interface LoopOpenerLine {
    kind: 'opener';
    iterator: LoopIteratorToken;
}

interface LoopNextLine {
    kind: 'next';
    closeCount: number;
    iterator?: LoopIteratorToken;
}

type LoopLineInfo = LoopOpenerLine | LoopNextLine;

/** The closing phrase expected for each block kind. */
const CLOSE_PHRASE: Record<BlockKind, string> = {
    Sub: 'End Sub',
    Function: 'End Function',
    Property: 'End Property',
    If: 'End If',
    With: 'End With',
    Select: 'End Select',
    Type: 'End Type',
    Enum: 'End Enum',
    For: 'Next',
    Do: 'Loop',
    While: 'Wend',
    PreprocessorIf: '#End If',
};

/** The opener keyword shown when a stray closer has no match. */
const OPEN_WORD: Record<BlockKind, string> = {
    Sub: 'Sub',
    Function: 'Function',
    Property: 'Property',
    If: 'If',
    With: 'With',
    Select: 'Select Case',
    Type: 'Type',
    Enum: 'Enum',
    For: 'For',
    Do: 'Do',
    While: 'While',
    PreprocessorIf: '#If',
};

/**
 * Replaces string-literal contents and trailing comments with spaces so that
 * keyword detection never trips over text inside quotes or comments, while
 * keeping every column aligned with the original line.
 */
export function stripVba(line: string): string {
    const chars = line.split('');
    let inString = false;
    for (let i = 0; i < chars.length; i++) {
        const c = chars[i];
        if (inString) {
            if (c === '"') {
                if (chars[i + 1] === '"') {
                    chars[i] = ' ';
                    chars[i + 1] = ' ';
                    i++;
                } else {
                    chars[i] = ' ';
                    inString = false;
                }
            } else {
                chars[i] = ' ';
            }
        } else if (c === '"') {
            chars[i] = ' ';
            inString = true;
        } else if (c === "'") {
            for (let j = i; j < chars.length; j++) { chars[j] = ' '; }
            break;
        }
    }
    let out = chars.join('');
    // Blank out a `Rem` comment (whole-line form) keeping column alignment.
    const rem = /^(\s*)Rem\b/i.exec(out);
    if (rem) {
        out = out.slice(0, rem[1].length) + ' '.repeat(out.length - rem[1].length);
    }
    return out;
}

/** Precomputes the absolute offset at which each physical line starts. */
export function lineStartOffsets(source: string): number[] {
    const starts = [0];
    for (let i = 0; i < source.length; i++) {
        if (source[i] === '\n') { starts.push(i + 1); }
    }
    return starts;
}

/** Returns the leading spaces/tabs for a source line or snippet of text. */
export function leadingWhitespace(text: string): string {
    return /^[ \t]*/.exec(text)?.[0] ?? '';
}

/**
 * Finds whole-word identifier occurrences while ignoring strings and comments.
 * Offsets are absolute source offsets so callers do not recompute line starts.
 */
export function findIdentifierOccurrences(
    source: string,
    name: string,
): VbaIdentifierOccurrence[] {
    const lines = source.split(/\r\n|\r|\n/);
    const starts = lineStartOffsets(source);
    const lower = name.toLowerCase();
    const out: VbaIdentifierOccurrence[] = [];
    for (let i = 0; i < lines.length; i++) {
        const stripped = stripVba(lines[i]);
        VBA_IDENTIFIER_WORD_RE.lastIndex = 0;
        let m: RegExpExecArray | null;
        while ((m = VBA_IDENTIFIER_WORD_RE.exec(stripped)) !== null) {
            if (m[0].toLowerCase() === lower) {
                out.push({
                    line: i,
                    column: m.index,
                    offset: (starts[i] ?? 0) + m.index,
                    text: m[0],
                });
            }
        }
    }
    return out;
}

/** Splits source, strips each line, and joins `_` line continuations. */
function toLogicalLines(source: string): { stripped: string[]; logical: LogicalLine[] } {
    const physical = source.split(/\r\n|\r|\n/);
    const stripped = physical.map(stripVba);
    const logical: LogicalLine[] = [];
    let i = 0;
    while (i < stripped.length) {
        let text = stripped[i];
        const startLine = i;
        while (/\s_[ \t]*$/.test(text) && i + 1 < stripped.length) {
            text = text.replace(/\s_[ \t]*$/, ' ') + stripped[i + 1];
            i++;
        }
        for (const segment of splitColonStatements(text)) {
            logical.push({ text: segment, line: startLine });
        }
        i++;
    }
    return { stripped, logical };
}

function splitColonStatements(text: string): string[] {
    const out: string[] = [];
    let start = 0;
    for (let i = 0; i < text.length; i++) {
        if (text[i] !== ':') {
            continue;
        }
        out.push(`${' '.repeat(start)}${text.slice(start, i)}`);
        start = i + 1;
    }
    out.push(`${' '.repeat(start)}${text.slice(start)}`);
    return out.filter((segment) => segment.trim().length > 0);
}

/** Detects a block closer on a stripped, trimmed logical line. */
function matchCloser(t: string): BlockKind | undefined {
    if (/^#\s*End\s*If\b/i.test(t) || /^#\s*EndIf\b/i.test(t)) { return 'PreprocessorIf'; }
    const end = /^End\s+(Sub|Function|Property|If|With|Select|Type|Enum)\b/i.exec(t);
    if (end) {
        const w = end[1].toLowerCase();
        if (w === 'sub') { return 'Sub'; }
        if (w === 'function') { return 'Function'; }
        if (w === 'property') { return 'Property'; }
        if (w === 'if') { return 'If'; }
        if (w === 'with') { return 'With'; }
        if (w === 'select') { return 'Select'; }
        if (w === 'type') { return 'Type'; }
        return 'Enum';
    }
    if (/^Loop\b/i.test(t)) { return 'Do'; }
    if (/^Wend\b/i.test(t)) { return 'While'; }
    return undefined;
}

/** Detects a block opener on a stripped, trimmed logical line. */
function matchOpener(t: string): OpenBlock | undefined {
    let m: RegExpExecArray | null;
    if (/^#\s*If\b/i.test(t)) { return { kind: 'PreprocessorIf', line: 0, label: '#If' }; }
    m = /^(?:(?:Public|Private|Friend|Global)\s+)?(?:Static\s+)?(Sub|Function)\s+([A-Za-z_]\w*)/i.exec(t);
    if (m) {
        const kind = (/^sub$/i.test(m[1]) ? 'Sub' : 'Function') as BlockKind;
        return { kind, line: 0, label: `${kind} ${m[2]}` };
    }
    m = /^(?:(?:Public|Private|Friend|Global)\s+)?(?:Static\s+)?Property\s+(Get|Let|Set)\s+([A-Za-z_]\w*)/i.exec(t);
    if (m) { return { kind: 'Property', line: 0, label: `Property ${m[2]}` }; }
    m = /^(?:(?:Public|Private|Global)\s+)?Type\s+([A-Za-z_]\w*)/i.exec(t);
    if (m) { return { kind: 'Type', line: 0, label: `Type ${m[1]}` }; }
    m = /^(?:(?:Public|Private|Global)\s+)?Enum\s+([A-Za-z_]\w*)/i.exec(t);
    if (m) { return { kind: 'Enum', line: 0, label: `Enum ${m[1]}` }; }
    if (/^Select\s+Case\b/i.test(t)) { return { kind: 'Select', line: 0, label: 'Select Case' }; }
    if (/^If\b/i.test(t) && /\bThen\s*$/i.test(t)) { return { kind: 'If', line: 0, label: 'If' }; }
    if (/^For\b/i.test(t)) { return { kind: 'For', line: 0, label: 'For' }; }
    if (/^Do\b/i.test(t)) { return { kind: 'Do', line: 0, label: 'Do' }; }
    if (/^While\b/i.test(t)) { return { kind: 'While', line: 0, label: 'While' }; }
    if (/^With\b/i.test(t)) { return { kind: 'With', line: 0, label: 'With' }; }
    return undefined;
}

function fullLineProblem(
    physical: string[], line: number, message: string, severity: 'error' | 'warning',
    details: Pick<VbaStructuralDiagnostic, 'code' | 'expectedClose' | 'insertLine'> = {},
    span = defaultDiagnosticColumnSpan(physical[line] ?? ''),
): VbaStructuralDiagnostic {
    return {
        line,
        startCol: span.startCol,
        endCol: span.endCol,
        message,
        severity,
        ...details,
    };
}

function defaultDiagnosticColumnSpan(raw: string): ColumnSpan {
    const startCol = raw.length - raw.trimStart().length;
    return {
        startCol,
        endCol: Math.max(raw.length, startCol + 1),
    };
}

function blockOpenerColumnSpan(raw: string, kind: BlockKind): ColumnSpan | undefined {
    const stripped = stripVba(raw);
    switch (kind) {
        case 'Sub':
            return capturedColumnSpan(stripped, /^(\s*)(?:(?:Public|Private|Friend|Global)\s+)?(?:Static\s+)?(Sub)\b/i);
        case 'Function':
            return capturedColumnSpan(stripped, /^(\s*)(?:(?:Public|Private|Friend|Global)\s+)?(?:Static\s+)?(Function)\b/i);
        case 'Property':
            return capturedColumnSpan(stripped, /^(\s*)(?:(?:Public|Private|Friend|Global)\s+)?(?:Static\s+)?(Property\s+(?:Get|Let|Set))\b/i);
        case 'If':
            return capturedColumnSpan(stripped, /^(\s*)(If)\b/i);
        case 'With':
            return capturedColumnSpan(stripped, /^(\s*)(With)\b/i);
        case 'Select':
            return capturedColumnSpan(stripped, /^(\s*)(Select\s+Case)\b/i);
        case 'Type':
            return capturedColumnSpan(stripped, /^(\s*)(?:(?:Public|Private|Global)\s+)?(Type)\b/i);
        case 'Enum':
            return capturedColumnSpan(stripped, /^(\s*)(?:(?:Public|Private|Global)\s+)?(Enum)\b/i);
        case 'For':
            return capturedColumnSpan(stripped, /^(\s*)(For(?:\s+Each)?)\b/i);
        case 'Do':
            return capturedColumnSpan(stripped, /^(\s*)(Do(?:\s+(?:While|Until))?)\b/i);
        case 'While':
            return capturedColumnSpan(stripped, /^(\s*)(While)\b/i);
        case 'PreprocessorIf':
            return capturedColumnSpan(stripped, /^(\s*)(#\s*If)\b/i);
    }
}

function blockCloserColumnSpan(raw: string, kind: BlockKind): ColumnSpan | undefined {
    const stripped = stripVba(raw);
    switch (kind) {
        case 'Sub':
            return capturedColumnSpan(stripped, /^(\s*)(End\s+Sub)\b/i);
        case 'Function':
            return capturedColumnSpan(stripped, /^(\s*)(End\s+Function)\b/i);
        case 'Property':
            return capturedColumnSpan(stripped, /^(\s*)(End\s+Property)\b/i);
        case 'If':
            return capturedColumnSpan(stripped, /^(\s*)(End\s+If)\b/i);
        case 'With':
            return capturedColumnSpan(stripped, /^(\s*)(End\s+With)\b/i);
        case 'Select':
            return capturedColumnSpan(stripped, /^(\s*)(End\s+Select)\b/i);
        case 'Type':
            return capturedColumnSpan(stripped, /^(\s*)(End\s+Type)\b/i);
        case 'Enum':
            return capturedColumnSpan(stripped, /^(\s*)(End\s+Enum)\b/i);
        case 'For':
            return capturedColumnSpan(stripped, /^(\s*)(Next)\b/i);
        case 'Do':
            return capturedColumnSpan(stripped, /^(\s*)(Loop)\b/i);
        case 'While':
            return capturedColumnSpan(stripped, /^(\s*)(Wend)\b/i);
        case 'PreprocessorIf':
            return capturedColumnSpan(stripped, /^(\s*)(#\s*End\s*If|#\s*EndIf)\b/i);
    }
}

function preprocessorBranchColumnSpan(raw: string): ColumnSpan | undefined {
    return capturedColumnSpan(stripVba(raw), /^(\s*)(#\s*ElseIf|#\s*Else)\b/i);
}

function moduleDeclarationInProcedureHit(raw: string): ModuleDeclarationInProcedureHit | undefined {
    const stripped = stripVba(raw);
    const procedure = capturedColumnSpan(
        stripped,
        /^(\s*)(?:(?:Public|Private|Friend|Global)\s+)?(?:Static\s+)?(Sub|Function)\b/i,
    );
    if (procedure) {
        return { label: 'Procedure declarations', span: procedure };
    }

    const property = capturedColumnSpan(
        stripped,
        /^(\s*)(?:(?:Public|Private|Friend|Global)\s+)?(?:Static\s+)?(Property\s+(?:Get|Let|Set))\b/i,
    );
    if (property) {
        return { label: 'Procedure declarations', span: property };
    }

    const typeDeclaration = capturedColumnSpan(
        stripped,
        /^(\s*)(?:(?:Public|Private|Global)\s+)?(Type)\b/i,
    );
    if (typeDeclaration) {
        return { label: 'Type declarations', span: typeDeclaration };
    }

    const enumDeclaration = capturedColumnSpan(
        stripped,
        /^(\s*)(?:(?:Public|Private|Global)\s+)?(Enum)\b/i,
    );
    if (enumDeclaration) {
        return { label: 'Enum declarations', span: enumDeclaration };
    }

    const declare = capturedColumnSpan(
        stripped,
        /^(\s*)(?:(?:Public|Private)\s+)?(Declare)\b/i,
    );
    if (declare) {
        return { label: 'Declare statements', span: declare };
    }

    return undefined;
}

function activeProcedureBlock(stack: OpenBlock[]): OpenBlock | undefined {
    for (let i = stack.length - 1; i >= 0; i--) {
        const block = stack[i];
        if (block.kind === 'Sub' || block.kind === 'Function' || block.kind === 'Property') {
            return block;
        }
    }
    return undefined;
}

function isProcedureBlockKind(kind: BlockKind): boolean {
    return kind === 'Sub' || kind === 'Function' || kind === 'Property';
}

function isConditionalAlternativeProcedureHeader(
    stack: readonly OpenBlock[],
    opener: OpenBlock,
    preprocessorDepth: number,
): boolean {
    if (preprocessorDepth === 0 || !isProcedureBlockKind(opener.kind)) {
        return false;
    }
    const active = stack[stack.length - 1];
    return !!active &&
        active.kind === opener.kind &&
        active.label.toLowerCase() === opener.label.toLowerCase();
}

function isTypeFieldNamedTypeInsideType(stack: readonly OpenBlock[], opener: OpenBlock, text: string): boolean {
    return opener.kind === 'Type' &&
        stack[stack.length - 1]?.kind === 'Type' &&
        /^Type\s+As\b/i.test(text.trim());
}

function isPreprocessorLine(trimmed: string): boolean {
    return /^#\s*(?:Const|If|ElseIf|Else|End\s*If|EndIf)\b/i.test(trimmed);
}

function capturedColumnSpan(stripped: string, pattern: RegExp): ColumnSpan | undefined {
    const match = pattern.exec(stripped);
    if (!match) {
        return undefined;
    }
    const prefix = match[1] ?? '';
    const text = match[2] ?? match[0].slice(prefix.length);
    const startCol = prefix.length + match[0].slice(prefix.length).indexOf(text);
    return {
        startCol,
        endCol: startCol + text.length,
    };
}

/**
 * Performs structural block-balance analysis on VBA source, reporting:
 *  - openers with no matching closer (missing `End Sub`, `Next`, ...),
 *  - closers with no matching opener (stray `End If`, `Loop`, ...),
 *  - mismatched nesting (an inner block left unclosed).
 */
export function analyzeVbaStructure(
    source: string,
    options: VbaStructuralAnalysisOptions = {},
): VbaStructuralDiagnostic[] {
    const physical = source.split(/\r\n|\r|\n/);
    const { logical } = toLogicalLines(source);
    const stack: OpenBlock[] = [];
    const preprocessorStack: number[] = [];
    const problems: VbaStructuralDiagnostic[] = [];

    const closeOne = (closerKind: BlockKind, line: number, closerWord: string): void => {
        let idx = -1;
        for (let k = stack.length - 1; k >= 0; k--) {
            if (stack[k].kind === closerKind) { idx = k; break; }
        }
        if (idx === -1) {
            problems.push(fullLineProblem(
                physical, line,
                `'${closerWord}' has no matching '${OPEN_WORD[closerKind]}'.`,
                'error',
                { code: 'unmatched-block-closer' },
                blockCloserColumnSpan(physical[line] ?? '', closerKind),
            ));
            const top = stack[stack.length - 1];
            if (top && isProcedureBlockKind(top.kind) && isProcedureBlockKind(closerKind)) {
                problems.push(fullLineProblem(
                    physical, top.line,
                    `Missing '${CLOSE_PHRASE[top.kind]}' for '${top.label}'.`,
                    'error',
                    {
                        code: 'missing-block-closer',
                        expectedClose: CLOSE_PHRASE[top.kind],
                        insertLine: line,
                    },
                    blockOpenerColumnSpan(physical[top.line] ?? '', top.kind),
                ));
                stack.pop();
            }
            return;
        }
        // Anything above the matched opener was never closed.
        for (let k = stack.length - 1; k > idx; k--) {
            const open = stack[k];
            problems.push(fullLineProblem(
                physical, open.line,
                `Missing '${CLOSE_PHRASE[open.kind]}' for '${open.label}'.`,
                'error',
                {
                    code: 'missing-block-closer',
                    expectedClose: CLOSE_PHRASE[open.kind],
                    insertLine: line,
                },
                blockOpenerColumnSpan(physical[open.line] ?? '', open.kind),
            ));
        }
        stack.length = idx;
    };

    for (const ll of logical) {
        const t = ll.text.trim();
        if (!t) { continue; }
        if (options.isInactiveLine?.(ll.line) && !isPreprocessorLine(t)) {
            continue;
        }

        if (/^#\s*If\b/i.test(t)) {
            preprocessorStack.push(ll.line);
            continue;
        }

        const preprocessorBranch = /^(#\s*ElseIf|#\s*Else)\b/i.exec(t);
        if (preprocessorBranch) {
            if (preprocessorStack.length === 0) {
                problems.push(fullLineProblem(
                    physical, ll.line,
                    `'${preprocessorBranch[1].replace(/\s+/g, ' ')}' has no matching '#If'.`,
                    'error',
                    { code: 'unmatched-block-closer' },
                    preprocessorBranchColumnSpan(physical[ll.line] ?? ''),
                ));
            }
            continue;
        }

        if (/^#\s*End\s*If\b/i.test(t) || /^#\s*EndIf\b/i.test(t)) {
            if (preprocessorStack.length === 0) {
                problems.push(fullLineProblem(
                    physical, ll.line,
                    `'#End If' has no matching '#If'.`,
                    'error',
                    { code: 'unmatched-block-closer' },
                    blockCloserColumnSpan(physical[ll.line] ?? '', 'PreprocessorIf'),
                ));
            } else {
                preprocessorStack.pop();
            }
            continue;
        }

        if (/^Next\b/i.test(t)) {
            const rest = t.replace(/^Next\b/i, '').trim();
            const count = rest === '' ? 1 : rest.split(',').length;
            for (let n = 0; n < count; n++) { closeOne('For', ll.line, 'Next'); }
            continue;
        }

        const closer = matchCloser(t);
        if (closer) {
            const word = closer === 'PreprocessorIf'
                ? '#End If'
                : /^End\b/i.test(t)
                    ? `End ${closer === 'Select' ? 'Select' : closer}`
                    : t.split(/\s+/)[0];
            closeOne(closer, ll.line, word);
            continue;
        }

        const opener = matchOpener(t);
        const isConditionalAlternativeHeader = !!opener &&
            isConditionalAlternativeProcedureHeader(stack, opener, preprocessorStack.length);
        const activeProcedure = activeProcedureBlock(stack);
        if (activeProcedure) {
            const declaration = moduleDeclarationInProcedureHit(physical[ll.line] ?? '');
            const procedureIndent = leadingWhitespace(physical[activeProcedure.line] ?? '').length;
            if (declaration && !isConditionalAlternativeHeader && declaration.span.startCol > procedureIndent) {
                problems.push(fullLineProblem(
                    physical,
                    ll.line,
                    `${declaration.label} must appear in the module declarations section, not inside a procedure.`,
                    'error',
                    { code: 'module-declaration-in-procedure' },
                    declaration.span,
                ));
            }
        }

        if (opener) {
            if (isTypeFieldNamedTypeInsideType(stack, opener, t)) {
                continue;
            }
            if (isConditionalAlternativeHeader) {
                continue;
            }
            stack.push({ ...opener, line: ll.line });
        }
    }

    for (const open of stack) {
        problems.push(fullLineProblem(
            physical, open.line,
            `Missing '${CLOSE_PHRASE[open.kind]}' for '${open.label}'.`,
            'error',
            {
                code: 'missing-block-closer',
                expectedClose: CLOSE_PHRASE[open.kind],
                insertLine: physical.length,
            },
            blockOpenerColumnSpan(physical[open.line] ?? '', open.kind),
        ));
    }
    for (const line of preprocessorStack) {
        problems.push(fullLineProblem(
            physical, line,
            `Missing '${CLOSE_PHRASE.PreprocessorIf}' for '#If'.`,
            'error',
            {
                code: 'missing-block-closer',
                expectedClose: CLOSE_PHRASE.PreprocessorIf,
                insertLine: physical.length,
            },
            blockOpenerColumnSpan(physical[line] ?? '', 'PreprocessorIf'),
        ));
    }

    problems.sort((a, b) => a.line - b.line || a.startCol - b.startCol);
    return problems;
}

/**
 * If the stripped line opens a block that is safe to complete immediately after
 * Enter, returns the closing statement and any body-line prefix.
 */
export function detectSmartBlockOpener(strippedLine: string): VbaSmartBlockOpener | undefined {
    const t = strippedLine.trim();
    if (!t || t.includes(':')) {
        return undefined;
    }

    const opener = matchOpener(t);
    if (!opener) {
        return undefined;
    }
    if (!isCompleteSmartBlockOpener(t, opener.kind)) {
        return undefined;
    }

    if (opener.kind === 'For') {
        const iterator = forIteratorName(t);
        return { endKeyword: iterator ? `Next ${iterator}` : 'Next' };
    }
    if (opener.kind === 'With') {
        return { endKeyword: CLOSE_PHRASE[opener.kind], bodyPrefix: '.' };
    }
    return { endKeyword: CLOSE_PHRASE[opener.kind] };
}

/**
 * Returns the still-open smart-block closers before an offset, ordered from
 * outermost to innermost. Used by keyword completion so close suggestions and
 * Smart Enter share the same block-opening rules.
 */
export function openSmartBlockClosersBefore(
    source: string,
    offset = source.length,
): string[] {
    const safeOffset = Math.max(0, Math.min(offset, source.length));
    const { logical } = toLogicalLines(source.slice(0, safeOffset));
    const stack: SmartOpenBlock[] = [];

    const closeOne = (kind: BlockKind): void => {
        let idx = -1;
        for (let k = stack.length - 1; k >= 0; k--) {
            if (stack[k].kind === kind) {
                idx = k;
                break;
            }
        }
        if (idx >= 0) {
            stack.length = idx;
        }
    };

    for (const ll of logical) {
        const t = ll.text.trim();
        if (!t) { continue; }
        if (/^(#\s*ElseIf|#\s*Else)\b/i.test(t)) { continue; }

        if (/^Next\b/i.test(t)) {
            const rest = t.replace(/^Next\b/i, '').trim();
            const count = rest === '' ? 1 : rest.split(',').length;
            for (let n = 0; n < count; n++) { closeOne('For'); }
            continue;
        }

        const closer = matchCloser(t);
        if (closer) {
            closeOne(closer);
            continue;
        }

        const opener = detectSmartBlockOpener(t);
        const matched = opener ? matchOpener(t) : undefined;
        if (opener && matched) {
            stack.push({ kind: matched.kind, closer: opener.endKeyword });
        }
    }

    return stack.map((open) => open.closer);
}

/**
 * When the edit position is on a simple `For` / `For Each` iterator or its
 * matching `Next name`, returns the paired iterator replacement.
 */
export function resolveLoopIteratorSyncEdit(
    source: string,
    offset: number,
): VbaLoopIteratorSyncEdit | undefined {
    const lines = physicalLines(source);
    if (lines.length === 0) { return undefined; }

    const safeOffset = Math.max(0, Math.min(offset, source.length));
    const lineIndex = physicalLineAtOffset(lines, safeOffset);
    const line = lines[lineIndex];
    const info = parseLoopLine(line);
    if (!info?.iterator || !offsetTouchesSpan(safeOffset, info.iterator.span)) {
        return undefined;
    }

    const counterpart = info.kind === 'opener'
        ? findMatchingNextLine(lines, lineIndex)
        : findMatchingOpenerLine(lines, lineIndex);
    if (!counterpart?.iterator) {
        return undefined;
    }
    if (counterpart.iterator.name === info.iterator.name) {
        return undefined;
    }

    return {
        span: counterpart.iterator.span,
        newText: info.iterator.name,
    };
}

/**
 * Computes the body-line indentation Smart Enter should own after a block
 * opener. If the editor already produced a deeper indent, keep it; otherwise
 * move one configured indentation unit deeper than the opener.
 */
export function smartBlockBodyIndent(
    openerLine: string,
    currentBodyLine = '',
    indentUnit = VBA_BLOCK_INDENT_UNIT,
): string {
    const openerIndent = leadingWhitespace(openerLine);
    const currentIndent = leadingWhitespace(currentBodyLine);
    const expectedIndent = openerIndent + indentUnit;
    if (currentIndent.startsWith(expectedIndent)) {
        return currentIndent;
    }
    return expectedIndent;
}

/**
 * Body-line text Smart Enter should leave after a block opener. For `With`,
 * this includes the seeded leading dot after the computed indentation.
 */
export function smartBlockBodyText(
    openerLine: string,
    currentBodyLine: string,
    opener: Pick<VbaSmartBlockOpener, 'bodyPrefix'>,
    indentUnit = VBA_BLOCK_INDENT_UNIT,
): string {
    return smartBlockBodyIndent(openerLine, currentBodyLine, indentUnit) + (opener.bodyPrefix ?? '');
}

/**
 * Builds the exact body-line replacement for Smart Enter. The normal shape is:
 * opener line, spacer line, one indented editable body line, spacer line, then
 * the closer on its own line.
 */
export function smartBlockInsertion(
    openerLine: string,
    currentBodyLine: string,
    opener: VbaSmartBlockOpener,
    options: {
        eol?: string;
        insertCloser?: boolean;
        indentUnit?: string;
        layout?: VbaSmartBlockLayout;
    } = {},
): VbaSmartBlockInsertion {
    const bodyText = smartBlockBodyText(
        openerLine,
        currentBodyLine,
        opener,
        options.indentUnit,
    );
    const eol = options.eol ?? '\n';
    if (options.insertCloser === false) {
        return {
            bodyText,
            bodyLineOffset: 0,
            replacementText: bodyText,
        };
    }
    const layout = options.layout ?? DEFAULT_VBA_SMART_BLOCK_LAYOUT;
    if (layout === 'compact') {
        return {
            bodyText,
            bodyLineOffset: 0,
            replacementText: `${bodyText}${eol}${leadingWhitespace(openerLine)}${opener.endKeyword}`,
        };
    }
    return {
        bodyText,
        bodyLineOffset: 1,
        replacementText: `${eol}${bodyText}${eol}${eol}${leadingWhitespace(openerLine)}${opener.endKeyword}`,
    };
}

/**
 * When Enter is pressed after a leading-dot member line inside an active With
 * block, the next line should keep the same indentation and seed another dot.
 */
export function withMemberContinuationText(
    source: string,
    previousLineIndex: number,
): string | undefined {
    const lines = source.split(/\r\n|\r|\n/);
    const previousLine = lines[previousLineIndex];
    if (previousLine === undefined) {
        return undefined;
    }
    if (!/^[ \t]*\./.test(stripVba(previousLine))) {
        return undefined;
    }

    const starts = lineStartOffsets(source);
    const previousLineEnd = (starts[previousLineIndex] ?? 0) + previousLine.length;
    if (!openSmartBlockClosersBefore(source, previousLineEnd).includes('End With')) {
        return undefined;
    }

    return `${leadingWhitespace(previousLine)}.`;
}

/**
 * Returns true if the smart-enter block opener already has a compatible closer
 * before the next procedure opener or end of file.
 * `strippedLines` must already have strings/comments removed.
 */
export function isSmartBlockClosedAhead(
    strippedLines: string[], openerIdx: number, opener: VbaSmartBlockOpener,
): boolean {
    const endRe = closerRegex(opener.endKeyword);
    for (let i = openerIdx + 1; i < strippedLines.length; i++) {
        const t = strippedLines[i].trim();
        if (!t) { continue; }
        if (endRe.test(t)) { return true; }
        if (isProcedureOpener(t)) { return false; }
    }
    return false;
}

function forIteratorName(t: string): string | undefined {
    return /^For\s+(?:Each\s+)?([A-Za-z_]\w*)\b/i.exec(t)?.[1];
}

function physicalLines(source: string): PhysicalLine[] {
    const lines: PhysicalLine[] = [];
    let start = 0;
    for (let i = 0; i < source.length; i++) {
        if (source[i] !== '\n') { continue; }
        const end = i > start && source[i - 1] === '\r' ? i - 1 : i;
        lines.push({ text: source.slice(start, end), start, end });
        start = i + 1;
    }
    lines.push({ text: source.slice(start), start, end: source.length });
    return lines;
}

function physicalLineAtOffset(lines: PhysicalLine[], offset: number): number {
    for (let i = 0; i < lines.length; i++) {
        if (offset >= lines[i].start && offset <= lines[i].end) {
            return i;
        }
    }
    return lines.length - 1;
}

function parseLoopLine(line: PhysicalLine): LoopLineInfo | undefined {
    const stripped = stripVba(line.text);

    let m = /^(\s*)For\s+Each\s+([A-Za-z_]\w*)\s+In\b/i.exec(stripped);
    if (m) {
        return { kind: 'opener', iterator: tokenFromMatch(line.start, m, 2) };
    }

    m = /^(\s*)For\s+([A-Za-z_]\w*)\s*=/i.exec(stripped);
    if (m) {
        return { kind: 'opener', iterator: tokenFromMatch(line.start, m, 2) };
    }

    const next = /^(\s*)Next\b(.*)$/i.exec(stripped);
    if (!next) {
        return undefined;
    }

    const closeCount = nextCloseCount(next[2].trim());
    m = /^(\s*)Next\s+([A-Za-z_]\w*)\s*$/i.exec(stripped);
    return {
        kind: 'next',
        closeCount,
        iterator: m ? tokenFromMatch(line.start, m, 2) : undefined,
    };
}

function tokenFromMatch(lineStart: number, match: RegExpExecArray, groupIndex: number): LoopIteratorToken {
    const name = match[groupIndex];
    const col = match[0].indexOf(name);
    return {
        name,
        span: { start: lineStart + col, end: lineStart + col + name.length },
    };
}

function nextCloseCount(rest: string): number {
    if (!rest) { return 1; }
    return Math.max(1, rest.split(',').map((part) => part.trim()).filter(Boolean).length);
}

function findMatchingNextLine(lines: PhysicalLine[], openerIndex: number): LoopNextLine | undefined {
    let depth = 0;
    for (let i = openerIndex + 1; i < lines.length; i++) {
        const info = parseLoopLine(lines[i]);
        if (!info) { continue; }
        if (info.kind === 'opener') {
            depth++;
            continue;
        }

        if (depth === 0) {
            return info.closeCount === 1 ? info : undefined;
        }

        depth -= info.closeCount;
        if (depth < 0) {
            return undefined;
        }
    }
    return undefined;
}

function findMatchingOpenerLine(lines: PhysicalLine[], nextIndex: number): LoopOpenerLine | undefined {
    let depth = 0;
    for (let i = nextIndex - 1; i >= 0; i--) {
        const info = parseLoopLine(lines[i]);
        if (!info) { continue; }
        if (info.kind === 'next') {
            depth += info.closeCount;
            continue;
        }

        if (depth === 0) {
            return info;
        }
        depth--;
    }
    return undefined;
}

function offsetTouchesSpan(offset: number, span: { start: number; end: number }): boolean {
    return offset >= span.start && offset <= span.end;
}

function isCompleteSmartBlockOpener(t: string, kind: BlockKind): boolean {
    switch (kind) {
        case 'With':
            return /^With\s+\S/i.test(t);
        case 'If':
            return /^If\s+\S.+\bThen\s*$/i.test(t);
        case 'For':
            return (
                /^For\s+Each\s+[A-Za-z_]\w*\s+In\s+\S/i.test(t) ||
                /^For\s+[A-Za-z_]\w*\s*=.+\bTo\b.+/i.test(t)
            );
        case 'Do':
            return /^Do\s*$/i.test(t) || /^Do\s+(?:While|Until)\s+\S/i.test(t);
        case 'Select':
            return /^Select\s+Case\s+\S/i.test(t);
        case 'While':
            return /^While\s+\S/i.test(t);
        case 'PreprocessorIf':
            return /^#\s*If\s+.+\bThen\s*$/i.test(t);
        default:
            return true;
    }
}

function isProcedureOpener(t: string): boolean {
    return /^(?:(?:Public|Private|Friend|Global)\s+)?(?:Static\s+)?(?:Sub|Function|Property)\b/i.test(t);
}

function closerRegex(endKeyword: string): RegExp {
    if (/^Next\b/i.test(endKeyword)) {
        return /^Next\b/i;
    }
    const escaped = endKeyword
        .replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
        .replace(/\s+/g, '\\s+');
    return new RegExp(`^${escaped}\\b`, 'i');
}
