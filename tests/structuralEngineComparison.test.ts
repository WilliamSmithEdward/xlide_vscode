// Dual-engine comparison harness for audit #74.
//
// XLIDE currently has two structural engines over the same VBA source:
//   - the legacy line/regex engine (src/vbaStructuralDiagnostics.ts), the live
//     producer of missing-block-closer / unmatched-block-closer diagnostics; and
//   - the token-based analyzer parser (src/analyzer/parser/parseModule.ts),
//     whose recovery state emits its own block-balance ParseDiagnostics.
//
// This harness runs BOTH engines over every VBA sample in the repository
// (every ```vba block in syntax_corpus/*.md, every oracle case source, and
// every .bas/.cls/.frm fixture under excel_test_workbook/), normalizes both
// outputs into comparable block-balance signals, and classifies every
// difference. The contract enforced here:
//
//   * every divergence must fall into one of the documented divergence
//     classes below -- an unclassified divergence fails the suite, so any
//     future drift between the engines becomes visible immediately; and
//   * each class has a minimal pinned repro asserting BOTH engines' behavior,
//     so neither engine can change inside a documented class unnoticed.
//
// Verdict recorded by this harness (audit #74): the diff between the engines
// is NOT empty and contains regressions in both directions, so the legacy
// engine cannot be retired for the structural-diagnostics consumer without
// changing user-visible behavior. The classes where the parser is the more
// correct engine are marked "parser correct" below; the classes where the
// legacy engine carries behavior the parser lacks are marked "legacy richer".
//
// Documented divergence classes:
//
//   preprocessor-balance       (legacy richer) The legacy engine checks
//       #If/#ElseIf/#Else/#End If balance; the parser treats directives as
//       opaque ConditionalDirective nodes and never reports them.
//
//   procedure-closer-repair    (legacy richer) For a procedure closed by the
//       WRONG procedure closer (Property Get ... End Function), the legacy
//       engine emits ONE missing-block-closer carrying expectedCloseReplacement
//       fix-it metadata; the parser emits an extra "Unexpected 'End X'"
//       diagnostic for the same line and carries no fix-it metadata.
//
//   next-multi-close           (legacy correct) `Next j, i` closes two For
//       blocks in the VBE (MS-VBAL 5.4.2.5 allows a Next variable list); the
//       parser consumes it as a single closer and falsely reports the outer
//       For as missing Next.
//
//   nested-procedure-header    (split) A procedure header while another
//       procedure is open: the legacy engine nests it silently (no diagnostic
//       when both closers are present); the parser ends the open procedure at
//       the new header and reports its missing closer, which matches the VBE
//       compile error but is new behavior for the diagnostics consumer.
//
//   module-level-block-fragment (legacy richer) A block opener outside any
//       procedure (the realtime-typing corpus fragments): the legacy engine
//       reports the missing closer; the parser does not track block balance
//       for module-level statements.
//
//   module-level-stray-closer  (legacy richer) An extra End Sub/Function/
//       Property after all procedures are closed: legacy reports it as
//       unmatched; the parser ignores stray closers at module level.
//
//   multi-modifier-procedure-header (parser correct) The legacy opener regex
//       accepts at most one visibility modifier, so "Public Private Sub X()"
//       is not recognized as an opener and its End Sub is falsely reported as
//       unmatched. The parser skips any run of leading modifiers.
//
//   line-numbered-statement    (parser correct) The legacy regexes are
//       anchored at line start, so a leading line number ("20  Select Case x")
//       hides the opener and the later closer is falsely reported as
//       unmatched. The parser strips leading line numbers per MS-VBAL 5.4.1.1.
//
//   rem-after-colon            (RESOLVED) Formerly a legacy false positive:
//       stripVba blanked only whole-line Rem comments, so block keywords inside
//       a trailing ": Rem ..." comment leaked into the legacy grammar as fake
//       closers. stripVba now blanks Rem at any statement start (MS-VBAL
//       3.3.5.2), so both engines agree; the class was retired and a pinned
//       convergence repro guards against regression.
//
//   declare-inside-procedure   (split) A Declare statement inside a procedure
//       body: the parser's recovery ends the procedure at the module-level
//       construct and reports the missing closer; the legacy engine keeps the
//       block balanced and reports the misplacement via its separate
//       module-declaration-in-procedure rule instead.

import { describe, expect, it } from 'vitest';
import { parseModule } from '../src/analyzer/parser/parseModule';
import { lineStartOffsets, stripVba } from '../src/vbaSourceScan';
import {
    analyzeVbaStructure,
    matchCloser,
    matchOpener,
} from '../src/vbaStructuralDiagnostics';
import { allStructuralComparisonSamples } from './helpers/structuralCorpus';

interface BlockBalanceRecord {
    kind: 'missing' | 'unmatched';
    /** 0-based physical line the diagnostic is anchored to. */
    line: number;
    /** Expected closing phrase for 'missing' records. */
    closer?: string;
}

// ---------------------------------------------------------------------------
// Normalization: each engine's output -> comparable block-balance records
// ---------------------------------------------------------------------------

function legacyRecords(source: string): BlockBalanceRecord[] {
    const records: BlockBalanceRecord[] = [];
    for (const problem of analyzeVbaStructure(source)) {
        if (problem.code === 'missing-block-closer' || problem.code === 'mismatched-end-keyword') {
            // A mismatched-end-keyword warning is the structural engine's
            // procedure-closer-repair signal (a wrong closer on an open procedure),
            // anchored at the opener like missing-block-closer.
            records.push({ kind: 'missing', line: problem.line, closer: problem.expectedClose });
        } else if (problem.code === 'unmatched-block-closer') {
            records.push({ kind: 'unmatched', line: problem.line });
        }
        // 'module-declaration-in-procedure' is a separate legacy-only rule, not
        // a block-balance signal; it is out of scope for this comparison.
    }
    return records;
}

const PARSER_MISSING_RES = [
    /^Procedure '.*' is missing (End (?:Sub|Function|Property))\.$/,
    /^Block is missing (End If|Next|Loop|Wend|End With|End Select)\.$/,
    /^(?:Type|Enum) block is missing (End Type|End Enum)\.$/,
];
const PARSER_UNMATCHED_RE = /^Unexpected '.+' without a matching opening block\.$/;

function parserRecords(source: string): BlockBalanceRecord[] {
    const starts = lineStartOffsets(source);
    const lineOf = (offset: number): number => {
        let line = 0;
        while (line + 1 < starts.length && starts[line + 1] <= offset) {
            line++;
        }
        return line;
    };
    const records: BlockBalanceRecord[] = [];
    for (const diagnostic of parseModule(source).diagnostics) {
        const missing = PARSER_MISSING_RES
            .map((re) => re.exec(diagnostic.message)?.[1])
            .find((closer) => closer !== undefined);
        if (missing) {
            records.push({ kind: 'missing', line: lineOf(diagnostic.span.start), closer: missing });
            continue;
        }
        if (PARSER_UNMATCHED_RE.test(diagnostic.message)) {
            records.push({ kind: 'unmatched', line: lineOf(diagnostic.span.start) });
        }
    }
    return records;
}

function recordKey(record: BlockBalanceRecord): string {
    return `${record.kind}:${record.line}:${record.closer ?? ''}`;
}

function sortedRecords(records: readonly BlockBalanceRecord[]): BlockBalanceRecord[] {
    return [...records].sort((a, b) => recordKey(a).localeCompare(recordKey(b)));
}

/** Multiset difference (a - b) over normalized record keys. */
function subtractRecords(
    a: readonly BlockBalanceRecord[],
    b: readonly BlockBalanceRecord[],
): BlockBalanceRecord[] {
    const remaining = new Map<string, number>();
    for (const record of b) {
        const key = recordKey(record);
        remaining.set(key, (remaining.get(key) ?? 0) + 1);
    }
    const out: BlockBalanceRecord[] = [];
    for (const record of a) {
        const key = recordKey(record);
        const count = remaining.get(key) ?? 0;
        if (count > 0) {
            remaining.set(key, count - 1);
        } else {
            out.push(record);
        }
    }
    return out;
}

// ---------------------------------------------------------------------------
// Divergence classification
// ---------------------------------------------------------------------------

type DivergenceClass =
    | 'preprocessor-balance'
    | 'procedure-closer-repair'
    | 'next-multi-close'
    | 'nested-procedure-header'
    | 'module-level-block-fragment'
    | 'module-level-stray-closer'
    | 'multi-modifier-procedure-header'
    | 'line-numbered-statement'
    | 'declare-inside-procedure'
    | 'UNCLASSIFIED';

const PROC_HEADER_RE =
    /^\s*(?:(?:Public|Private|Friend|Global|Static)\s+)*(?:Sub|Function|Property)\b/i;
const PROC_CLOSER_RE = /^\s*End\s+(?:Sub|Function|Property)\b/i;
const MULTI_MODIFIER_HEADER_RE =
    /^\s*(?:Public|Private|Friend|Global|Static)\s+(?:Public|Private|Friend|Global|Static)\s+(?:(?:Public|Private|Friend|Global|Static)\s+)*(?:Sub|Function|Property)\b/i;
const DECLARE_RE = /^\s*(?:(?:Public|Private)\s+)?Declare\b/i;
const LINE_NUMBER_PREFIX_RE = /^\s*\d+\s+/;
const MULTI_NEXT_RE = /^\s*(?:\d+\s+)?Next\s+\w+\s*,/i;
const PROC_CLOSER_PHRASE_RE = /^End (?:Sub|Function|Property)$/;

interface ClassificationContext {
    rawLines: string[];
    strippedLines: string[];
    /** Legacy diagnostics carrying the fix-it metadata, for repair detection. */
    legacyReplacementLines: ReadonlySet<number>;
}

function classificationContext(source: string): ClassificationContext {
    const rawLines = source.split(/\r\n|\r|\n/);
    const legacyReplacementLines = new Set<number>();
    for (const problem of analyzeVbaStructure(source)) {
        if (problem.expectedCloseReplacement) {
            legacyReplacementLines.add(problem.expectedCloseReplacement.line);
        }
    }
    return {
        rawLines,
        strippedLines: rawLines.map(stripVba),
        legacyReplacementLines,
    };
}

/** True when a line carries a leading line number hiding an opener/closer. */
function isLineNumberedBlockStatement(stripped: string): boolean {
    if (!LINE_NUMBER_PREFIX_RE.test(stripped)) {
        return false;
    }
    const rest = stripped.replace(LINE_NUMBER_PREFIX_RE, '').trim();
    return matchOpener(rest) !== undefined || matchCloser(rest) !== undefined;
}

function classifyRecord(
    record: BlockBalanceRecord,
    side: 'legacy' | 'parser',
    ctx: ClassificationContext,
): DivergenceClass {
    const stripped = ctx.strippedLines[record.line] ?? '';
    const trimmed = stripped.trim();
    const headerLines: number[] = [];
    const closerLines: number[] = [];
    for (let i = 0; i < ctx.strippedLines.length; i++) {
        if (PROC_HEADER_RE.test(ctx.strippedLines[i])) {
            headerLines.push(i);
        }
        if (PROC_CLOSER_RE.test(ctx.strippedLines[i])) {
            closerLines.push(i);
        }
    }

    if (side === 'legacy') {
        if (
            record.closer === '#End If' ||
            (record.kind === 'unmatched' && trimmed.startsWith('#'))
        ) {
            return 'preprocessor-balance';
        }
        if (
            record.kind === 'unmatched' &&
            ctx.strippedLines.some(
                (line, i) => i < record.line && isLineNumberedBlockStatement(line),
            )
        ) {
            return 'line-numbered-statement';
        }
        if (
            record.kind === 'missing' &&
            ctx.strippedLines.some(
                (line, i) => i > record.line && isLineNumberedBlockStatement(line),
            )
        ) {
            return 'line-numbered-statement';
        }
        if (record.kind === 'unmatched' && PROC_CLOSER_RE.test(stripped)) {
            const headersBefore = headerLines.filter((l) => l < record.line).length;
            const closersBefore = closerLines.filter((l) => l < record.line).length;
            if (headersBefore <= closersBefore) {
                return 'module-level-stray-closer';
            }
            if (ctx.strippedLines.some(
                (line, i) => i < record.line && MULTI_MODIFIER_HEADER_RE.test(line),
            )) {
                return 'multi-modifier-procedure-header';
            }
        }
        if (
            record.kind === 'missing' &&
            !headerLines.some((l) => l <= record.line)
        ) {
            return 'module-level-block-fragment';
        }
        return 'UNCLASSIFIED';
    }

    // side === 'parser'
    if (
        record.kind === 'unmatched' &&
        PROC_CLOSER_RE.test(stripped) &&
        ctx.legacyReplacementLines.has(record.line)
    ) {
        return 'procedure-closer-repair';
    }
    if (record.kind === 'missing' && record.closer === 'Next' &&
        ctx.strippedLines.some((line, i) => i >= record.line && MULTI_NEXT_RE.test(line))
    ) {
        return 'next-multi-close';
    }
    if (
        record.kind === 'missing' &&
        record.closer !== undefined &&
        PROC_CLOSER_PHRASE_RE.test(record.closer)
    ) {
        const nextHeader = headerLines.find((l) => l > record.line);
        const interveningCloser = closerLines.some(
            (l) => l > record.line && nextHeader !== undefined && l < nextHeader,
        );
        if (nextHeader !== undefined && !interveningCloser) {
            return 'nested-procedure-header';
        }
        const declareLine = ctx.strippedLines.findIndex(
            (line, i) => i > record.line && DECLARE_RE.test(line),
        );
        if (
            declareLine > record.line &&
            !headerLines.some((l) => l > record.line && l < declareLine)
        ) {
            return 'declare-inside-procedure';
        }
    }
    return 'UNCLASSIFIED';
}

// ---------------------------------------------------------------------------
// Corpus-wide comparison
// ---------------------------------------------------------------------------

describe('structural engine comparison (audit #74)', () => {
    const samples = allStructuralComparisonSamples();

    it('collects a meaningful sample population', () => {
        expect(samples.length).toBeGreaterThan(700);
    });

    it('classifies every block-balance divergence between the engines', () => {
        const counts = new Map<DivergenceClass, number>();
        const unclassified: string[] = [];
        let divergentSamples = 0;

        for (const sample of samples) {
            const legacy = legacyRecords(sample.source);
            const parser = parserRecords(sample.source);
            const legacyOnly = subtractRecords(legacy, parser);
            const parserOnly = subtractRecords(parser, legacy);
            if (legacyOnly.length === 0 && parserOnly.length === 0) {
                continue;
            }
            divergentSamples++;
            const ctx = classificationContext(sample.source);
            const report = (record: BlockBalanceRecord, side: 'legacy' | 'parser'): void => {
                const cls = classifyRecord(record, side, ctx);
                counts.set(cls, (counts.get(cls) ?? 0) + 1);
                if (cls === 'UNCLASSIFIED') {
                    unclassified.push(
                        `${sample.id}: ${side}-only ${record.kind}@${record.line}` +
                        `${record.closer ? ` (${record.closer})` : ''}: ` +
                        `${(ctx.rawLines[record.line] ?? '').trim()}`,
                    );
                }
            };
            for (const record of legacyOnly) {
                report(record, 'legacy');
            }
            for (const record of parserOnly) {
                report(record, 'parser');
            }
        }

        // eslint-disable-next-line no-console
        console.log(
            `[audit #74] samples=${samples.length} divergentSamples=${divergentSamples} ` +
            `classes=${JSON.stringify(Object.fromEntries(counts))}`,
        );

        expect(
            unclassified,
            `Unexplained block-balance divergence between the legacy structural engine ` +
            `and the analyzer parser:\n${unclassified.join('\n')}`,
        ).toEqual([]);
    });
});

// ---------------------------------------------------------------------------
// Pinned minimal repros: one per documented divergence class. These assert
// BOTH engines' current behavior so any change inside a documented class is
// caught here with a readable diff.
// ---------------------------------------------------------------------------

describe('structural engine divergence repros (audit #74)', () => {
    it('preprocessor-balance: only the legacy engine checks #If balance', () => {
        const src = '#If WIN64 Then\nDebug.Print 1\n';
        expect(sortedRecords(legacyRecords(src))).toEqual([
            { kind: 'missing', line: 0, closer: '#End If' },
        ]);
        expect(parserRecords(src)).toEqual([]);
    });

    it('procedure-closer-repair: legacy folds a wrong procedure closer into one repairable diagnostic', () => {
        const src =
            'Public Property Get Measurement() As Double\n' +
            '    Measurement = 1\n' +
            'End Function\n';
        expect(sortedRecords(legacyRecords(src))).toEqual([
            { kind: 'missing', line: 0, closer: 'End Property' },
        ]);
        // The fix-it metadata the parser route does not carry today:
        expect(analyzeVbaStructure(src)[0].expectedCloseReplacement).toMatchObject({
            line: 2,
            text: 'End Property',
        });
        expect(sortedRecords(parserRecords(src))).toEqual([
            { kind: 'missing', line: 0, closer: 'End Property' },
            { kind: 'unmatched', line: 2 },
        ]);
    });

    it('next-multi-close: the parser does not honor a Next variable list', () => {
        const src =
            'Sub S()\n' +
            '    For i = 1 To 2\n' +
            '        For j = 1 To 2\n' +
            '            x = x + 1\n' +
            '    Next j, i\n' +
            'End Sub\n';
        expect(legacyRecords(src)).toEqual([]);
        expect(sortedRecords(parserRecords(src))).toEqual([
            { kind: 'missing', line: 1, closer: 'Next' },
        ]);
    });

    it('nested-procedure-header: the parser ends the outer procedure at the inner header', () => {
        const src =
            'Public Sub Outer()\n' +
            '    Public Sub Inner()\n' +
            '    End Sub\n' +
            'End Sub\n';
        expect(legacyRecords(src)).toEqual([]);
        expect(sortedRecords(parserRecords(src))).toEqual([
            { kind: 'missing', line: 0, closer: 'End Sub' },
        ]);
    });

    it('module-level-block-fragment: only the legacy engine tracks blocks outside procedures', () => {
        const src = 'If x > 0 Then\n';
        expect(sortedRecords(legacyRecords(src))).toEqual([
            { kind: 'missing', line: 0, closer: 'End If' },
        ]);
        expect(parserRecords(src)).toEqual([]);
    });

    it('module-level-stray-closer: only the legacy engine flags an extra End Sub', () => {
        const src = 'Public Sub Demo()\nEnd Sub\nEnd Sub\n';
        expect(sortedRecords(legacyRecords(src))).toEqual([
            { kind: 'unmatched', line: 2 },
        ]);
        expect(parserRecords(src)).toEqual([]);
    });

    it('multi-modifier-procedure-header: the legacy opener regex misses doubled modifiers', () => {
        const src = 'Public Private Sub Demo()\nEnd Sub\n';
        expect(sortedRecords(legacyRecords(src))).toEqual([
            { kind: 'unmatched', line: 1 },
        ]);
        expect(parserRecords(src)).toEqual([]);
    });

    it('line-numbered-statement: the legacy engine misses openers behind line numbers', () => {
        const src =
            'Sub S()\n' +
            '10 Select Case x\n' +
            '    Case 1\n' +
            '        y = 1\n' +
            'End Select\n' +
            'End Sub\n';
        expect(sortedRecords(legacyRecords(src))).toEqual([
            { kind: 'unmatched', line: 4 },
        ]);
        expect(parserRecords(src)).toEqual([]);
    });

    it('rem-after-colon: RESOLVED - legacy stripping now blanks a trailing ": Rem ..." comment like the parser', () => {
        // Previously a legacy-only divergence: stripVba blanked only whole-line
        // Rem comments, so block keywords inside a `: Rem ...` comment leaked as
        // phantom closers (a live false positive). stripVba now blanks Rem at any
        // statement start, so both engines agree this is clean code.
        const src =
            'Sub S()\n' +
            '    Debug.Print 1: Rem hidden: End If\n' +
            'End Sub\n';
        expect(legacyRecords(src)).toEqual([]);
        expect(parserRecords(src)).toEqual([]);
    });

    it('declare-inside-procedure: parser recovery ends the procedure at the Declare', () => {
        const src =
            'Public Sub S()\n' +
            '    Private Declare PtrSafe Sub Sleep Lib "kernel32" ()\n' +
            'End Sub\n';
        expect(legacyRecords(src)).toEqual([]);
        expect(sortedRecords(parserRecords(src))).toEqual([
            { kind: 'missing', line: 0, closer: 'End Sub' },
        ]);
    });
});
