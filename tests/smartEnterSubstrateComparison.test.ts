// Smart-Enter substrate regression gate for audit #74.
//
// Smart Enter (src/vbaSmartEnter.ts, consumed by src/vbaTypingAutomation.ts
// and keyword completion) historically sat on the legacy regex line scanner
// (stripVba + toLogicalLines from src/vbaSourceScan.ts). This harness was
// built to compare that scanner against a lexer-derived equivalent and decide
// whether rebuilding Smart Enter on the analyzer lexer was safe. The recorded
// verdict was UNBLOCKED -- across 1,103 repository samples the substrates
// disagreed on exactly one line class, and the lexer was strictly more
// correct -- and the migration has since landed: Smart Enter's production
// substrate IS the lexer substrate (lexerStrippedLines / lexerStrippedLine in
// src/analyzer/lexer/strippedLines.ts), asserted directly below.
//
// The harness now serves as the regression gate for the migrated substrate.
// It re-diffs the production substrate against legacy stripVba over every VBA
// sample in the repository, both per source (full lexer context) and per
// isolated line (the per-line call-site contract), and requires that every
// difference falls in the single deliberate divergence class:
//
//   rem-after-colon            (lexer correct; pinned below as EXPECTED
//       production behavior) stripVba blanks only whole-line Rem comments;
//       the lexer recognizes Rem at any statement start (MS-VBAL 3.3.5.2), so
//       a trailing ": Rem ..." comment is blanked by the production substrate
//       but leaks through stripVba. This divergence is a deliberate
//       improvement, not drift; any other difference fails as UNCLASSIFIED.
//
// A second class -- file-number-date-literal, where the lexer's date-literal
// scan treated the file-number '#' of `Write #ff, "csv", 1, #1/1/2026#` as a
// date-literal opener and swallowed `#ff, "csv", 1, #` as one dateLiteral
// token, leaving the enclosed string un-blanked -- was a verified lexer bug.
// The lexer now validates the text between a '#' pair against the MS-VBAL
// 3.3.3 date-or-time grammar (tokenize.ts, isDateLiteralBody) and lexes a
// non-matching '#' as an operator, so both substrates blank such lines
// identically. The pinned repro below asserts the corrected lexing; if the
// divergence ever reappears, this harness fails it as UNCLASSIFIED.
//
// The structural block-balance diagnostics (src/vbaStructuralDiagnostics.ts)
// intentionally remain on the legacy stripVba/toLogicalLines scanner: that
// migration is still blocked by its own divergence classes.

import { describe, expect, it } from 'vitest';
import { lexerStrippedLine, lexerStrippedLines } from '../src/analyzer/lexer/strippedLines';
import { tokenize } from '../src/analyzer/lexer/tokenize';
import { openSmartBlockClosersBefore } from '../src/vbaSmartEnter';
import { stripVba } from '../src/vbaSourceScan';
import { allStructuralComparisonSamples } from './helpers/structuralCorpus';

type SubstrateDivergenceClass =
    | 'rem-after-colon'
    | 'UNCLASSIFIED';

const MIDLINE_REM_RE = /:\s*Rem\b/i;

function classifyLine(raw: string): SubstrateDivergenceClass {
    if (MIDLINE_REM_RE.test(raw)) {
        return 'rem-after-colon';
    }
    return 'UNCLASSIFIED';
}

describe('smart-enter substrate regression gate (audit #74)', () => {
    const samples = allStructuralComparisonSamples();

    it('every divergence from the legacy scanner is the pinned rem-after-colon improvement', () => {
        const counts = new Map<SubstrateDivergenceClass, number>();
        const unclassified: string[] = [];

        for (const sample of samples) {
            const rawLines = sample.source.split(/\r\n|\r|\n/);
            const legacyLines = rawLines.map(stripVba);
            const lexerLines = lexerStrippedLines(sample.source);
            expect(lexerLines).toHaveLength(legacyLines.length);

            for (let i = 0; i < rawLines.length; i++) {
                const pairs: [string, string][] = [
                    [legacyLines[i], lexerLines[i]],
                    // The Smart-Enter call sites also strip single lines in
                    // isolation; compare that contract too.
                    [stripVba(rawLines[i]), lexerStrippedLine(rawLines[i])],
                ];
                for (const [legacy, lexer] of pairs) {
                    if (legacy === lexer) {
                        continue;
                    }
                    const cls = classifyLine(rawLines[i]);
                    counts.set(cls, (counts.get(cls) ?? 0) + 1);
                    if (cls === 'UNCLASSIFIED') {
                        unclassified.push(
                            `${sample.id}@${i}: raw=${JSON.stringify(rawLines[i])} ` +
                            `legacy=${JSON.stringify(legacy)} lexer=${JSON.stringify(lexer)}`,
                        );
                    }
                    break; // count each physical line once
                }
            }
        }

        // eslint-disable-next-line no-console
        console.log(
            `[audit #74] substrate comparison: samples=${samples.length} ` +
            `divergentLines=${JSON.stringify(Object.fromEntries(counts))}`,
        );

        expect(
            unclassified,
            `Unexplained stripping divergence between stripVba and the production ` +
            `lexer substrate:\n${unclassified.join('\n')}`,
        ).toEqual([]);

        // The corpus must keep exercising the deliberate divergence; if this
        // drops to zero the gate has lost its rem-after-colon coverage.
        expect(
            counts.get('rem-after-colon') ?? 0,
            'corpus no longer contains a divergent ": Rem ..." line',
        ).toBeGreaterThan(0);
    });

    // -- Pinned minimal repros: the deliberate divergence class, plus a
    // regression pin for the fixed file-number/date-literal lexer bug -------

    it('rem-after-colon: the production substrate blanks a trailing ": Rem ..." comment that stripVba leaks', () => {
        // Deliberate divergence from stripVba (lexer correct, MS-VBAL
        // 3.3.5.2): Rem is a comment at any statement start, including after
        // a ':' separator, so the trailing comment -- and the ": End If"
        // INSIDE it -- must be blanked. stripVba only handles whole-line Rem
        // comments and leaks the text, which made downstream colon-splitting
        // see a phantom `End If` statement.
        const line = '    Debug.Print 1: Rem hidden: End If';
        expect(stripVba(line)).toBe('    Debug.Print 1: Rem hidden: End If');
        expect(lexerStrippedLine(line)).toBe('    Debug.Print 1:                   ');
    });

    it('rem-after-colon: Smart Enter consumes the lexer substrate, so a ": Rem ... End If" tail no longer closes the block', () => {
        // End-to-end pin that the production Smart-Enter path sits on the
        // lexer substrate: under legacy stripVba the leaked comment text was
        // colon-split into a phantom `End If` logical line that incorrectly
        // closed the open If block.
        const src = 'Sub T()\n    If ready Then\n        Debug.Print 1: Rem hidden: End If\n        \n';
        expect(openSmartBlockClosersBefore(src, src.length)).toEqual(['End Sub', 'End If']);
    });

    it('file-number-date-literal: the lexer lexes a Write # file number as an operator, not a date-literal opener', () => {
        const line = '    Write #ff, "csv", 1, #1/1/2026#';
        // Both substrates blank the string and keep the genuine date literal.
        expect(stripVba(line)).toBe('    Write #ff,      , 1, #1/1/2026#');
        expect(lexerStrippedLine(line)).toBe('    Write #ff,      , 1, #1/1/2026#');
        const kinds = tokenize(line).map((t) => `${t.kind}:${t.rawText}`);
        expect(kinds).toContain('operator:#');
        expect(kinds).toContain('stringLiteral:"csv"');
        expect(kinds).toContain('dateLiteral:#1/1/2026#');
        expect(kinds).not.toContain('dateLiteral:#ff, "csv", 1, #');
    });
});
