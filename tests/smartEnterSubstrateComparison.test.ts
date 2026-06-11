// Smart-Enter substrate comparison harness for audit #74.
//
// Smart Enter (src/vbaSmartEnter.ts, consumed by src/vbaTypingAutomation.ts
// and keyword completion) sits entirely on the legacy regex line scanner:
// stripVba + toLogicalLines from src/vbaSourceScan.ts. Audit #74 proposed
// rebuilding it on the analyzer lexer (src/analyzer/lexer/tokenize.ts) so a
// single lexer defines string/comment/continuation semantics.
//
// This harness compares the two lexical substrates -- legacy stripVba vs a
// lexer-derived equivalent (string-literal and comment token spans blanked,
// column alignment preserved) -- over every VBA sample in the repository,
// both per source (full lexer context) and per isolated line (the contract
// the Smart-Enter call sites actually use).
//
// Verdict recorded by this harness: the migration is UNBLOCKED. Across 1,103
// samples the substrates disagree on exactly one class of line:
//
//   rem-after-colon            (lexer correct) stripVba blanks only
//       whole-line Rem comments; the lexer recognizes Rem at any statement
//       start (MS-VBAL 3.3.5.2), so a trailing ": Rem ..." comment is blanked
//       by the lexer but leaks through stripVba. Migrating would be a strict
//       improvement here.
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
// With the only remaining divergence class a verified improvement, the audit
// #74 work rule (migrate only where the diff is empty or every difference is
// a verified improvement) no longer blocks rebuilding Smart Enter on the
// lexer substrate. src/vbaSmartEnter.ts still sits on legacy stripVba; the
// migration itself is tracked as follow-up work.

import { describe, expect, it } from 'vitest';
import { tokenize } from '../src/analyzer/lexer/tokenize';
import { stripVba } from '../src/vbaSourceScan';
import { allStructuralComparisonSamples } from './helpers/structuralCorpus';

/**
 * The lexer-derived substrate under evaluation: every physical line of
 * `source` with string-literal and comment token spans blanked to spaces,
 * preserving length and column alignment (the stripVba contract). Comments
 * and string literals never span physical lines (MS-VBAL 3.3.1 / 3.3.4), so
 * blanking on the token's start line covers the whole token.
 */
function lexerStrippedLines(source: string): string[] {
    const chars = source.split(/\r\n|\r|\n/).map((line) => line.split(''));
    for (const token of tokenize(source)) {
        if (token.kind !== 'comment' && token.kind !== 'stringLiteral') {
            continue;
        }
        const lineChars = chars[token.line];
        if (!lineChars) {
            continue;
        }
        const end = Math.min(token.character + token.rawText.length, lineChars.length);
        for (let col = token.character; col < end; col++) {
            lineChars[col] = ' ';
        }
    }
    return chars.map((lineChars) => lineChars.join(''));
}

/** Isolated-line variant matching the per-line stripVba call sites. */
function lexerStrippedLine(line: string): string {
    return lexerStrippedLines(line)[0] ?? line;
}

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

describe('smart-enter substrate comparison (audit #74)', () => {
    const samples = allStructuralComparisonSamples();

    it('classifies every stripping divergence between regex and lexer substrates', () => {
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
            `Unexplained stripping divergence between stripVba and the analyzer ` +
            `lexer:\n${unclassified.join('\n')}`,
        ).toEqual([]);
    });

    // -- Pinned minimal repros: the surviving divergence class, plus a
    // regression pin for the fixed file-number/date-literal lexer bug -------

    it('rem-after-colon: only the lexer blanks a trailing ": Rem ..." comment', () => {
        const line = '    Debug.Print 1: Rem hidden: End If';
        expect(stripVba(line)).toBe('    Debug.Print 1: Rem hidden: End If');
        expect(lexerStrippedLine(line)).toBe('    Debug.Print 1:                   ');
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
