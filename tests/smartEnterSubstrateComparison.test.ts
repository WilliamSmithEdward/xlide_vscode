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
// Verdict recorded by this harness: the migration is BLOCKED. Across 1,103
// samples the substrates disagree on exactly two classes of line:
//
//   rem-after-colon            (lexer correct) stripVba blanks only
//       whole-line Rem comments; the lexer recognizes Rem at any statement
//       start (MS-VBAL 3.3.5.2), so a trailing ": Rem ..." comment is blanked
//       by the lexer but leaks through stripVba. Migrating would be a strict
//       improvement here.
//
//   file-number-date-literal   (legacy correct -- BLOCKS the migration) In
//       `Write #ff, "csv", 1, #1/1/2026#` the lexer's date-literal scan
//       (tokenize.ts, '#' branch) treats the file-number '#' as the start of
//       a date literal and swallows `#ff, "csv", 1, #` as one dateLiteral
//       token, so the enclosed string literal is NOT blanked. stripVba blanks
//       it correctly. Under the lexer substrate, a string such as
//       `Write #1, "x: End If"` would leak a fake block closer into the
//       Smart-Enter grammar -- a regression, not an improvement.
//
// Because one divergence class is a lexer regression, Smart Enter keeps the
// legacy substrate (audit #74 work rule: migrate only where the diff is empty
// or every difference is a verified improvement; do not tweak either engine
// to force equivalence). When the lexer learns to distinguish file numbers
// from date literals, the 'file-number-date-literal' repro below will fail --
// that is the signal that this migration is unblocked.

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
    | 'file-number-date-literal'
    | 'UNCLASSIFIED';

const MIDLINE_REM_RE = /:\s*Rem\b/i;
// A '#' file-number argument later mistaken for a date-literal opener: the
// statement carries `#<name-or-number>` and a second '#' later on the line.
const FILE_NUMBER_RE = /#\s*[A-Za-z_0-9]/;

function classifyLine(raw: string, legacy: string, lexer: string): SubstrateDivergenceClass {
    if (MIDLINE_REM_RE.test(raw)) {
        return 'rem-after-colon';
    }
    // The lexer under-blanks (keeps text legacy blanked) when a file-number
    // '#' swallowed the rest of the statement as a date literal.
    const lexerKeptWhatLegacyBlanked = [...raw].some(
        (_, i) => legacy[i] === ' ' && lexer[i] !== ' ',
    );
    if (
        lexerKeptWhatLegacyBlanked &&
        FILE_NUMBER_RE.test(raw) &&
        raw.split('#').length > 2
    ) {
        return 'file-number-date-literal';
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
                    const cls = classifyLine(rawLines[i], legacy, lexer);
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

    // -- Pinned minimal repros, one per divergence class --------------------

    it('rem-after-colon: only the lexer blanks a trailing ": Rem ..." comment', () => {
        const line = '    Debug.Print 1: Rem hidden: End If';
        expect(stripVba(line)).toBe('    Debug.Print 1: Rem hidden: End If');
        expect(lexerStrippedLine(line)).toBe('    Debug.Print 1:                   ');
    });

    it('file-number-date-literal: the lexer swallows a Write # statement as a date literal (blocks migration)', () => {
        const line = '    Write #ff, "csv", 1, #1/1/2026#';
        // Legacy strips the string correctly.
        expect(stripVba(line)).toBe('    Write #ff,      , 1, #1/1/2026#');
        // The lexer leaves the string un-blanked: `#ff, "csv", 1, #` lexed as
        // one dateLiteral token. When this assertion fails, the lexer has
        // learned file numbers and the Smart-Enter migration is unblocked.
        expect(lexerStrippedLine(line)).toBe('    Write #ff, "csv", 1, #1/1/2026#');
        const kinds = tokenize(line).map((t) => `${t.kind}:${t.rawText}`);
        expect(kinds).toContain('dateLiteral:#ff, "csv", 1, #');
    });
});
