// Type-suffixed numeric literal overflow (MS-VBAL number-token grammar).
//
// A VBA integer literal may carry the `%` type-declaration suffix, which forces
// the literal to Integer. When the literal's magnitude exceeds the Integer range,
// VBE rejects the token at COMPILE time as a "Syntax error" -- the suffix is
// evaluated against the literal before any surrounding context, so the error is
// intrinsic to the token and independent of where it appears.
//
// Oracle evidence (suffix_integer_pct_*_compile):
//   - 40000% / 32768% / -32768% -> rejected (compile); 32767% -> accepted.
// The negated-min probe (-32768%) is ALSO rejected, proving VBE rejects the
// over-range suffixed token regardless of a leading sign. So the check is purely
// magnitude-vs-Integer-max with no sign special-casing.
//
// Why only `%` (and NOT the `&` Long suffix): `%` is unambiguously a type suffix,
// but `&` is BOTH the Long suffix AND the string-concatenation operator. When an
// over-range digit run is glued to `&` followed by an operand, VBE re-parses the
// `&` as concatenation rather than a (overflowing) Long suffix and ACCEPTS it --
// oracle suffix_long_amp_glued_concat_accepted: `s = 3000000000&"x"` is accepted
// as `3000000000 & "x"`. The lexer greedily glues the trailing `&` into the
// number token, so flagging `&`-suffixed overflow would false-positive on that
// valid concat form. Per the no-false-positive rule, `&` overflow is deferred
// (a known, conservative false negative) until the suffix-vs-concat boundary is
// fully oracle-mapped by next-token context.
//
// No-false-positive discipline:
//   - Only `integerLiteral` tokens that are pure decimal digits followed by a
//     single `%` are considered (`/^\d+%$/`). Hex (`&HFFFF`), octal (`&O777`),
//     `&`/`^` suffixes, and float-suffix (`!`/`#`/`@`) literals never match.
//   - Bounds come from the shared `numericLiteralBounds('integer')`, so this rule
//     and the coercion-overflow rule can never disagree on the Integer range.
//   - Provably-inactive conditional-compilation regions are skipped.

import type { ConditionalActivityTracker } from '../../conditional/conditionalCompilation';
import type { PushFn } from '../analysisContext';
import { numericLiteralBounds } from '../typeInference';
import { tokenizeCached } from '../../lexer/tokenize';

/** A pure-decimal integer literal with the `%` (Integer) type suffix. */
const INTEGER_SUFFIXED_LITERAL = /^(\d+)%$/;

export function checkSuffixedLiteralOverflow(
	source: string,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	const bounds = numericLiteralBounds('integer');
	if (!bounds) {
		return;
	}
	const tokens = tokenizeCached(source);
	for (let index = 0; index < tokens.length; index++) {
		const tok = tokens[index];
		if (tok.kind === 'floatLiteral') {
			const span = { start: tok.start, end: tok.end };
			if (activity?.isInactive(span)) {
				continue;
			}
			// `1E400` is past the Double range and `1.5%` puts the Integer
			// suffix on a fractional literal: both are compile-time "Syntax
			// error" (issue #125, measured in Excel 16.0).
			const value = Number(tok.rawText.replace(/[dD]/g, 'E').replace(/[!#@]$/, ''));
			if (!Number.isFinite(value)) {
				push(
					'floatLiteralOverflow',
					`The literal '${tok.rawText}' is outside the Double range (about 1.8E+308). VBE rejects this at compile time as a Syntax error.`,
					span,
				);
				continue;
			}
			const next = tokens[index + 1];
			if (next && next.kind === 'unknown' && next.rawText === '%' && next.start === tok.end) {
				push(
					'suffixedLiteralOverflow',
					`The literal '${tok.rawText}%' puts the '%' Integer type suffix on a fractional number, which has no Integer form. VBE rejects this at compile time as a Syntax error.`,
					{ start: tok.start, end: next.end },
				);
			}
			continue;
		}
		if (tok.kind !== 'integerLiteral') {
			continue;
		}
		const match = INTEGER_SUFFIXED_LITERAL.exec(tok.rawText);
		if (!match) {
			continue;
		}
		// The digit run is unsigned, so only the upper bound is reachable. A value
		// beyond Number.isSafeInteger is still definitively over the Integer max.
		if (Number(match[1]) <= bounds.max) {
			continue;
		}
		const span = { start: tok.start, end: tok.end };
		if (activity?.isInactive(span)) {
			continue;
		}
		push(
			'suffixedLiteralOverflow',
			`The literal '${tok.rawText}' is outside the ${bounds.label} range ${bounds.min} to ${bounds.max} of its '%' type suffix. VBE rejects this at compile time as a Syntax error.`,
			span,
		);
	}
}
