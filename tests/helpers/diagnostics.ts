import { expect } from 'vitest';
import type { VbaDiagnostic } from '../../src/analyzer';

/** Returns the diagnostics whose code matches `code`. */
export function byCode(diags: readonly VbaDiagnostic[], code: string): VbaDiagnostic[] {
	return diags.filter((diag) => diag.code === code);
}

/** Resolves the source substring a diagnostic span covers. */
export function spanText(source: string, diag: VbaDiagnostic): string {
	return source.slice(diag.span.start, diag.span.end);
}

/**
 * Expected shape of one diagnostic. Fields left undefined are not asserted.
 * `message` is a substring check; exact wording is pinned once per rule in the
 * dedicated message-wording suite, not in behavior tests.
 */
export interface ExpectedDiagnostic {
	severity?: VbaDiagnostic['severity'];
	span?: string;
	message?: string | readonly string[];
}

/** Asserts `diags` contains exactly one diagnostic with `code`, matching `expected`. */
export function expectDiagnostic(
	source: string,
	diags: readonly VbaDiagnostic[],
	code: string,
	expected: ExpectedDiagnostic = {},
): VbaDiagnostic {
	return expectDiagnostics(source, diags, code, [expected])[0];
}

/**
 * Asserts the diagnostics with `code` match `expected` in order: rule code,
 * count, and each entry's severity / covered span / message substring.
 * Returns the matched diagnostics for further structured assertions.
 */
export function expectDiagnostics(
	source: string,
	diags: readonly VbaDiagnostic[],
	code: string,
	expected: readonly ExpectedDiagnostic[],
): VbaDiagnostic[] {
	const hits = byCode(diags, code);
	const actual = hits.map((hit) => ({
		code: hit.code,
		severity: hit.severity,
		span: spanText(source, hit),
	}));
	const wanted = expected.map((want, index) => ({
		code,
		severity: want.severity ?? hits[index]?.severity,
		span: want.span ?? actual[index]?.span,
	}));
	expect(actual).toEqual(wanted);
	expected.forEach((want, index) => {
		if (want.message === undefined) {
			return;
		}
		const parts = typeof want.message === 'string' ? [want.message] : want.message;
		for (const part of parts) {
			expect(hits[index].message).toContain(part);
		}
	});
	return hits;
}
