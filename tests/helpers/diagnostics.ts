import type { VbaDiagnostic } from '../../src/analyzer';

/** Returns the diagnostics whose code matches `code`. */
export function byCode(diags: readonly VbaDiagnostic[], code: string): VbaDiagnostic[] {
	return diags.filter((diag) => diag.code === code);
}

/** Resolves the source substring a diagnostic span covers. */
export function spanText(source: string, diag: VbaDiagnostic): string {
	return source.slice(diag.span.start, diag.span.end);
}
