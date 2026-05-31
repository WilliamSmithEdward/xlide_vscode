// Diagnostic rule catalog (MS-VBAL Phase 5: Active Diagnostics).
//
// Every active diagnostic the analyzer can raise is described here once: a
// stable `code` (shown in the Problems panel and usable in `// xlide-disable`
// style tooling later), a human title, a default severity, the MS-VBAL section
// it enforces, and a confidence level. Only *high-confidence*, deterministic
// rules are enabled - rules that would need an expression-level binder or a
// complete host catalogue to avoid false positives (undeclared-variable,
// unknown-call) are deliberately NOT implemented here. See the roadmap.
//
// Pure data: no `vscode` dependency. The VS Code layer maps `severity` onto
// `vscode.DiagnosticSeverity` and `code`/`source` onto the diagnostic.

/** Severity of a diagnostic, independent of the VS Code enum. */
export type DiagnosticSeverity = 'error' | 'warning' | 'information' | 'hint';

/** Static description of a single diagnostic rule. */
export interface DiagnosticRuleMetadata {
	/** Stable identifier shown as the diagnostic code. */
	code: string;
	/** Short human-readable title. */
	title: string;
	/** Severity used unless the user overrides it. */
	defaultSeverity: DiagnosticSeverity;
	/** Always 'XLIDE' - the diagnostic source label. */
	source: 'XLIDE';
	/** MS-VBAL section (or other authority) the rule enforces. */
	specReference?: string;
	/** True when the rule needs the whole project, not a single module. */
	requiresWholeProject?: boolean;
	/** How certain the rule is that a flagged construct is genuinely wrong. */
	confidence: 'high' | 'medium' | 'low';
}

/** The catalogue of active diagnostic rules, keyed by a stable rule name. */
export const DIAGNOSTIC_RULES = {
	unterminatedString: {
		code: 'unterminated-string',
		title: 'Unterminated string literal',
		defaultSeverity: 'error',
		source: 'XLIDE',
		specReference: 'MS-VBAL 3.3.4',
		confidence: 'high',
	},
	duplicateProcedure: {
		code: 'duplicate-procedure',
		title: 'Ambiguous (duplicate) procedure name in module',
		defaultSeverity: 'error',
		source: 'XLIDE',
		specReference: 'MS-VBAL 5.3',
		confidence: 'high',
	},
	duplicateDeclaration: {
		code: 'duplicate-declaration',
		title: 'Duplicate declaration in the current scope',
		defaultSeverity: 'error',
		source: 'XLIDE',
		specReference: 'MS-VBAL 5.2 / 5.3',
		confidence: 'high',
	},
	duplicateModuleMember: {
		code: 'duplicate-module-variable',
		title: 'Duplicate module-level declaration',
		defaultSeverity: 'error',
		source: 'XLIDE',
		specReference: 'MS-VBAL 5.2.3',
		confidence: 'high',
	},
	constAssignment: {
		code: 'const-assignment',
		title: 'Assignment to a constant',
		defaultSeverity: 'error',
		source: 'XLIDE',
		specReference: 'MS-VBAL 5.4.3.1',
		confidence: 'high',
	},
	optionExplicitMissing: {
		code: 'option-explicit-missing',
		title: 'Option Explicit is not specified',
		defaultSeverity: 'warning',
		source: 'XLIDE',
		specReference: 'MS-VBAL 5.2.4.1.1',
		confidence: 'high',
	},
} satisfies Record<string, DiagnosticRuleMetadata>;

/** Stable rule-name keys of {@link DIAGNOSTIC_RULES}. */
export type DiagnosticRuleName = keyof typeof DIAGNOSTIC_RULES;
