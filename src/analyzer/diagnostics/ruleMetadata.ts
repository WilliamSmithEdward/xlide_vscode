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
// The one cross-module rule that is enabled - `unknownCallStatement` - fires
// only on a *call statement* whose callee is a bare (non-member) identifier (the
// lone-identifier `DoStartup`, the parenless `MsgBox "hi"`, or `Call Foo`) whose
// name resolves to no procedure anywhere in the project, no VBA runtime
// function/statement, no host global or Application member, and no in-scope
// declaration. That is the unambiguous VBE "Sub or Function not defined"
// compile error, so it stays high-confidence and low-noise. The broader
// undeclared-variable / arbitrary-expression unknown-call cases still need a
// full expression binder and remain unimplemented.
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
	invalidProcedureHeader: {
		code: 'invalid-proc-header',
		title: 'Invalid procedure declaration',
		defaultSeverity: 'error',
		source: 'XLIDE',
		specReference: 'MS-VBAL 5.3.1',
		confidence: 'high',
	},
	unbalancedParens: {
		code: 'unbalanced-parens',
		title: 'Unbalanced parentheses',
		defaultSeverity: 'error',
		source: 'XLIDE',
		specReference: 'MS-VBAL 3.3.1',
		confidence: 'high',
	},
	argumentCount: {
		code: 'argument-count',
		title: 'Wrong number of arguments',
		defaultSeverity: 'error',
		source: 'XLIDE',
		specReference: 'MS-VBAL 5.4.2.1',
		confidence: 'high',
	},
	unknownCallStatement: {
		code: 'unknown-call',
		title: 'Sub or Function not defined',
		defaultSeverity: 'error',
		source: 'XLIDE',
		specReference: 'MS-VBAL 5.4.2.1',
		requiresWholeProject: true,
		confidence: 'high',
	},
	dimInitializer: {
		code: 'dim-initializer',
		title: 'Declaration cannot include an initializer (VB.NET syntax)',
		defaultSeverity: 'error',
		source: 'XLIDE',
		specReference: 'MS-VBAL 5.2.3.1',
		confidence: 'high',
	},
	callRequiresParens: {
		code: 'call-requires-parens',
		title: 'Call statement requires parentheses around arguments',
		defaultSeverity: 'error',
		source: 'XLIDE',
		specReference: 'MS-VBAL 5.4.2.1',
		confidence: 'high',
	},
	requiredParamAfterOptional: {
		code: 'required-param-after-optional',
		title: 'A required parameter cannot follow an Optional parameter',
		defaultSeverity: 'error',
		source: 'XLIDE',
		specReference: 'MS-VBAL 5.3.1.5',
		confidence: 'high',
	},
	paramArrayNotLast: {
		code: 'paramarray-not-last',
		title: 'ParamArray must be the final parameter',
		defaultSeverity: 'error',
		source: 'XLIDE',
		specReference: 'MS-VBAL 5.3.1.6',
		confidence: 'high',
	},
	exitWrongProcedure: {
		code: 'exit-wrong-proc',
		title: 'Exit statement does not match the enclosing procedure',
		defaultSeverity: 'error',
		source: 'XLIDE',
		specReference: 'MS-VBAL 5.4.1.3',
		confidence: 'high',
	},
	optionAfterDeclaration: {
		code: 'option-after-declaration',
		title: 'Option statement must precede all declarations',
		defaultSeverity: 'error',
		source: 'XLIDE',
		specReference: 'MS-VBAL 5.2.1',
		confidence: 'high',
	},
} satisfies Record<string, DiagnosticRuleMetadata>;

/** Stable rule-name keys of {@link DIAGNOSTIC_RULES}. */
export type DiagnosticRuleName = keyof typeof DIAGNOSTIC_RULES;
