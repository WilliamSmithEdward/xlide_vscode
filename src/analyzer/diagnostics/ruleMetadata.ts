// Diagnostic rule catalog (MS-VBAL Phase 5: Active Diagnostics).
//
// Every active diagnostic the analyzer can raise is described here once: a
// stable `code` (shown in the Problems panel and usable in `// xlide-disable`
// style tooling later), a human title, a default severity, category, VBE compile
// equivalence, the MS-VBAL section it enforces, and a confidence level. Only
// *high-confidence*, deterministic rules are enabled - rules that would need an
// expression-level binder or a complete host catalogue to avoid false positives
// remain deliberately out of scope until the binder can prove them.
//
// The cross-module call rule that is enabled - `unknownCallStatement` - fires
// only on a *call statement* whose callee is a bare (non-member) identifier (the
// lone-identifier `DoStartup`, the parenless `MsgBox "hi"`, or `Call Foo`) whose
// name resolves to no procedure anywhere in the project, no VBA runtime
// function/statement, no host global or Application member, and no in-scope
// declaration. The `undeclaredVariable` rule follows the same shape: it only
// fires when project-visible identifiers are available, and it scans
// high-confidence write/read positions while skipping type-name, label,
// named-argument, and unresolved external-call positions.
//
// Pure data: no `vscode` dependency. The VS Code layer maps `severity` onto
// `vscode.DiagnosticSeverity` and `code`/`source` onto the diagnostic.

/** Severity of a diagnostic, independent of the VS Code enum. */
export type DiagnosticSeverity = 'error' | 'warning' | 'information' | 'hint';

/** Broad purpose bucket used by tests, docs, and future filtering. */
export type DiagnosticCategory =
	| 'syntax'
	| 'lexer'
	| 'parser'
	| 'realtime-recovery'
	| 'declaration'
	| 'semantic'
	| 'project-symbol'
	| 'module-kind'
	| 'excel-host'
	| 'style';

/** Why a rule is surfaced at its default severity. */
export type DiagnosticEvidenceKind =
	| 'compile-error'
	| 'deterministic-runtime-error'
	| 'runtime-risk'
	| 'style-policy';

/** Static description of a single diagnostic rule. */
export interface DiagnosticRuleMetadata {
	/** Stable identifier shown as the diagnostic code. */
	code: string;
	/** Short human-readable title. */
	title: string;
	/** Severity used unless the user overrides it. */
	defaultSeverity: DiagnosticSeverity;
	/** Broad purpose bucket for the diagnostic. */
	category: DiagnosticCategory;
	/** True when the diagnostic is expected to match a VBE compile failure. */
	vbeCompileEquivalent: boolean;
	/** Distinguishes compile-time errors from deterministic runtime errors. */
	diagnosticKind: DiagnosticEvidenceKind;
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
		category: 'syntax',
		vbeCompileEquivalent: true,
		diagnosticKind: 'compile-error',
		source: 'XLIDE',
		specReference: 'MS-VBAL 3.3.4',
		confidence: 'high',
	},
	duplicateProcedure: {
		code: 'duplicate-procedure',
		title: 'Ambiguous (duplicate) procedure name in module',
		defaultSeverity: 'error',
		category: 'declaration',
		vbeCompileEquivalent: true,
		diagnosticKind: 'compile-error',
		source: 'XLIDE',
		specReference: 'MS-VBAL 5.3',
		confidence: 'high',
	},
	duplicateDeclaration: {
		code: 'duplicate-declaration',
		title: 'Duplicate declaration in the current scope',
		defaultSeverity: 'error',
		category: 'declaration',
		vbeCompileEquivalent: true,
		diagnosticKind: 'compile-error',
		source: 'XLIDE',
		specReference: 'MS-VBAL 5.2 / 5.3',
		confidence: 'high',
	},
	duplicateModuleMember: {
		code: 'duplicate-module-variable',
		title: 'Duplicate module-level declaration',
		defaultSeverity: 'error',
		category: 'declaration',
		vbeCompileEquivalent: true,
		diagnosticKind: 'compile-error',
		source: 'XLIDE',
		specReference: 'MS-VBAL 5.2.3',
		confidence: 'high',
	},
	constAssignment: {
		code: 'const-assignment',
		title: 'Assignment to a constant',
		defaultSeverity: 'error',
		category: 'semantic',
		vbeCompileEquivalent: true,
		diagnosticKind: 'compile-error',
		source: 'XLIDE',
		specReference: 'MS-VBAL 5.4.3.1',
		confidence: 'high',
	},
	optionExplicitMissing: {
		code: 'option-explicit-missing',
		title: 'Option Explicit is not specified',
		defaultSeverity: 'warning',
		category: 'style',
		vbeCompileEquivalent: false,
		diagnosticKind: 'style-policy',
		source: 'XLIDE',
		specReference: 'MS-VBAL 5.2.4.1.1',
		confidence: 'high',
	},
	undeclaredVariable: {
		code: 'undeclared-variable',
		title: 'Variable not defined',
		defaultSeverity: 'error',
		category: 'project-symbol',
		vbeCompileEquivalent: true,
		diagnosticKind: 'compile-error',
		source: 'XLIDE',
		specReference: 'MS-VBAL 5.2.4.1.1',
		requiresWholeProject: true,
		confidence: 'high',
	},
	invalidProcedureHeader: {
		code: 'invalid-proc-header',
		title: 'Invalid procedure declaration',
		defaultSeverity: 'error',
		category: 'declaration',
		vbeCompileEquivalent: true,
		diagnosticKind: 'compile-error',
		source: 'XLIDE',
		specReference: 'MS-VBAL 5.3.1',
		confidence: 'high',
	},
	invalidDeclarationName: {
		code: 'invalid-declaration-name',
		title: 'Reserved keyword used as a declaration name',
		defaultSeverity: 'error',
		category: 'declaration',
		vbeCompileEquivalent: true,
		diagnosticKind: 'compile-error',
		source: 'XLIDE',
		specReference: 'MS-VBAL 3.3.5.2',
		confidence: 'high',
	},
	unbalancedParens: {
		code: 'unbalanced-parens',
		title: 'Unbalanced parentheses',
		defaultSeverity: 'error',
		category: 'syntax',
		vbeCompileEquivalent: true,
		diagnosticKind: 'compile-error',
		source: 'XLIDE',
		specReference: 'MS-VBAL 3.3.1',
		confidence: 'high',
	},
	argumentCount: {
		code: 'argument-count',
		title: 'Wrong number of arguments',
		defaultSeverity: 'error',
		category: 'semantic',
		vbeCompileEquivalent: true,
		diagnosticKind: 'compile-error',
		source: 'XLIDE',
		specReference: 'MS-VBAL 5.4.2.1',
		confidence: 'high',
	},
	argumentTypeMismatch: {
		code: 'argument-type-mismatch',
		title: 'Argument type mismatch',
		defaultSeverity: 'error',
		category: 'semantic',
		vbeCompileEquivalent: false,
		diagnosticKind: 'deterministic-runtime-error',
		source: 'XLIDE',
		specReference: 'MS-VBAL 5.3.1 / runtime type coercion',
		confidence: 'high',
	},
	argumentObjectTypeMismatch: {
		code: 'argument-object-type-mismatch',
		title: 'Object argument type mismatch',
		defaultSeverity: 'error',
		category: 'semantic',
		vbeCompileEquivalent: true,
		diagnosticKind: 'compile-error',
		source: 'XLIDE',
		specReference: 'MS-VBAL 5.3.1',
		confidence: 'high',
	},
	assignmentTypeMismatch: {
		code: 'assignment-type-mismatch',
		title: 'Assignment type mismatch',
		defaultSeverity: 'error',
		category: 'semantic',
		vbeCompileEquivalent: false,
		diagnosticKind: 'deterministic-runtime-error',
		source: 'XLIDE',
		specReference: 'MS-VBAL 5.4.3 / runtime type coercion',
		confidence: 'high',
	},
	missingReturnAssignment: {
		code: 'missing-return-assignment',
		title: 'Function has no return assignment',
		defaultSeverity: 'warning',
		category: 'semantic',
		vbeCompileEquivalent: false,
		diagnosticKind: 'runtime-risk',
		source: 'XLIDE',
		specReference: 'VBA Function return variable semantics',
		confidence: 'high',
	},
	assignmentObjectTypeMismatch: {
		code: 'assignment-object-type-mismatch',
		title: 'Object assignment type mismatch',
		defaultSeverity: 'error',
		category: 'semantic',
		vbeCompileEquivalent: true,
		diagnosticKind: 'compile-error',
		source: 'XLIDE',
		specReference: 'MS-VBAL 5.4.3 / Set statement',
		confidence: 'high',
	},
	readonlyMemberAssignment: {
		code: 'readonly-member-assignment',
		title: 'Assignment to a read-only member',
		defaultSeverity: 'error',
		category: 'semantic',
		vbeCompileEquivalent: true,
		diagnosticKind: 'compile-error',
		source: 'XLIDE',
		specReference: "VBE oracle: Can't assign to read-only property",
		confidence: 'high',
	},
	setRequired: {
		code: 'set-required',
		title: 'Object assignment requires Set',
		defaultSeverity: 'error',
		category: 'semantic',
		vbeCompileEquivalent: true,
		diagnosticKind: 'compile-error',
		source: 'XLIDE',
		specReference: 'MS-VBAL 5.4.3 / Set statement',
		confidence: 'high',
	},
	memberNotFound: {
		code: 'member-not-found',
		title: 'Object member not found',
		defaultSeverity: 'error',
		category: 'semantic',
		vbeCompileEquivalent: true,
		diagnosticKind: 'compile-error',
		source: 'XLIDE',
		specReference: 'VBE oracle: Method or data member not found',
		requiresWholeProject: true,
		confidence: 'high',
	},
	stringArithmeticCoercion: {
		code: 'string-arithmetic-coercion',
		title: 'Nonnumeric string in numeric expression',
		defaultSeverity: 'error',
		category: 'semantic',
		vbeCompileEquivalent: false,
		diagnosticKind: 'deterministic-runtime-error',
		source: 'XLIDE',
		specReference: 'MS-VBAL 5.6 / runtime type coercion',
		confidence: 'high',
	},
	unknownCallStatement: {
		code: 'unknown-call',
		title: 'Sub or Function not defined',
		defaultSeverity: 'error',
		category: 'project-symbol',
		vbeCompileEquivalent: true,
		diagnosticKind: 'compile-error',
		source: 'XLIDE',
		specReference: 'MS-VBAL 5.4.2.1',
		requiresWholeProject: true,
		confidence: 'high',
	},
	nonCallableCallStatement: {
		code: 'non-callable-call',
		title: 'Identifier is not callable',
		defaultSeverity: 'error',
		category: 'semantic',
		vbeCompileEquivalent: true,
		diagnosticKind: 'compile-error',
		source: 'XLIDE',
		specReference: 'MS-VBAL 5.4.2.1',
		confidence: 'high',
	},
	dimInitializer: {
		code: 'dim-initializer',
		title: 'Declaration cannot include an initializer (VB.NET syntax)',
		defaultSeverity: 'error',
		category: 'declaration',
		vbeCompileEquivalent: true,
		diagnosticKind: 'compile-error',
		source: 'XLIDE',
		specReference: 'MS-VBAL 5.2.3.1',
		confidence: 'high',
	},
	unexpectedDeclarationToken: {
		code: 'unexpected-declaration-token',
		title: 'Unexpected token after declaration type',
		defaultSeverity: 'error',
		category: 'declaration',
		vbeCompileEquivalent: true,
		diagnosticKind: 'compile-error',
		source: 'XLIDE',
		specReference: 'MS-VBAL 5.2.3.1 / VBE oracle: Syntax error',
		confidence: 'high',
	},
	objectModulePublicMember: {
		code: 'object-module-public-member',
		title: 'Invalid public member in object module',
		defaultSeverity: 'error',
		category: 'declaration',
		vbeCompileEquivalent: true,
		diagnosticKind: 'compile-error',
		source: 'XLIDE',
		specReference: 'VBE oracle: public object-module member restrictions',
		confidence: 'high',
	},
	eventHandlerWrongModule: {
		code: 'event-handler-module-scope',
		title: 'Event handler is not wired in this module',
		defaultSeverity: 'information',
		category: 'module-kind',
		vbeCompileEquivalent: false,
		diagnosticKind: 'style-policy',
		source: 'XLIDE',
		specReference: 'Excel document-module event binding',
		confidence: 'high',
	},
	invalidAsTypeName: {
		code: 'invalid-as-type-name',
		title: 'Invalid type name',
		defaultSeverity: 'error',
		category: 'declaration',
		vbeCompileEquivalent: true,
		diagnosticKind: 'compile-error',
		source: 'XLIDE',
		specReference: 'MS-VBAL 3.3.5.2 / 5.2.3.1',
		confidence: 'high',
	},
	setRequiresObject: {
		code: 'set-requires-object',
		title: 'Set assignment requires an object variable',
		defaultSeverity: 'error',
		category: 'semantic',
		vbeCompileEquivalent: true,
		diagnosticKind: 'compile-error',
		source: 'XLIDE',
		specReference: 'MS-VBAL 5.4.3',
		confidence: 'high',
	},
	scalarMemberAccess: {
		code: 'scalar-member-access',
		title: 'Member access on scalar variable',
		defaultSeverity: 'error',
		category: 'semantic',
		vbeCompileEquivalent: true,
		diagnosticKind: 'compile-error',
		source: 'XLIDE',
		specReference: 'VBE oracle: Invalid qualifier / Syntax error',
		confidence: 'high',
	},
	callRequiresParens: {
		code: 'call-requires-parens',
		title: 'Call statement requires parentheses around arguments',
		defaultSeverity: 'error',
		category: 'syntax',
		vbeCompileEquivalent: true,
		diagnosticKind: 'compile-error',
		source: 'XLIDE',
		specReference: 'MS-VBAL 5.4.2.1',
		confidence: 'high',
	},
	invalidExplicitCallTarget: {
		code: 'invalid-explicit-call-target',
		title: 'Invalid explicit Call target',
		defaultSeverity: 'error',
		category: 'syntax',
		vbeCompileEquivalent: true,
		diagnosticKind: 'compile-error',
		source: 'XLIDE',
		specReference: 'VBE oracle: Syntax error',
		confidence: 'high',
	},
	callStatementForbidsParens: {
		code: 'call-statement-forbids-parens',
		title: 'Standalone zero-argument call cannot use empty parentheses',
		defaultSeverity: 'error',
		category: 'syntax',
		vbeCompileEquivalent: true,
		diagnosticKind: 'compile-error',
		source: 'XLIDE',
		specReference: 'MS-VBAL 5.4.2.1',
		confidence: 'high',
	},
	expressionCallRequiresParens: {
		code: 'expression-call-requires-parens',
		title: 'Function call in an expression requires parentheses around arguments',
		defaultSeverity: 'error',
		category: 'syntax',
		vbeCompileEquivalent: true,
		diagnosticKind: 'compile-error',
		source: 'XLIDE',
		specReference: 'MS-VBAL 5.6.9',
		confidence: 'high',
	},
	invalidExpressionSyntax: {
		code: 'invalid-expression-syntax',
		title: 'Invalid expression syntax',
		defaultSeverity: 'error',
		category: 'syntax',
		vbeCompileEquivalent: true,
		diagnosticKind: 'compile-error',
		source: 'XLIDE',
		specReference: 'MS-VBAL 5.6 / VBE oracle: Syntax error',
		confidence: 'high',
	},
	requiredParamAfterOptional: {
		code: 'required-param-after-optional',
		title: 'A required parameter cannot follow an Optional parameter',
		defaultSeverity: 'error',
		category: 'declaration',
		vbeCompileEquivalent: true,
		diagnosticKind: 'compile-error',
		source: 'XLIDE',
		specReference: 'MS-VBAL 5.3.1.5',
		confidence: 'high',
	},
	paramArrayNotLast: {
		code: 'paramarray-not-last',
		title: 'ParamArray must be the final parameter',
		defaultSeverity: 'error',
		category: 'declaration',
		vbeCompileEquivalent: true,
		diagnosticKind: 'compile-error',
		source: 'XLIDE',
		specReference: 'MS-VBAL 5.3.1.6',
		confidence: 'high',
	},
	exitWrongProcedure: {
		code: 'exit-wrong-proc',
		title: 'Exit statement does not match the enclosing procedure',
		defaultSeverity: 'error',
		category: 'semantic',
		vbeCompileEquivalent: true,
		diagnosticKind: 'compile-error',
		source: 'XLIDE',
		specReference: 'MS-VBAL 5.4.1.3',
		confidence: 'high',
	},
	optionAfterDeclaration: {
		code: 'option-after-declaration',
		title: 'Option statement must precede all declarations',
		defaultSeverity: 'error',
		category: 'declaration',
		vbeCompileEquivalent: true,
		diagnosticKind: 'compile-error',
		source: 'XLIDE',
		specReference: 'MS-VBAL 5.2.1',
		confidence: 'high',
	},
	ifMissingThen: {
		code: 'if-missing-then',
		title: 'If statement is missing Then',
		defaultSeverity: 'error',
		category: 'syntax',
		vbeCompileEquivalent: true,
		diagnosticKind: 'compile-error',
		source: 'XLIDE',
		specReference: 'MS-VBAL 5.4.2.1',
		confidence: 'high',
	},
	caseOutsideSelect: {
		code: 'case-outside-select',
		title: 'Case statement outside Select Case',
		defaultSeverity: 'error',
		category: 'semantic',
		vbeCompileEquivalent: true,
		diagnosticKind: 'compile-error',
		source: 'XLIDE',
		specReference: 'MS-VBAL 5.4.2.4',
		confidence: 'high',
	},
	memberAccessOutsideWith: {
		code: 'member-access-outside-with',
		title: 'Leading member access outside With block',
		defaultSeverity: 'error',
		category: 'semantic',
		vbeCompileEquivalent: true,
		diagnosticKind: 'compile-error',
		source: 'XLIDE',
		specReference: 'MS-VBAL 5.4.2.6',
		confidence: 'high',
	},
	exitOutsideBlock: {
		code: 'exit-outside-block',
		title: 'Loop exit statement outside matching loop',
		defaultSeverity: 'error',
		category: 'semantic',
		vbeCompileEquivalent: true,
		diagnosticKind: 'compile-error',
		source: 'XLIDE',
		specReference: 'MS-VBAL 5.4.1.3',
		confidence: 'high',
	},
} satisfies Record<string, DiagnosticRuleMetadata>;

/** Stable rule-name keys of {@link DIAGNOSTIC_RULES}. */
export type DiagnosticRuleName = keyof typeof DIAGNOSTIC_RULES;
