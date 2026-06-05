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
export type DiagnosticSeverity = 'error' | 'warning' | 'information';

/** User-configurable severity override value for an analysis rule. */
export type DiagnosticSeverityOverride = DiagnosticSeverity | 'off';

/** Analysis suppression scopes exposed by the workbook analysis UI. */
export type DiagnosticSuppressionScope = 'block' | 'member' | 'module';

const DEFAULT_DIAGNOSTIC_SUPPRESSION_SCOPES: readonly DiagnosticSuppressionScope[] = [
	'block',
	'member',
	'module',
];

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
	/** Suppression scopes that make sense for this diagnostic rule before source-position filtering. */
	suppressionScopes?: readonly DiagnosticSuppressionScope[];
	/** True when an error rule is allowed to be downgraded to warning by user settings. */
	allowSeverityDowngrade?: boolean;
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
	invalidLineContinuation: {
		code: 'invalid-line-continuation',
		title: 'Invalid line continuation',
		defaultSeverity: 'error',
		category: 'syntax',
		vbeCompileEquivalent: true,
		diagnosticKind: 'compile-error',
		source: 'XLIDE',
		specReference: 'MS-VBAL 3.2.2',
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
	declareMissingPtrSafe: {
		code: 'declare-missing-ptrsafe',
		title: 'Declare statement missing PtrSafe for 64-bit Office',
		defaultSeverity: 'error',
		category: 'declaration',
		vbeCompileEquivalent: true,
		diagnosticKind: 'compile-error',
		source: 'XLIDE',
		specReference: 'VBA 7 Declare statement PtrSafe requirement for 64-bit Office',
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
		suppressionScopes: ['module'],
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
		allowSeverityDowngrade: true,
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
	invalidIdentifierStart: {
		code: 'invalid-identifier-start',
		title: 'Invalid identifier start',
		defaultSeverity: 'error',
		category: 'declaration',
		vbeCompileEquivalent: true,
		diagnosticKind: 'compile-error',
		source: 'XLIDE',
		specReference: 'MS-VBAL 3.3.5',
		confidence: 'high',
	},
	moduleDeclarationInProcedure: {
		code: 'module-declaration-in-procedure',
		title: 'Module-level declaration inside procedure',
		defaultSeverity: 'error',
		category: 'declaration',
		vbeCompileEquivalent: true,
		diagnosticKind: 'compile-error',
		source: 'XLIDE',
		specReference: 'MS-VBAL 5.2 / 5.3',
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
		specReference: 'MS-VBAL 5.3.1 / runtime type coercion and numeric overflow',
		allowSeverityDowngrade: true,
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
	byRefArgumentTypeMismatch: {
		code: 'byref-argument-type-mismatch',
		title: 'ByRef argument type mismatch',
		defaultSeverity: 'error',
		category: 'semantic',
		vbeCompileEquivalent: true,
		diagnosticKind: 'compile-error',
		source: 'XLIDE',
		specReference: 'MS-VBAL 5.3.1 / VBE oracle: ByRef argument type mismatch',
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
		specReference: 'MS-VBAL 5.4.3 / runtime type coercion and numeric overflow',
		allowSeverityDowngrade: true,
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
		allowSeverityDowngrade: true,
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
		allowSeverityDowngrade: true,
		confidence: 'high',
	},
	divisionByZero: {
		code: 'division-by-zero',
		title: 'Division by zero',
		defaultSeverity: 'error',
		category: 'semantic',
		vbeCompileEquivalent: false,
		diagnosticKind: 'deterministic-runtime-error',
		source: 'XLIDE',
		specReference: 'MS-VBAL 5.6 / runtime division by zero',
		allowSeverityDowngrade: true,
		confidence: 'high',
	},
	runtimeArgumentValue: {
		code: 'runtime-argument-value',
		title: 'Invalid runtime argument value',
		defaultSeverity: 'error',
		category: 'semantic',
		vbeCompileEquivalent: false,
		diagnosticKind: 'deterministic-runtime-error',
		source: 'XLIDE',
		specReference: 'MS-VBAL 5.6 / VBA runtime argument bounds and VBE oracle runtime error 5',
		allowSeverityDowngrade: true,
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
		allowSeverityDowngrade: true,
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
	fixedLengthStringSize: {
		code: 'fixed-length-string-size',
		title: 'Invalid fixed-length String size',
		defaultSeverity: 'error',
		category: 'declaration',
		vbeCompileEquivalent: true,
		diagnosticKind: 'compile-error',
		source: 'XLIDE',
		specReference: 'MS-VBAL fixed-length String bounds / VBE oracle: Invalid length for fixed-length string',
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
	eventDeclarationModuleKind: {
		code: 'event-declaration-module-kind',
		title: 'Event declaration is not valid in this module',
		defaultSeverity: 'error',
		category: 'module-kind',
		vbeCompileEquivalent: true,
		diagnosticKind: 'compile-error',
		source: 'XLIDE',
		specReference: 'MS-VBAL 5.2.5: Event declarations belong to object modules',
		confidence: 'high',
	},
	withEventsDeclaration: {
		code: 'withevents-declaration',
		title: 'Invalid WithEvents declaration',
		defaultSeverity: 'error',
		category: 'module-kind',
		vbeCompileEquivalent: true,
		diagnosticKind: 'compile-error',
		source: 'XLIDE',
		specReference: 'MS-VBAL 5.2.3: WithEvents object variable declarations',
		confidence: 'high',
	},
	implementsStatementPlacement: {
		code: 'implements-statement-placement',
		title: 'Invalid Implements statement',
		defaultSeverity: 'error',
		category: 'module-kind',
		vbeCompileEquivalent: true,
		diagnosticKind: 'compile-error',
		source: 'XLIDE',
		specReference: 'MS-VBAL Implements statement: module-level object-module declaration',
		confidence: 'high',
	},
	raiseEventUndeclaredEvent: {
		code: 'raiseevent-undeclared-event',
		title: 'RaiseEvent target is not declared',
		defaultSeverity: 'error',
		category: 'semantic',
		vbeCompileEquivalent: true,
		diagnosticKind: 'compile-error',
		source: 'XLIDE',
		specReference: 'MS-VBAL RaiseEvent statement: event name resolution',
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
	invalidNewTypeName: {
		code: 'invalid-new-type-name',
		title: 'Type cannot be created with New',
		defaultSeverity: 'error',
		category: 'declaration',
		vbeCompileEquivalent: true,
		diagnosticKind: 'compile-error',
		source: 'XLIDE',
		specReference: 'MS-VBAL 5.2.3.1 / 5.6.9',
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
	parameterDefaultTypeMismatch: {
		code: 'parameter-default-type-mismatch',
		title: 'Parameter default type mismatch',
		defaultSeverity: 'error',
		category: 'declaration',
		vbeCompileEquivalent: true,
		diagnosticKind: 'compile-error',
		source: 'XLIDE',
		specReference: 'MS-VBAL 5.3.1 / VBE oracle: Type mismatch',
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
	undefinedLabel: {
		code: 'undefined-label',
		title: 'Label not defined',
		defaultSeverity: 'error',
		category: 'semantic',
		vbeCompileEquivalent: true,
		diagnosticKind: 'compile-error',
		source: 'XLIDE',
		specReference: 'MS-VBAL 5.4.1 / VBE oracle: Label not defined',
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
	vbaTestDirective: {
		code: 'vba-test-directive',
		title: 'Invalid XLIDE VBA test directive',
		defaultSeverity: 'warning',
		category: 'style',
		vbeCompileEquivalent: false,
		diagnosticKind: 'style-policy',
		source: 'XLIDE',
		specReference: 'docs/xlide_vba_com_test_runner.md',
		confidence: 'high',
	},
	analysisSuppressionDirective: {
		code: 'analysis-suppression-directive',
		title: 'Invalid XLIDE analysis suppression directive',
		defaultSeverity: 'warning',
		category: 'style',
		vbeCompileEquivalent: false,
		diagnosticKind: 'style-policy',
		source: 'XLIDE',
		specReference: 'docs/xlide_vba_analysis_suppression_comments.md',
		confidence: 'high',
	},
} satisfies Record<string, DiagnosticRuleMetadata>;

/** Structural diagnostics emitted by the dependency-free block-balance analyzer. */
export const STRUCTURAL_DIAGNOSTIC_RULES = {
	missingBlockCloser: {
		code: 'missing-block-closer',
		title: 'Missing block closer',
		defaultSeverity: 'error',
		category: 'syntax',
		vbeCompileEquivalent: true,
		diagnosticKind: 'compile-error',
		source: 'XLIDE',
		confidence: 'high',
	},
	unmatchedBlockCloser: {
		code: 'unmatched-block-closer',
		title: 'Unmatched block closer',
		defaultSeverity: 'error',
		category: 'syntax',
		vbeCompileEquivalent: true,
		diagnosticKind: 'compile-error',
		source: 'XLIDE',
		confidence: 'high',
	},
} satisfies Record<string, DiagnosticRuleMetadata>;

const DIAGNOSTIC_METADATA_BY_CODE = new Map<string, DiagnosticRuleMetadata>(
	[
		...Object.values(DIAGNOSTIC_RULES),
		...Object.values(STRUCTURAL_DIAGNOSTIC_RULES),
	].map((rule) => [rule.code, rule]),
);

export const XLIDE_DIAGNOSTIC_SOURCE = 'XLIDE';

/** Returns metadata for semantic and structural diagnostic codes. */
export function diagnosticMetadataForCode(
	code: string | undefined,
): DiagnosticRuleMetadata | undefined {
	const normalized = normalizeDiagnosticRuleCode(code);
	if (!normalized) {
		return undefined;
	}
	return DIAGNOSTIC_METADATA_BY_CODE.get(normalized);
}

/** Metadata for every active diagnostic rule, sorted by stable diagnostic code. */
export function allDiagnosticRuleMetadata(): DiagnosticRuleMetadata[] {
	return Array.from(DIAGNOSTIC_METADATA_BY_CODE.values())
		.sort((left, right) => left.code.localeCompare(right.code));
}

/** True when a diagnostic code is part of the active XLIDE rule catalog. */
export function isKnownDiagnosticRuleCode(code: string | undefined): boolean {
	return diagnosticMetadataForCode(code) !== undefined;
}

/** Severity override choices allowed for one diagnostic code. */
export function allowedDiagnosticSeverityOverridesForCode(
	code: string | undefined,
): readonly DiagnosticSeverityOverride[] {
	const meta = diagnosticMetadataForCode(code);
	if (!meta) {
		return [];
	}
	if (meta.defaultSeverity === 'error') {
		return meta.allowSeverityDowngrade ? ['warning'] : [];
	}
	return ['off'];
}

/** Normalizes one guarded severity override. Invalid or disallowed values are ignored. */
export function normalizeDiagnosticSeverityOverride(
	code: string | undefined,
	value: unknown,
): DiagnosticSeverityOverride | undefined {
	if (!isDiagnosticSeverityOverride(value)) {
		return undefined;
	}
	const allowed = allowedDiagnosticSeverityOverridesForCode(code);
	return allowed.includes(value) ? value : undefined;
}

/** Normalizes a user-provided map keyed by stable diagnostic code. */
export function normalizeDiagnosticSeverityOverrides(
	value: unknown,
): Record<string, DiagnosticSeverityOverride> {
	if (!value || typeof value !== 'object' || Array.isArray(value)) {
		return {};
	}
	const normalized: Record<string, DiagnosticSeverityOverride> = {};
	for (const [rawCode, rawSeverity] of Object.entries(value)) {
		const code = normalizeDiagnosticRuleCode(rawCode);
		const severity = normalizeDiagnosticSeverityOverride(code, rawSeverity);
		if (code && severity) {
			normalized[code] = severity;
		}
	}
	return sortDiagnosticSeverityOverrides(normalized);
}

/** True when a raw value is part of the severity-override vocabulary. */
export function isDiagnosticSeverityOverride(value: unknown): value is DiagnosticSeverityOverride {
	return value === 'off' || value === 'information' || value === 'warning' || value === 'error';
}

function normalizeDiagnosticRuleCode(code: unknown): string | undefined {
	return typeof code === 'string' ? code.trim().toLowerCase() || undefined : undefined;
}

function sortDiagnosticSeverityOverrides(
	overrides: Record<string, DiagnosticSeverityOverride>,
): Record<string, DiagnosticSeverityOverride> {
	return Object.fromEntries(
		Object.entries(overrides).sort(([left], [right]) => left.localeCompare(right)),
	);
}

/** Rule-level suppression scopes before narrowing by the diagnostic's source position. */
export function diagnosticSuppressionScopesForCode(
	code: string | undefined,
): readonly DiagnosticSuppressionScope[] {
	return diagnosticMetadataForCode(code)?.suppressionScopes ?? DEFAULT_DIAGNOSTIC_SUPPRESSION_SCOPES;
}

/** Problems-panel source label for a diagnostic code. */
export function diagnosticSourceForCode(code: string | undefined): string {
	const meta = diagnosticMetadataForCode(code);
	if (!meta) {
		return XLIDE_DIAGNOSTIC_SOURCE;
	}
	if (meta.vbeCompileEquivalent) {
		return `${XLIDE_DIAGNOSTIC_SOURCE}/VBE`;
	}
	switch (meta.diagnosticKind) {
		case 'deterministic-runtime-error':
			return `${XLIDE_DIAGNOSTIC_SOURCE}/runtime`;
		case 'runtime-risk':
			return `${XLIDE_DIAGNOSTIC_SOURCE}/risk`;
		case 'style-policy':
			return `${XLIDE_DIAGNOSTIC_SOURCE}/style`;
		case 'compile-error':
			return XLIDE_DIAGNOSTIC_SOURCE;
	}
}

/** True for canonical and metadata-expanded XLIDE diagnostic source labels. */
export function isXlideDiagnosticSource(source: string | undefined): boolean {
	return source === XLIDE_DIAGNOSTIC_SOURCE || source?.startsWith(`${XLIDE_DIAGNOSTIC_SOURCE}/`) === true;
}

/** Stable rule-name keys of {@link DIAGNOSTIC_RULES}. */
export type DiagnosticRuleName = keyof typeof DIAGNOSTIC_RULES;
