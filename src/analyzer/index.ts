// Public surface of the ground-up VBA language analyzer.
//
// This module intentionally has no dependency on the `vscode` module so the
// whole analyzer can be unit-tested under vitest (Node environment) and reused
// from any host. See docs/xlide_vba_language_service_roadmap.md.

export * from './lexer/tokenKinds';
export { tokenize } from './lexer/tokenize';
export {
	canonicalKeyword,
	isReservedIdentifier,
	RESERVED_IDENTIFIERS,
	VBA_KEYWORDS,
} from './lexer/keywordTable';
export * from './parser/nodes';
export { parseModule } from './parser/parseModule';
export { splitLogicalStatements } from './parser/parserState';
export {
	collectConditionalDirectives,
	conditionalActivityAtOffset,
	conditionalActivityForSpan,
	conditionalCompilerConstants,
	createConditionalActivityTracker,
	evaluateConditionalExpression,
	indexConditionalCompilation,
	ConditionalActivity,
	ConditionalActivityTracker,
	ConditionalCompilationEnvironment,
	ConditionalCompilationIndex,
	ConditionalConstDefinition,
	ConditionalDirectiveOccurrence,
	ConditionalValue,
} from './conditional/conditionalCompilation';
export {
	EXCEL_OBJECT_MODEL,
	HostConstant,
	HostMember,
	HostMemberKind,
	HostObjectModel,
	HostType,
} from './host/excelObjectModel';
export {
	getHostConstants,
	getHostGlobals,
	getHostMembers,
	getHostType,
	HostGlobal,
	resolveHostAlias,
	resolveHostConstant,
	resolveHostGlobal,
	resolveHostMemberSignature,
	resolveMemberReturnType,
} from './host/hostModel';
export {
	MemberCompletion,
	MemberCompletionContext,
	resolveMemberCompletions,
	resolveReceiverTypeAt,
} from './completion/memberAccess';
export {
	CanonicalCaseBoundaryKind,
	CanonicalCaseContext,
	CanonicalCaseEdit,
	CanonicalCaseSpan,
	canonicalCaseBoundaryKind,
	resolveCanonicalCaseEdit,
	resolveCanonicalCaseEdits,
} from './completion/canonicalCasing';
export {
	isCreatableTypeCompletion,
	ProjectTypeName,
	resolveTypeCompletions,
	TypeCompletion,
	TypeCompletionContext,
	TypeCompletionKind,
} from './completion/typeCompletion';
export {
	bareCallStatementTarget,
	explicitCallStatementBareRuntimeRewrite,
	explicitCallStatementArgumentListWithoutParens,
	explicitCallStatementArgumentWithoutParens,
	explicitCallStatementTarget,
	findActiveCallSite,
	standaloneEmptyParenthesizedCallStatement,
	BareCallStatementTarget,
	ExplicitCallStatementBareRuntimeRewrite,
	ExplicitCallStatementArgumentList,
	ParenthesizedCallStatementTarget,
	VbaCallSite,
	VbaTextSpan,
} from './call/callContext';
export {
	IdentifierCompletion,
	IdentifierCompletionContext,
	IdentifierCompletionKind,
	callableCompletionShouldInsertParens,
	resolveIdentifierCompletions,
} from './completion/identifierCompletion';
export {
	KeywordCompletion,
	KeywordCompletionKind,
	KeywordCompletionOptions,
	KeywordCompletionResult,
	materializeKeywordSnippet,
	resolveKeywordCompletions,
} from './completion/keywordCompletion';
export {
	collectProcedureLabelReferences,
	collectProcedureLabels,
	resolveProcedureLabelCompletions,
	resolveProcedureLabelDefinitionAt,
	statementLabelReferences,
	VbaProcedureLabel,
	VbaProcedureLabelCompletion,
	VbaProcedureLabelDefinition,
	VbaProcedureLabelReference,
} from './flow/procedureLabels';
export {
	EventHandlerCompletion,
	EventHandlerCompletionContext,
	EventHandlerDocumentType,
	EventHandlerProcedureMatch,
	eventHandlerDocumentTypeForContext,
	eventHandlerProcedureForName,
	resolveEventHandlerCompletions,
} from './completion/eventHandlers';
export {
	collectTypeNameReferences,
	ResolvedTypeReference,
	resolveTypeReferenceAt,
	resolveTypeSemanticTokens,
	TypeNameReference,
	TypeNameReferenceKind,
	TypeSemanticToken,
	TypeSemanticTokenType,
	typeReferenceLookupName,
} from './semantic/typeSemanticTokens';
export {
	HoverContext,
	HoverInfo,
	resolveHover,
} from './hover/resolveHover';
export {
	resolveRuntimeConstant,
	resolveRuntimeFunction,
	resolveRuntimeObject,
	resolveRuntimeObjectType,
	runtimeAllowsExplicitCall,
	VBA_RUNTIME_CONSTANTS,
	VBA_RUNTIME_FUNCTIONS,
	VBA_RUNTIME_OBJECTS,
	VbaRuntimeConstant,
	VbaRuntimeFunction,
	VbaRuntimeObject,
	VbaRuntimeObjectMember,
} from './runtime/vbaRuntime';
export {
	resolveSignatureHelp,
	SignatureHelpContext,
	SignatureInfo,
	SignatureParameter,
} from './signature/signatureHelp';
export {
	hasDocContent,
	renderDocMarkdown,
	renderParamDocMarkdown,
	renderSignatureDocMarkdown,
	VbaDoc,
	VbaDocParam,
	VbaDocSource,
} from './docs/docModel';
export { extractLeadingDoc, extractModuleHeaderDoc, parseDocBody } from './docs/docComment';
export { ExternalDocEntry, parseMetadataFile } from './docs/externalDoc';
export { DocRegistry } from './docs/docRegistry';
export {
	isProcedureKind,
	ModuleSymbolKind,
	ModuleSymbols,
	qualifiedProcedureKey,
	SymbolVisibility,
	VbaSymbol,
	VbaSymbolKind,
	VbaProcedureParam,
	VbaProcedureSignature,
	VbaProjectClassMemberDefinition,
	VbaProjectClassMember,
	VbaProjectClassMembers,
	VbaProjectTypeKind,
	VbaProjectTypeName,
	VbaSymbolAttribute,
} from './symbols/symbolModel';
export { buildModuleSymbols, BuildModuleSymbolsOptions } from './symbols/buildModuleSymbols';
export {
	BareIdentifierContext,
	BareIdentifierResolution,
	BareIdentifierResolutionScope,
} from './symbols/nameResolution';
export {
	ModuleInput,
	ProjectIndex,
	ProjectIndexOptions,
	ReferenceScope,
	ReferenceScopeKind,
	ShadowedSpan,
} from './symbols/projectIndex';
export {
	analyzeModule,
	AnalyzeModuleOptions,
	DiagnosticSeverityOverrides,
	incompleteExpressionEditSpan,
	VbaCreateProcedureStubData,
	VbaDiagnostic,
	VbaDiagnosticData,
	VbaMissingRequiredArgumentPlaceholderData,
} from './diagnostics/analyzeModule';
export {
	normalizeDiagnosticCode,
	resolveDiagnosticCodeActions,
	VbaDiagnosticCodeAction,
	VbaDiagnosticCodeActionInput,
	VbaTextEdit,
} from './codeActions/diagnosticCodeActions';
export {
	DIAGNOSTIC_RULES,
	STRUCTURAL_DIAGNOSTIC_RULES,
	diagnosticMetadataForCode,
	diagnosticSuppressionScopesForCode,
	diagnosticSourceForCode,
	allowedDiagnosticSeverityOverridesForCode,
	DiagnosticCategory,
	DiagnosticEvidenceKind,
	DiagnosticRuleMetadata,
	DiagnosticRuleName,
	DiagnosticSeverity,
	DiagnosticSeverityOverride,
	DiagnosticSuppressionScope,
	isDiagnosticSeverityOverride,
	isXlideDiagnosticSource,
	normalizeDiagnosticSeverityOverride,
	normalizeDiagnosticSeverityOverrides,
	XLIDE_DIAGNOSTIC_SOURCE,
} from './diagnostics/ruleMetadata';
export {
	filterDiagnosticsWithSuppressions,
	ANALYSIS_SUPPRESSION_DIRECTIVE_CODE,
	AnalysisSuppressionAnalysis,
	AnalysisSuppressionFilterResult,
	scanAnalysisSuppressions,
} from './diagnostics/analysisSuppressions';
