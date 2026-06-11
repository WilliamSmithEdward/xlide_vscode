// Public surface of the ground-up VBA language analyzer.
//
// This module intentionally has no dependency on the `vscode` module so the
// whole analyzer can be unit-tested under vitest (Node environment) and reused
// from any host. See docs/xlide_vba_language_service_roadmap.md.

export * from './lexer/tokenKinds';
export { tokenize, tokenizeCached } from './lexer/tokenize';
export {
	canonicalKeyword,
	isReservedIdentifier,
	VBA_KEYWORDS,
} from './lexer/keywordTable';
export * from './parser/nodes';
export { parseModule } from './parser/parseModule';
export { splitLogicalStatements } from './parser/parserState';
export {
	collectConditionalDirectives,
	conditionalActivityAtOffset,
	conditionalCompilerConstants,
	createConditionalActivityTracker,
	evaluateConditionalExpression,
	indexConditionalCompilation,
	ConditionalActivityTracker,
	ConditionalCompilationEnvironment,
} from './conditional/conditionalCompilation';
export {
	getExcelObjectModel,
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
	resolveHostAlias,
	resolveHostConstant,
	resolveHostGlobal,
	resolveHostMemberSignature,
	resolveMemberReturnType,
} from './host/hostModel';
export {
	MemberCompletion,
	MemberCompletionContext,
	resolveMemberCompletionNamed,
	resolveMemberCompletions,
	resolveMemberDefinitionsAt,
} from './completion/memberAccess';
export {
	CanonicalCaseContext,
	CanonicalCaseEdit,
	canonicalCaseBoundaryKind,
	resolveCanonicalCaseEdit,
	resolveCanonicalCaseEdits,
} from './completion/canonicalCasing';
export {
	CompletionCursorContext,
	completionCursorContext,
	identifierSpanEndingAt,
	spaceTriggerMayComplete,
} from './completion/cursorContext';
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
	VbaCallSite,
} from './call/callContext';
export {
	IdentifierCompletion,
	IdentifierCompletionContext,
	callableCompletionShouldInsertParens,
	resolveIdentifierCompletions,
} from './completion/identifierCompletion';
export {
	KeywordCompletion,
	materializeKeywordSnippet,
	resolveKeywordCompletions,
} from './completion/keywordCompletion';
export {
	collectProcedureLabelReferences,
	collectProcedureLabels,
	resolveProcedureLabelCompletions,
	resolveProcedureLabelDefinitionAt,
	VbaProcedureLabelCompletion,
} from './flow/procedureLabels';
export {
	EventHandlerCompletion,
	EventHandlerCompletionContext,
	EventHandlerDocumentType,
	eventHandlerDocumentTypeForContext,
	eventHandlerProcedureForName,
	resolveEventHandlerCompletions,
} from './completion/eventHandlers';
export {
	collectTypeNameReferences,
	ResolvedTypeReference,
	resolveTypeReferenceAt,
	resolveTypeSemanticTokens,
	TypeNameReferenceKind,
	TypeSemanticTokenType,
	typeReferenceLookupName,
} from './semantic/typeSemanticTokens';
export {
	HoverContext,
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
} from './runtime/vbaRuntime';
export {
	resolveSignatureHelp,
	SignatureHelpContext,
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
} from './symbols/nameResolution';
export {
	ProjectIndex,
	ReferenceScope,
} from './symbols/projectIndex';
export {
	analyzeModule,
	AnalyzeModuleOptions,
	incompleteExpressionEditSpan,
	VbaDiagnostic,
	VbaDiagnosticData,
} from './diagnostics/analyzeModule';
export {
	normalizeDiagnosticCode,
	resolveDiagnosticCodeActions,
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
	isXlideDiagnosticSource,
	normalizeDiagnosticSeverityOverride,
	normalizeDiagnosticSeverityOverrides,
} from './diagnostics/ruleMetadata';
export {
	filterDiagnosticsWithSuppressions,
	ANALYSIS_SUPPRESSION_DIRECTIVE_CODE,
	scanAnalysisSuppressions,
} from './diagnostics/analysisSuppressions';
