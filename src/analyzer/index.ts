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
	CANONICAL_COMPOUND_FORMS,
} from './lexer/keywordTable';
export * from './parser/nodes';
export { parseModule } from './parser/parseModule';
export { splitLogicalStatements } from './parser/parserState';
export {
	EXCEL_OBJECT_MODEL,
	HostMember,
	HostMemberKind,
	HostObjectModel,
	HostType,
} from './host/excelObjectModel';
export {
	getHostGlobals,
	getHostMembers,
	getHostType,
	HostGlobal,
	resolveHostAlias,
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
	ProjectTypeName,
	resolveTypeCompletions,
	TypeCompletion,
	TypeCompletionContext,
	TypeCompletionKind,
} from './completion/typeCompletion';
export {
	IdentifierCompletion,
	IdentifierCompletionContext,
	IdentifierCompletionKind,
	resolveIdentifierCompletions,
} from './completion/identifierCompletion';
export {
	HoverContext,
	HoverInfo,
	resolveHover,
} from './hover/resolveHover';
export {
	resolveRuntimeFunction,
	VBA_RUNTIME_FUNCTIONS,
	VbaRuntimeFunction,
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
export { extractLeadingDoc, parseDocBody } from './docs/docComment';
export { ExternalDocEntry, parseMetadataFile } from './docs/externalDoc';
export { DocRegistry } from './docs/docRegistry';
export {
	isProcedureKind,
	ModuleSymbolKind,
	ModuleSymbols,
	SymbolVisibility,
	VbaSymbol,
	VbaSymbolKind,
} from './symbols/symbolModel';
export { buildModuleSymbols } from './symbols/buildModuleSymbols';
export {
	ModuleInput,
	ProjectIndex,
	ReferenceScope,
	ReferenceScopeKind,
	ShadowedSpan,
} from './symbols/projectIndex';
export {
	analyzeModule,
	AnalyzeModuleOptions,
	SeverityOverrides,
	VbaDiagnostic,
} from './diagnostics/analyzeModule';
export {
	DIAGNOSTIC_RULES,
	DiagnosticCategory,
	DiagnosticRuleMetadata,
	DiagnosticRuleName,
	DiagnosticSeverity,
} from './diagnostics/ruleMetadata';
