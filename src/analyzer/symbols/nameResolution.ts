import {
	isProcedureKind,
	type ModuleSymbols,
	type VbaSymbol,
} from './symbolModel';

export type BareIdentifierContext =
	| 'expression'
	| 'call'
	| 'assignmentTarget'
	| 'memberReceiver'
	| 'typeName'
	| 'newExpression';

export type BareIdentifierResolutionScope =
	| 'local'
	| 'module'
	| 'project'
	| 'ambiguous'
	| 'unresolved';

export interface BareIdentifierResolution {
	name: string;
	lowerName: string;
	context: BareIdentifierContext;
	scope: BareIdentifierResolutionScope;
	tier?: Exclude<BareIdentifierResolutionScope, 'ambiguous' | 'unresolved'>;
	definitions: readonly VbaSymbol[];
	reason: string;
}

export interface BareIdentifierResolutionInput {
	currentModule: ModuleSymbols;
	name: string;
	context: BareIdentifierContext;
	enclosingProcedure?: VbaSymbol;
	offset?: number;
	projectVisibleSymbols?: readonly VbaSymbol[];
}

/** Context-aware source resolver for bare identifiers, ordered local -> module -> project. */
export function resolveBareIdentifierBinding(
	input: BareIdentifierResolutionInput,
): BareIdentifierResolution {
	const lowerName = input.name.toLowerCase();
	const local = localIdentifierMatches(input.enclosingProcedure, lowerName, input.context, input.offset);
	if (local.length > 0) {
		return resolution(input, lowerName, ambiguousScope(local, 'local'), local);
	}

	const module = moduleLevelIdentifierMatches(input.currentModule, lowerName, input.context);
	if (module.length > 0) {
		return resolution(input, lowerName, ambiguousScope(module, 'module'), module);
	}

	const currentLower = input.currentModule.moduleName.toLowerCase();
	const project = (input.projectVisibleSymbols ?? [])
		.filter((symbol) => symbol.moduleName.toLowerCase() !== currentLower)
		.filter((symbol) => symbol.name.toLowerCase() === lowerName)
		.filter((symbol) => symbolAllowedInContext(symbol, input.context));
	if (project.length > 0) {
		return resolution(input, lowerName, ambiguousScope(project, 'project'), project);
	}

	return {
		name: input.name,
		lowerName,
		context: input.context,
		scope: 'unresolved',
		definitions: [],
		reason: `No source-backed ${input.context} binding found for '${input.name}'.`,
	};
}

export function localIdentifierMatches(
	procedure: VbaSymbol | undefined,
	lowerName: string,
	context: BareIdentifierContext,
	offset?: number,
): VbaSymbol[] {
	if (!procedure || context === 'typeName' || context === 'newExpression') {
		return [];
	}
	const out: VbaSymbol[] = [];
	const returnVariable = procedureReturnVariable(procedure, context, offset);
	if (returnVariable?.name.toLowerCase() === lowerName) {
		out.push(returnVariable);
	}
	out.push(...(procedure.children ?? [])
		.filter((symbol) => isLocalIdentifierSymbol(symbol))
		.filter((symbol) => symbol.name.toLowerCase() === lowerName));
	return out;
}

export function moduleLevelIdentifierMatches(
	mod: ModuleSymbols,
	lowerName: string,
	context: BareIdentifierContext,
): VbaSymbol[] {
	const out: VbaSymbol[] = [];
	for (const symbol of mod.root.children ?? []) {
		if (
			symbol.name.toLowerCase() === lowerName &&
			symbolAllowedInContext(symbol, context)
		) {
			out.push(symbol);
		}
		if (symbol.kind === 'enum') {
			for (const member of symbol.children ?? []) {
				if (
					member.name.toLowerCase() === lowerName &&
					symbolAllowedInContext(member, context)
				) {
					out.push(member);
				}
			}
		}
	}
	return out;
}

export function sourceIdentifierNames(
	input: Pick<
		BareIdentifierResolutionInput,
		'currentModule' | 'enclosingProcedure' | 'projectVisibleSymbols'
	>,
): Set<string> {
	const out = new Set<string>();
	const returnVariable = procedureReturnVariable(input.enclosingProcedure, 'expression');
	if (returnVariable) {
		out.add(returnVariable.name.toLowerCase());
	}
	for (const symbol of input.enclosingProcedure?.children ?? []) {
		if (isLocalIdentifierSymbol(symbol)) {
			out.add(symbol.name.toLowerCase());
		}
	}
	for (const symbol of input.currentModule.root.children ?? []) {
		out.add(symbol.name.toLowerCase());
		if (symbol.kind === 'enum') {
			for (const member of symbol.children ?? []) {
				out.add(member.name.toLowerCase());
			}
		}
	}
	for (const symbol of input.projectVisibleSymbols ?? []) {
		out.add(symbol.name.toLowerCase());
	}
	return out;
}

function procedureReturnVariable(
	procedure: VbaSymbol | undefined,
	context: BareIdentifierContext,
	offset?: number,
): VbaSymbol | undefined {
	if (!procedureReturnsThroughName(procedure) || context === 'call') {
		return undefined;
	}
	if (offset !== undefined && offset <= procedure.nameSpan.end) {
		return undefined;
	}
	return {
		name: procedure.name,
		kind: 'localVariable',
		nameSpan: procedure.nameSpan,
		fullSpan: procedure.nameSpan,
		moduleName: procedure.moduleName,
		containerName: procedure.name,
		asType: procedure.asType,
		isArray: procedure.isArray,
		doc: procedure.doc,
	};
}

function procedureReturnsThroughName(
	procedure: VbaSymbol | undefined,
): procedure is VbaSymbol {
	return procedure?.kind === 'function' || procedure?.kind === 'propertyGet';
}

export function isLocalIdentifierSymbol(symbol: VbaSymbol): boolean {
	return (
		symbol.kind === 'parameter' ||
		symbol.kind === 'localVariable' ||
		(symbol.kind === 'constant' && !!symbol.containerName)
	);
}

function symbolAllowedInContext(
	symbol: VbaSymbol,
	context: BareIdentifierContext,
): boolean {
	if (context === 'typeName' || context === 'newExpression') {
		return symbol.kind === 'type' || symbol.kind === 'enum';
	}
	return true;
}

function ambiguousScope(
	definitions: readonly VbaSymbol[],
	fallback: Exclude<BareIdentifierResolutionScope, 'ambiguous' | 'unresolved'>,
): BareIdentifierResolutionScope {
	if (definitions.length <= 1 || isPropertyAccessorFamily(definitions)) {
		return fallback;
	}
	return 'ambiguous';
}

function isPropertyAccessorFamily(definitions: readonly VbaSymbol[]): boolean {
	if (definitions.length <= 1) {
		return false;
	}
	const first = definitions[0];
	return definitions.every(
		(symbol) =>
			symbol.moduleName.toLowerCase() === first.moduleName.toLowerCase() &&
			symbol.name.toLowerCase() === first.name.toLowerCase() &&
			(
				symbol.kind === 'propertyGet' ||
				symbol.kind === 'propertyLet' ||
				symbol.kind === 'propertySet'
			),
	);
}

function resolution(
	input: BareIdentifierResolutionInput,
	lowerName: string,
	scope: BareIdentifierResolutionScope,
	definitions: readonly VbaSymbol[],
): BareIdentifierResolution {
	const tier = scope === 'ambiguous'
		? definitionTier(input.currentModule, definitions)
		: scope === 'unresolved'
			? undefined
			: scope;
	const owner = definitions[0]?.moduleName ?? input.currentModule.moduleName;
	const label = scope === 'ambiguous'
		? 'ambiguous source-backed'
		: `${scope} source-backed`;
	return {
		name: input.name,
		lowerName,
		context: input.context,
		scope,
		...(tier ? { tier } : {}),
		definitions,
		reason: `${label} ${input.context} binding for '${input.name}' in ${owner}.`,
	};
}

function definitionTier(
	currentModule: ModuleSymbols,
	definitions: readonly VbaSymbol[],
): Exclude<BareIdentifierResolutionScope, 'ambiguous' | 'unresolved'> | undefined {
	const first = definitions[0];
	if (!first) {
		return undefined;
	}
	if (isLocalIdentifierSymbol(first)) {
		return 'local';
	}
	return first.moduleName.toLowerCase() === currentModule.moduleName.toLowerCase()
		? 'module'
		: 'project';
}
