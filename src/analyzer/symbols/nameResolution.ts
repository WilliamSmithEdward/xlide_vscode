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
	/**
	 * The procedure enclosing the resolved offset, echoed back from the input so
	 * hot-path callers (referenceScope) can reuse it instead of re-running the
	 * O(n) enclosing-procedure scan.
	 */
	enclosingProcedure?: VbaSymbol;
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

// Per-module index of module-level symbols (and enum members) by lowercased
// name. Identifier resolution runs per reference, so a linear scan over the
// module's declarations makes large modules quadratic. The module root symbol
// is stable for the duration of an analysis pass; the index preserves the
// original scan order (declaration order, enum members inline after their enum)
// so context filtering returns the same symbols in the same order.
const MODULE_MATCH_INDEXES = new WeakMap<VbaSymbol, Map<string, VbaSymbol[]>>();

function moduleMatchIndex(root: VbaSymbol): Map<string, VbaSymbol[]> {
	const cached = MODULE_MATCH_INDEXES.get(root);
	if (cached) {
		return cached;
	}
	const index = new Map<string, VbaSymbol[]>();
	const add = (symbol: VbaSymbol): void => {
		const key = symbol.name.toLowerCase();
		const bucket = index.get(key);
		if (bucket) {
			bucket.push(symbol);
		} else {
			index.set(key, [symbol]);
		}
	};
	for (const symbol of root.children ?? []) {
		add(symbol);
		if (symbol.kind === 'enum') {
			for (const member of symbol.children ?? []) {
				add(member);
			}
		}
	}
	MODULE_MATCH_INDEXES.set(root, index);
	return index;
}

export function moduleLevelIdentifierMatches(
	mod: ModuleSymbols,
	lowerName: string,
	context: BareIdentifierContext,
): VbaSymbol[] {
	const bucket = moduleMatchIndex(mod.root).get(lowerName);
	if (!bucket) {
		return [];
	}
	return bucket.filter((symbol) => symbolAllowedInContext(symbol, context));
}

// Lowercased module-level names (declarations + enum members), cached per
// module root: sourceIdentifierNames runs once per procedure, so re-lowercasing
// every module-level declaration each time is O(procedures x declarations).
const MODULE_LOWER_NAMES = new WeakMap<VbaSymbol, readonly string[]>();

function moduleLowerNames(root: VbaSymbol): readonly string[] {
	const cached = MODULE_LOWER_NAMES.get(root);
	if (cached) {
		return cached;
	}
	const out: string[] = [];
	for (const symbol of root.children ?? []) {
		out.push(symbol.name.toLowerCase());
		if (symbol.kind === 'enum') {
			for (const member of symbol.children ?? []) {
				out.push(member.name.toLowerCase());
			}
		}
	}
	MODULE_LOWER_NAMES.set(root, out);
	return out;
}

// Same caching for the project-visible symbol list (stable array identity per pass).
const PROJECT_LOWER_NAMES = new WeakMap<readonly VbaSymbol[], readonly string[]>();

function projectLowerNames(symbols: readonly VbaSymbol[] | undefined): readonly string[] {
	if (!symbols || symbols.length === 0) {
		return [];
	}
	const cached = PROJECT_LOWER_NAMES.get(symbols);
	if (cached) {
		return cached;
	}
	const out = symbols.map((symbol) => symbol.name.toLowerCase());
	PROJECT_LOWER_NAMES.set(symbols, out);
	return out;
}

// Merged module + project names, cached per (module root, project list)
// pair. Every procedure shares the same base union; re-adding those names
// into a fresh Set per procedure was ~40 ms of a full analysis pass on the
// giant corpus class, so the base is built once and the per-procedure layer
// stays as small as the procedure's own locals.
const MERGED_BASE_NAMES = new WeakMap<VbaSymbol, WeakMap<object, ReadonlySet<string>>>();
const NO_PROJECT_SYMBOLS: readonly VbaSymbol[] = [];

function mergedBaseNames(
	root: VbaSymbol,
	projectVisibleSymbols: readonly VbaSymbol[] | undefined,
): ReadonlySet<string> {
	const projectKey = projectVisibleSymbols ?? NO_PROJECT_SYMBOLS;
	let byProject = MERGED_BASE_NAMES.get(root);
	if (!byProject) {
		byProject = new WeakMap();
		MERGED_BASE_NAMES.set(root, byProject);
	}
	const cached = byProject.get(projectKey);
	if (cached) {
		return cached;
	}
	const out = new Set<string>(moduleLowerNames(root));
	for (const lower of projectLowerNames(projectVisibleSymbols)) {
		out.add(lower);
	}
	byProject.set(projectKey, out);
	return out;
}

/**
 * Membership view over every identifier name the source declares around a
 * statement: the enclosing procedure's locals and return variable layered
 * over the module- and project-level names. Only `has` is offered so the
 * shared base set is never copied per procedure.
 */
export function sourceIdentifierNames(
	input: Pick<
		BareIdentifierResolutionInput,
		'currentModule' | 'enclosingProcedure' | 'projectVisibleSymbols'
	>,
): { has(lowerName: string): boolean } {
	const locals = new Set<string>();
	const returnVariable = procedureReturnVariable(input.enclosingProcedure, 'expression');
	if (returnVariable) {
		locals.add(returnVariable.name.toLowerCase());
	}
	for (const symbol of input.enclosingProcedure?.children ?? []) {
		if (isLocalIdentifierSymbol(symbol)) {
			locals.add(symbol.name.toLowerCase());
		}
	}
	const base = mergedBaseNames(input.currentModule.root, input.projectVisibleSymbols);
	return {
		has: (lowerName: string): boolean => locals.has(lowerName) || base.has(lowerName),
	};
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
		...(input.enclosingProcedure ? { enclosingProcedure: input.enclosingProcedure } : {}),
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
