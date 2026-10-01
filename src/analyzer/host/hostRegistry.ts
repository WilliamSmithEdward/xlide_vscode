// Which Office host a module's VBA belongs to, and the object model that
// answers for it.
//
// The resolvers have always accepted any HostObjectModel and defaulted to
// Excel's; this is the seam that lets a caller choose (issue #24). The token
// vocabulary is the one xlide_vbide already normalises from the process image,
// so an embedder passes the string it has.
//
// The semantics are deliberately asymmetric:
//
// - ABSENT means Excel. Every existing caller keeps exactly the behavior it
//   had, and xlide's own workbooks are Excel workbooks.
// - A NAMED host with no model yet means NO HOST KNOWLEDGE - an empty model,
//   not Excel's. Telling Word's ThisDocument it has Cells and Range was the
//   bug that motivated the seam; silence is the honest answer until the
//   host's own model exists.
//
// Only Excel is built in. The other hosts' models are most of the analyzer's
// size, so a caller registers the ones it analyzes: the extension registers
// every host it ships (registerBuiltInHostModels), and an embedder that only
// analyzes Excel registers none and bundles none of them.

import type { HostObjectModel } from './excelObjectModel';
import { getExcelObjectModel } from './excelObjectModel';

/**
 * The host tokens xlide_vbide sends with project/open, plus 'vb6': a VB6
 * project is not an Office host at all, but the analyzer selects an object
 * model by this token, and a VB6 form's code-behind needs the VB runtime's
 * surface (App, Screen, Printer, the intrinsic controls), not Excel's.
 */
export type VbaHostToken =
	| 'excel'
	| 'word'
	| 'powerpoint'
	| 'access'
	| 'outlook'
	| 'visio'
	| 'project'
	| 'vb6'
	| 'other';

/** A model that knows nothing: every lookup misses, so nothing is asserted. */
export const EMPTY_HOST_MODEL: HostObjectModel = Object.freeze({
	source: 'none: a host whose object model is not yet available',
	types: Object.freeze({}),
	aliases: Object.freeze({}),
	globals: Object.freeze({}),
});

const MODELS_BY_TOKEN = new Map<string, () => HostObjectModel>([
	['excel', getExcelObjectModel],
]);

/** Merged models, keyed by the token list that produced them. */
const MERGED_BY_KEY = new Map<string, HostObjectModel>();

/**
 * The hosts XLIDE ships a model for besides Excel: what
 * registerBuiltInHostModels registers. Listed here, not imported, so naming
 * them costs nothing.
 */
export const BUILT_IN_HOST_TOKENS: readonly VbaHostToken[] = Object.freeze(['word', 'powerpoint', 'access', 'vb6']);

/**
 * The tokens in `tokens` that XLIDE ships a model for but nobody registered.
 * Analyzing with one of them gives that host no knowledge, so a caller that
 * forgot to register is told about potentially missing or false findings.
 */
export function unregisteredBuiltInHosts(tokens: readonly string[]): VbaHostToken[] {
	const out: VbaHostToken[] = [];
	for (const raw of tokens) {
		const token = raw.trim().toLowerCase() as VbaHostToken;
		if (BUILT_IN_HOST_TOKENS.includes(token) && !MODELS_BY_TOKEN.has(token) && !out.includes(token)) {
			out.push(token);
		}
	}
	return out;
}

/**
 * Registers a host's model under its token. A host that is never registered
 * answers the empty model. Missing host knowledge can suppress valid findings
 * and introduce false findings about host names.
 */
export function registerHostObjectModel(token: VbaHostToken, model: () => HostObjectModel): void {
	MODELS_BY_TOKEN.set(token, model);
	// A merged model built before this host was known left it out.
	MERGED_BY_KEY.clear();
}

/**
 * The model a host token selects. Absent (or unrecognised casing of `excel`)
 * answers undefined so downstream `?? getExcelObjectModel()` defaults keep
 * today's behavior; any other named host answers its model, or the empty
 * model when none is registered yet.
 */
export function hostObjectModelForToken(host: string | undefined): HostObjectModel | undefined {
	if (host === undefined) {
		return undefined;
	}
	const token = host.trim().toLowerCase();
	if (token === '' || token === 'excel') {
		return undefined;
	}
	return MODELS_BY_TOKEN.get(token)?.() ?? EMPTY_HOST_MODEL;
}

/**
 * One model answering for a project's own host and every library it
 * references, in the order VBA resolves them.
 *
 * A project that references another application's library can name its types
 * and call its members, so a Word document with a reference to Excel has to
 * be analyzed against both. The FIRST token wins every shared name, which is
 * how VBA resolves an ambiguous one: by the reference list's order, the
 * project's own host at the top. That is the same rule each host model
 * already applies to the shared Office library it folds in.
 *
 * Every library's globals are merged, because a library marks its global
 * object APPOBJECT in its type library and VBA binds that object's members
 * bare for anyone who references it - Excel, Word and PowerPoint through a
 * hidden `Global`, Access through `Application`. The host's own still wins a
 * collision.
 *
 * Returns undefined when the list adds nothing to what a single token would
 * have given, so every existing caller keeps the model it had - including the
 * bare `excel` that rides as the downstream `?? getExcelObjectModel()`
 * default.
 */
export function hostObjectModelForTokens(
	tokens: readonly string[],
): HostObjectModel | undefined {
	const known = tokens.filter((token) => MODELS_BY_TOKEN.has(token));
	if (known.length <= 1) {
		return hostObjectModelForToken(known[0] ?? tokens[0]);
	}
	const key = known.join('+');
	const cached = MERGED_BY_KEY.get(key);
	if (cached) {
		return cached;
	}
	// Later models are spread first so the earlier ones overwrite them: the
	// project's own host wins every name it shares with a referenced library.
	const models = known.map((token) => MODELS_BY_TOKEN.get(token)!());
	const layered = [...models].reverse();
	// A merged model has one hostName - the project's own host, by
	// construction - and every label built from it named that host, so a Word
	// member in an Excel workbook read as Excel's (issue #77). A type key
	// already names its library; an enum key does not, so each referenced
	// library's enums carry theirs from here on.
	const labelled = layered.map((model) => (model === models[0]
		? model.enums ?? {}
		: Object.fromEntries(Object.entries(model.enums ?? {}).map(
			([name, entry]) => [name, { ...entry, library: model.hostName }],
		))));
	const merged: HostObjectModel = {
		source: models.map((one) => one.source).join(' + '),
		hostName: models[0].hostName,
		globalType: models[0].globalType,
		types: Object.assign({}, ...layered.map((one) => one.types)),
		aliases: Object.assign({}, ...layered.map((one) => one.aliases)),
		globals: Object.assign({}, ...layered.map((one) => one.globals)),
		constants: Object.assign({}, ...layered.map((one) => one.constants ?? {})),
		enums: Object.assign({}, ...labelled),
		memberSignatures: Object.assign({}, ...layered.map((one) => one.memberSignatures ?? {})),
	};
	MERGED_BY_KEY.set(key, merged);
	return merged;
}

/**
 * The host a macro container implies, from its file name. XLIDE's own file
 * surfaces are the caller here, so the analyzer knows what kind of file a
 * module came from without anyone having to say.
 */
export function hostTokenForFileName(fileName: string): VbaHostToken | undefined {
	const match = /\.([a-z0-9]+)$/i.exec(fileName);
	switch (match?.[1]?.toLowerCase()) {
		case 'xlsm':
		case 'xlsb':
		case 'xlam':
		case 'xltm':
		case 'xls':
		case 'xlt':
		case 'xla':
			return 'excel';
		case 'docm':
		case 'dotm':
		case 'doc':
		case 'dot':
			return 'word';
		case 'pptm':
		case 'potm':
		case 'ppsm':
		case 'ppam':
		case 'ppt':
		case 'ppa':
			return 'powerpoint';
		case 'accdb':
		case 'accda':
		case 'mdb':
		case 'mda':
			return 'access';
		case 'vbp':
			return 'vb6';
		default:
			return undefined;
	}
}
