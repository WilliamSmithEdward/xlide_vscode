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

import type { HostObjectModel } from './excelObjectModel';
import { getExcelObjectModel } from './excelObjectModel';
import { getWordObjectModel } from './wordObjectModel';
import { getPowerPointObjectModel } from './powerpointObjectModel';
import { getAccessObjectModel } from './accessObjectModel';
import { getVb6ObjectModel } from './vb6ObjectModel';

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
	['word', getWordObjectModel],
	['powerpoint', getPowerPointObjectModel],
	['access', getAccessObjectModel],
	['vb6', getVb6ObjectModel],
]);

/**
 * Registers a host's model under its token. Called by each host model module
 * at load; exported so tests can register throwaway models.
 */
export function registerHostObjectModel(token: VbaHostToken, model: () => HostObjectModel): void {
	MODELS_BY_TOKEN.set(token, model);
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

/** Merged models, keyed by the token list that produced them. */
const MERGED_BY_KEY = new Map<string, HostObjectModel>();

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
	const merged: HostObjectModel = {
		source: models.map((one) => one.source).join(' + '),
		hostName: models[0].hostName,
		globalType: models[0].globalType,
		types: Object.assign({}, ...layered.map((one) => one.types)),
		aliases: Object.assign({}, ...layered.map((one) => one.aliases)),
		globals: Object.assign({}, ...layered.map((one) => one.globals)),
		constants: Object.assign({}, ...layered.map((one) => one.constants ?? {})),
		enums: Object.assign({}, ...layered.map((one) => one.enums ?? {})),
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
