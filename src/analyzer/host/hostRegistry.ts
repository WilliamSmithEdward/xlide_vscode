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

/** The host tokens xlide_vbide sends with project/open. */
export type VbaHostToken =
	| 'excel'
	| 'word'
	| 'powerpoint'
	| 'access'
	| 'outlook'
	| 'visio'
	| 'project'
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
		default:
			return undefined;
	}
}
