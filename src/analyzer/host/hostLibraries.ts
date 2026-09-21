// The Office type libraries a VBA project can reference, and the host model
// that answers for each.
//
// A project that references another application's library can name its types
// and call its members: a Word document with a reference to the Excel library
// compiles `Dim xl As Excel.Application`, and `xl.Calculate` is checked
// against Excel's object model rather than Word's. Reading only the host a
// file's extension implies answers nothing for that code.
//
// A reference is stored as a libid, `*\G{guid}#major.minor#lcid#path#name`.
// The GUID is the identity: the path is a hint the host resolves through the
// registry, which is why a file written on one machine loads on another whose
// libraries sit elsewhere. So the GUID is what this maps, never the path or
// the description.
//
// Each GUID below was read from the registered type library on a machine with
// Office 16, together with the library's own name, which is the qualifier VBA
// writes: Excel.Application, Word.Document, PowerPoint.Slide,
// Access.Application.

import type { VbaHostToken } from './hostRegistry';

/** The library GUID a reference declares -> the host whose model answers. */
const HOST_BY_LIBRARY_GUID: ReadonlyMap<string, VbaHostToken> = new Map([
	['{00020813-0000-0000-C000-000000000046}', 'excel' as const],
	['{00020905-0000-0000-C000-000000000046}', 'word' as const],
	['{91493440-5A91-11CF-8700-00AA0060263B}', 'powerpoint' as const],
	['{4AFFC9A0-5F99-101B-AF4E-00AA003F0F07}', 'access' as const],
]);

/** The name each library gives itself, which is the qualifier VBA writes. */
export const HOST_LIBRARY_NAMES: Readonly<Record<VbaHostToken, string>> = Object.freeze({
	excel: 'Excel',
	word: 'Word',
	powerpoint: 'PowerPoint',
	access: 'Access',
	outlook: 'Outlook',
	visio: 'Visio',
	project: 'MSProject',
	vb6: 'VB',
	other: '',
});

/** The GUID inside a libid, upper-cased with its braces, or undefined. */
export function libraryGuidOf(libid: string): string | undefined {
	const found = /\{[0-9a-fA-F-]{36}\}/.exec(libid);
	return found ? found[0].toUpperCase() : undefined;
}

/**
 * The host whose object model answers for a reference, or undefined for one
 * XLIDE has no model for - stdole, the shared Office library, a third-party
 * DLL. Undefined means no knowledge, and silence is the honest answer for it;
 * the same rule the registry applies to an unmodelled host.
 */
export function hostTokenForLibid(libid: string): VbaHostToken | undefined {
	const guid = libraryGuidOf(libid);
	return guid === undefined ? undefined : HOST_BY_LIBRARY_GUID.get(guid);
}

/** One reference as the dir stream declares it, narrowed to what this needs. */
export interface DeclaredReference {
	name: string;
	libid: string;
}

/**
 * The hosts a project's references bring in, in the order VBA resolves them:
 * the project's own host first, then each referenced library in the order the
 * project declares it. VBA resolves an ambiguous name by that order, so a
 * Word document referencing Excel keeps Word's Range for `Dim r As Range`.
 *
 * The host itself is included even when its library is not in the reference
 * list, because a project's own host library is implicit.
 */
export function hostTokensForProject(
	host: VbaHostToken | undefined,
	references: readonly DeclaredReference[],
): VbaHostToken[] {
	const out: VbaHostToken[] = [];
	const seen = new Set<VbaHostToken>();
	for (const token of [host, ...references.map((one) => hostTokenForLibid(one.libid))]) {
		if (token === undefined || seen.has(token)) { continue; }
		seen.add(token);
		out.push(token);
	}
	return out;
}

/**
 * The same list without the project's own host: the libraries its code can
 * name because the project references them. The analyzer takes the host and
 * these separately, because the host also decides what `Me` is and which
 * host-specific rules apply, which a referenced library never does.
 */
export function referencedHostTokens(
	host: VbaHostToken | undefined,
	references: readonly DeclaredReference[],
): VbaHostToken[] {
	return hostTokensForProject(host, references).filter((token) => token !== host);
}
