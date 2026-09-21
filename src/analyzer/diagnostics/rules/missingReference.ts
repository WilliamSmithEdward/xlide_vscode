// Rule family: naming another application's library without referencing it.
//
// `Dim xl As Excel.Application` in a Word document compiles only when the
// project references the Excel type library. Without it the VBE refuses the
// declaration outright - "User-defined type not defined" - and the whole
// project stops compiling, so this is a real compile error rather than a
// style note.
//
// EARLY BINDING ONLY. Late binding needs no reference at all:
//
//     Dim xl As Object
//     Set xl = CreateObject("Excel.Application")
//
// names nothing from the library, resolves through IDispatch at run time,
// and is the usual way to drive another application without one. So the rule
// fires on a name the compiler has to resolve - a qualified type in an As
// clause, a New, or a qualified constant - and never on a string.
//
// It also fires only on the QUALIFIED spelling. An unqualified `Dim wb As
// Workbook` in a Word project is indistinguishable from a project class that
// does not exist yet, and guessing between the two would put a reference
// suggestion on ordinary broken code.

import type { HostObjectModel } from '../../host/excelObjectModel';
import { HOST_LIBRARY_NAMES } from '../../host/hostLibraries';
import { statementTokens, tokenName } from '../../lexer/tokenHelpers';
import type { Span } from '../../parser/nodes';
import type { PushFn } from '../analysisContext';

/** The libraries a project can be given a reference to, lowercased. */
const ADDABLE = new Map<string, string>(
	(['excel', 'word', 'powerpoint', 'access'] as const)
		.map((token) => [HOST_LIBRARY_NAMES[token].toLowerCase(), HOST_LIBRARY_NAMES[token]]),
);

/** Which libraries the model can already answer for, from its own type keys. */
function librariesInModel(model: HostObjectModel | undefined): Set<string> {
	const out = new Set<string>();
	for (const qualified of Object.keys(model?.types ?? {})) {
		const dot = qualified.indexOf('.');
		if (dot > 0) { out.add(qualified.slice(0, dot).toLowerCase()); }
	}
	return out;
}

/**
 * Every `Library.Member` in the module where the compiler has to resolve
 * `Library`: in an `As` clause, after `New`, or standing as a value.
 *
 * The whole module is scanned rather than each procedure's statements,
 * because `Dim xl As Excel.Application` - the commonest early binding there
 * is, and a module-level `Private mApp As Excel.Application` with it - is a
 * declaration, and the per-statement walk does not reach declarations.
 *
 * A string literal is one token, so `CreateObject("Excel.Application")` is
 * not a match: late binding names nothing the compiler has to resolve.
 */
function qualifiedNamesIn(source: string): Array<{ library: string; span: Span }> {
	const toks = statementTokens(source, { start: 0, end: source.length });
	const out: Array<{ library: string; span: Span }> = [];
	for (let i = 0; i < toks.length - 2; i++) {
		const library = tokenName(toks[i]);
		if (toks[i].kind !== 'identifier' || !library) { continue; }
		if (toks[i + 1].rawText !== '.') { continue; }
		if (toks[i + 2].kind !== 'identifier' && toks[i + 2].kind !== 'keyword') { continue; }
		// A member access further along a chain (`a.b.c`) is not a library
		// qualifier: `b` there is a member of whatever `a` is.
		if (i > 0 && toks[i - 1].rawText === '.') { continue; }
		out.push({ library, span: { start: toks[i].start, end: toks[i + 2].end } });
	}
	return out;
}

/**
 * Module rule: a type or constant qualified with an Office library the
 * project does not reference.
 *
 * Reported once per library per module. A project missing a reference names
 * it on every line that uses it, and one diagnostic per line would bury the
 * module in the same message with the same one fix.
 */
export function checkMissingLibraryReference(
	source: string,
	model: HostObjectModel | undefined,
	push: PushFn,
): void {
	const present = librariesInModel(model);
	// Nothing is known about any library, so nothing can be said about one
	// being absent. A caller that passes no model gets silence.
	if (present.size === 0) {
		return;
	}
	const seen = new Set<string>();
	for (const found of qualifiedNamesIn(source)) {
		const lower = found.library.toLowerCase();
		const library = ADDABLE.get(lower);
		if (library === undefined || present.has(lower) || seen.has(lower)) { continue; }
		seen.add(lower);
		push(
			'missingLibraryReference',
			`'${library}' is not referenced by this project, so ${library}.* cannot be resolved. `
			+ `Add a reference to the ${library} object library, or use late binding: `
			+ `Dim x As Object: Set x = CreateObject("${library}.Application").`,
			found.span,
			{ addLibraryReference: { library: lower } },
		);
	}
}
