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
import { tokenName } from '../../lexer/tokenHelpers';
import { tokenizeCached } from '../../lexer/tokenize';
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
	// The pass has the module's token stream already; a second full lex of
	// the module here was 2% of a large project's analysis (issue #139).
	const toks = tokenizeCached(source).filter((t) => t.kind !== 'comment' && t.kind !== 'newline');
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
 * The libraries the module names early bound, lowercased. Removing a
 * reference is the other half of this rule: a project can be told which of
 * its modules would stop compiling before the reference goes, which is what
 * the VBE's own Tools > References dialog never says.
 */
export function librariesNamedIn(source: string): Set<string> {
	const out = new Set<string>();
	for (const found of qualifiedNamesIn(source)) {
		out.add(found.library.toLowerCase());
	}
	return out;
}

/** The Scripting library's types a module names unqualified, which no default reference brings. */
const SCRIPTING_TYPES: ReadonlySet<string> = new Set(['dictionary', 'filesystemobject', 'textstream']);

/**
 * Module rule: an early-bound Scripting type in a project whose references
 * are known and do not include the Scripting Runtime. `Dim d As
 * Scripting.Dictionary` and `Dim d As New Dictionary` then do not compile,
 * "User-defined type not defined" (issue #349, measured in Excel 16.0).
 * Silent when the references are not known, and for a name the project
 * declares itself. Reported once per module, as a missing library is.
 */
export function checkMissingScriptingReference(
	source: string,
	referencedLibraries: readonly string[] | undefined,
	projectTypes: ReadonlySet<string>,
	push: PushFn,
): void {
	if (referencedLibraries === undefined || referencedLibraries.some((name) => name.toLowerCase() === 'scripting')) {
		return;
	}
	const toks = tokenizeCached(source).filter((t) => t.kind !== 'comment' && t.kind !== 'newline');
	for (let i = 0; i + 1 < toks.length; i++) {
		const word = toks[i].rawText.toLowerCase();
		if (word !== 'as' && word !== 'new') {
			continue;
		}
		let at = i + 1;
		if (word === 'as' && toks[at]?.rawText.toLowerCase() === 'new') {
			at++;
		}
		const qualified = toks[at]?.rawText.toLowerCase() === 'scripting' && toks[at + 1]?.rawText === '.' && tokenName(toks[at + 2]) !== undefined;
		const name = qualified ? tokenName(toks[at + 2])! : tokenName(toks[at]);
		const lower = name?.toLowerCase();
		if (!name || !lower || (!qualified && (!SCRIPTING_TYPES.has(lower) || projectTypes.has(lower) || toks[at + 1]?.rawText === '.'))) {
			continue;
		}
		push(
			'missingLibraryReference',
			`'${name}' is the Scripting Runtime's, which this project does not reference. Add a reference to Microsoft Scripting Runtime, or bind late: `
			+ `Dim x As Object: Set x = CreateObject("Scripting.${name}"). This is a VBE compile error: User-defined type not defined.`,
			{ start: toks[at].start, end: (qualified ? toks[at + 2] : toks[at]).end },
			{ addLibraryReference: { library: 'scripting' } },
		);
		return;
	}
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
	/** The project's module names, lowercased: a module named Word is called as Word.Hi (issue #357). */
	projectModules: ReadonlySet<string> = new Set(),
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
		if (library === undefined || present.has(lower) || projectModules.has(lower) || seen.has(lower)) { continue; }
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
