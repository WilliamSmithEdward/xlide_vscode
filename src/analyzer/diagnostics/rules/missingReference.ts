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

import type { ReferencedProjectSurface } from '../../symbols/referencedProjects';
import { buildModuleSymbols } from '../../symbols/buildModuleSymbols';
import { isProcedureKind, type ModuleSymbols, type VbaSymbol } from '../../symbols/symbolModel';
import { resolveBareIdentifierBinding } from '../../symbols/nameResolution';

interface SourceQualifierBindings {
	symbols: ModuleSymbols;
	projectVisibleSymbols?: readonly VbaSymbol[];
}

import type { HostObjectModel } from '../../host/excelObjectModel';
import { HOST_LIBRARY_NAMES } from '../../host/hostLibraries';
import { tokenName } from '../../lexer/tokenHelpers';
import { tokenizeCached } from '../../lexer/tokenize';
import type { Span } from '../../parser/nodes';
import type { PushFn } from '../analysisContext';

/** The libraries a project can be given a reference to, lowercased. */
const ADDABLE = new Map<string, string>(
	(['excel', 'word', 'powerpoint', 'access', 'outlook'] as const)
		.map((token) => [HOST_LIBRARY_NAMES[token].toLowerCase(), HOST_LIBRARY_NAMES[token]]),
);

// Models are immutable for their identity; a reference change replaces the model.
const MODEL_LIBRARIES = new WeakMap<HostObjectModel, ReadonlySet<string>>();

/** Which libraries the model can already answer for, from its own type keys. */
function librariesInModel(model: HostObjectModel | undefined): ReadonlySet<string> {
	if (!model) { return new Set(); }
	const cached = MODEL_LIBRARIES.get(model);
	if (cached) { return cached; }
	const out = new Set<string>();
	for (const qualified of Object.keys(model.types)) {
		const dot = qualified.indexOf('.');
		if (dot > 0) { out.add(qualified.slice(0, dot).toLowerCase()); }
	}
	MODEL_LIBRARIES.set(model, out);
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
function qualifiedNamesIn(source: string, bindings?: SourceQualifierBindings): Array<{ library: string; span: Span; typeQualifier: boolean }> {
	// The pass has the module's token stream already; a second full lex of
	// the module here was 2% of a large project's analysis (issue #139).
	const tokens = tokenizeCached(source);
	const toks = tokens.filter((t) => t.kind !== 'comment' && t.kind !== 'newline');
	const out: Array<{ library: string; span: Span; typeQualifier: boolean }> = [];
	const procedures = bindings?.symbols.root.children?.filter(symbol => isProcedureKind(symbol.kind)) ?? [];
	// Direct assignments introduce implicit locals without Option Explicit. Keep
	// them scoped to their procedure; a member write does not introduce its root.
	const implicitWrites = new Map<VbaSymbol, Set<string>>();
	let writeProcedureIndex = 0;
	let previous = '';
	for (let i = 0; i < tokens.length; i++) {
		const token = tokens[i];
		if (token.kind === 'comment') { continue; }
		while (writeProcedureIndex < procedures.length && procedures[writeProcedureIndex].fullSpan.end < token.start) { writeProcedureIndex++; }
		const procedure = procedures[writeProcedureIndex];
		if (procedure && procedure.fullSpan.start <= token.start && token.kind === 'identifier'
			&& tokens[i + 1]?.rawText === '=' && ['', 'set', 'let', 'then', 'else'].includes(previous)) {
			let names = implicitWrites.get(procedure);
			if (!names) { names = new Set(); implicitWrites.set(procedure, names); }
			names.add(tokenName(token)!.toLowerCase());
		}
		previous = token.kind === 'newline' || token.kind === 'colon' ? '' : token.rawText.toLowerCase();
	}
	let procedureIndex = 0;
	// Where each logical line ends; a ` _` continuation is trivia, not one.
	const newlines = tokens.filter((t) => t.kind === 'newline').map((t) => t.start);
	let nextNewline = 0;
	for (let i = 0; i < toks.length - 2; i++) {
		const library = tokenName(toks[i]);
		if (toks[i].kind !== 'identifier' || !library) { continue; }
		if (toks[i + 1].rawText !== '.') { continue; }
		// The dot must be on the same logical line. Inside a With block,
		// `With New Collection` and `.Add 1` below it are two statements,
		// not `Collection.Add`.
		while (nextNewline < newlines.length && newlines[nextNewline] < toks[i].end) { nextNewline++; }
		if (nextNewline < newlines.length && newlines[nextNewline] < toks[i + 1].start) { continue; }
		if (toks[i + 2].kind !== 'identifier' && toks[i + 2].kind !== 'keyword') { continue; }
		// A member access further along a chain (`a.b.c`) is not a library
		// qualifier: `b` there is a member of whatever `a` is.
		if (i > 0 && toks[i - 1].rawText === '.') { continue; }
		const typeQualifier = ['as', 'new', 'implements'].includes(toks[i - 1]?.rawText.toLowerCase() ?? '');
		if (bindings && !typeQualifier) {
			while (procedureIndex < procedures.length && procedures[procedureIndex].fullSpan.end < toks[i].start) { procedureIndex++; }
			const procedure = procedures[procedureIndex];
			const binding = resolveBareIdentifierBinding({currentModule: bindings.symbols, projectVisibleSymbols: bindings.projectVisibleSymbols,
				enclosingProcedure: procedure && procedure.fullSpan.start <= toks[i].start ? procedure : undefined,
				name: library, context: 'memberReceiver', offset: toks[i].start});
			if (binding.scope !== 'unresolved' || (procedure && implicitWrites.get(procedure)?.has(library.toLowerCase()))) { continue; }
		}
		out.push({ library, typeQualifier, span: { start: toks[i].start, end: toks[i + 2].end } });
	}
	return out;
}

/**
 * The libraries the module names early bound, lowercased. Removing a
 * reference is the other half of this rule: a project can be told which
 * modules may depend on it before the reference goes, which is what
 * the VBE's own Tools > References dialog never says.
 */
export function librariesNamedIn(source: string, bindings?: SourceQualifierBindings, projectModules: ReadonlySet<string> = new Set()): Set<string> {
	const out = new Set<string>();
	for (const found of qualifiedNamesIn(source, bindings ?? {symbols: buildModuleSymbols('', 'standard', source)})) {
		if (!projectModules.has(found.library.toLowerCase())) { out.add(found.library.toLowerCase()); }
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

/** Where the project is not known well enough to judge a qualifier. */
const NO_JUDGEMENT: ReadonlySet<string> = new Set();

/** What the project knows of itself, for judging a qualifier no library XLIDE can add spells. */
export interface ProjectNames {
	/** The libraries the dir stream records; undefined where they were not read. */
	referencedLibraries?: readonly string[];
	/** The project's own name, `VBAProject` or `Project` or an Access database's. */
	projectName?: string;
	/** Every module of the project, lowercased; undefined without the project. */
	projectModuleNames?: ReadonlySet<string>;
	/** What each referenced VBA project XLIDE read offers this one, its types included. */
	referencedProjectSurfaces?: readonly ReferencedProjectSurface[];
}

/**
 * The types a read referenced project offers by name, lowercased, or
 * undefined for a qualifier that is no such project (issue #1503): its
 * exposed classes and the Public Types and Enums of its visible standard
 * modules. A class it does not expose is no type of the referring project:
 * `Dim t As ProbeLibrary.Thing` is "User-defined type not defined".
 */
function referencedProjectTypes(qualifier: string, project: ProjectNames | undefined): ReadonlySet<string> | undefined {
	const surface = project?.referencedProjectSurfaces?.find((candidate) => candidate.name.toLowerCase() === qualifier);
	return surface && new Set(surface.modules.flatMap((module) => module.types.map((type) => type.toLowerCase())));
}

/**
 * The names that qualify a type in this project, lowercased, or undefined
 * when the project is not known well enough to say a name does not (issue
 * #1413, measured in Excel, Word, PowerPoint and Access 16.0): every
 * library the dir stream records, which it records by the name the library
 * gives itself; VBA and the host's own library, which it never records; the
 * project's name and its modules. The Scripting Runtime is
 * checkMissingScriptingReference's.
 */
function knownQualifiers(model: HostObjectModel | undefined, projectModules: ReadonlySet<string>, project: ProjectNames | undefined): ReadonlySet<string> | undefined {
	if (!project?.referencedLibraries || !project.projectName || !project.projectModuleNames) {
		return undefined;
	}
	return new Set([
		'vba', 'scripting', project.projectName, model?.hostName ?? 'Excel', ...project.referencedLibraries,
		...project.projectModuleNames, ...projectModules,
	].map((name) => name.toLowerCase()));
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
	bindings?: SourceQualifierBindings,
	project?: ProjectNames,
): void {
	const present = librariesInModel(model);
	// Nothing is known about any library, so nothing can be said about one
	// being absent. A caller that passes no model gets silence.
	if (present.size === 0) {
		return;
	}
	const seen = new Set<string>();
	let qualifiers: ReadonlySet<string> | undefined;
	const tokens = tokenizeCached(source).filter(token => token.kind !== 'comment' && token.kind !== 'newline');
	const explicit = tokens.some((token, i) => token.rawText.toLowerCase() === 'option' && tokens[i + 1]?.rawText.toLowerCase() === 'explicit');
	for (const found of qualifiedNamesIn(source, bindings ?? {symbols: buildModuleSymbols('', 'standard', source)})) {
		// Without Option Explicit an unknown value receiver is an implicit Variant,
		// not a missing-library compile error (native Excel control).
		if (!found.typeQualifier && !explicit) { continue; }
		const lower = found.library.toLowerCase();
		const referencedTypes = found.typeQualifier ? referencedProjectTypes(lower, project) : undefined;
		if (referencedTypes) {
			const typeName = source.slice(found.span.start, found.span.end).split('.').pop()?.trim().replace(/^\[(.*)\]$/, '$1') ?? '';
			if (!referencedTypes.has(typeName.toLowerCase())) {
				push(
					'missingLibraryReference',
					`'${found.library}' is a VBA project this project references, and it offers no type '${typeName}': a class it does not expose, or a Type or Enum in a module it keeps private, cannot be named here. This is a VBE compile error: User-defined type not defined.`,
					found.span,
				);
			}
			continue;
		}
		const library = ADDABLE.get(lower);
		if (library === undefined) {
			// A library XLIDE cannot add: judged only in a type position, and
			// only where the project's references and modules are known.
			qualifiers ??= knownQualifiers(model, projectModules, project) ?? NO_JUDGEMENT;
			if (found.typeQualifier && qualifiers !== NO_JUDGEMENT && !qualifiers.has(lower) && !seen.has(lower)) {
				seen.add(lower);
				push(
					'missingLibraryReference',
					`'${found.library}' is no library this project references and none of its modules, so ${found.library}.* cannot be resolved. `
					+ 'Add a reference to the library that defines it, or use late binding: '
					+ `Dim x As Object: Set x = CreateObject("${found.library}.<class>"). This is a VBE compile error: User-defined type not defined.`,
					found.span,
				);
			}
			continue;
		}
		if (present.has(lower) || projectModules.has(lower) || seen.has(lower)) { continue; }
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
