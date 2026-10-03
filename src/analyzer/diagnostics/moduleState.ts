// Module-level and class-level state the runtime rules can rely on (issue
// #241). A module variable that no statement ever writes keeps its initial
// value everywhere: an object is Nothing, a dynamic array has no elements, a
// number is 0, a String is "" and a Variant is Empty. Measured in Excel 16.0
// (build 20326): `Private m As Collection` then `m.Count` raises 91, and
// `Private z As Long` then `10 / z` raises 11.
//
// What counts as a write is read from the tokens alone and errs toward
// writing: an assignment's target, any name in a Set, ReDim, Erase, For,
// Input, Get, Line Input, LSet, RSet or Mid statement, and a whole name passed
// to anything but a VBA library function, which may take it ByRef. A Private
// variable needs only its own module; a Public one needs every module of the
// project, which the project index supplies as `projectWrittenNames`.

import { topLevelEqualsIndex } from '../lexer/tokenHelpers';
import { tokenizeCached } from '../lexer/tokenize';
import type { VbaToken } from '../lexer/tokenKinds';
import type { ProcedureNode } from '../parser/nodes';
import { resolveRuntimeFunction } from '../runtime/vbaRuntime';
import type { buildModuleSymbols } from '../symbols/buildModuleSymbols';
import type { VbaSymbol } from '../symbols/symbolModel';
import { procedureSymbolFor } from './analysisContext';
import { tokenName, tokenText } from './walker';

type ModuleSymbols = ReturnType<typeof buildModuleSymbols>;

/** Statement heads that write every name they mention. */
const WRITING_HEADS: ReadonlySet<string> = new Set(['set', 'let', 'redim', 'erase', 'for', 'input', 'get', 'line', 'lset', 'rset', 'mid', 'mid$']);

/** Statement heads that declare rather than run. */
const DECLARING_HEADS: ReadonlySet<string> = new Set([
	'dim', 'private', 'public', 'global', 'friend', 'static', 'const', 'sub', 'function', 'property',
	'declare', 'type', 'enum', 'end', 'option', 'attribute', 'implements', 'event',
]);

/** Heads after which an `=` compares rather than assigns. */
const COMPARING_HEADS: ReadonlySet<string> = new Set(['if', 'elseif', 'while', 'do', 'loop', 'until', 'case', 'select', 'debug', 'print', 'return', 'call']);

/**
 * The lowercased names a module's code may write, by the rules at the top
 * of this file. Over-reporting a write only keeps a rule quiet.
 */
export function writtenNamesIn(source: string): ReadonlySet<string> {
	const written = new Set<string>();
	const procedures = procedureNamesIn(source);
	let statement: VbaToken[] = [];
	const flush = (): void => {
		for (const segment of segmentsOf(statement)) {
			markWrites(segment, procedures, written);
		}
		statement = [];
	};
	for (const tok of tokenizeCached(source)) {
		if (tok.kind === 'newline' || tok.kind === 'colon') {
			flush();
		} else if (tok.kind !== 'comment') {
			statement.push(tok);
		}
	}
	flush();
	return written;
}

/** The procedures a module declares, which shadow the VBA library's functions. */
function procedureNamesIn(source: string): ReadonlySet<string> {
	const out = new Set<string>();
	const toks = tokenizeCached(source).filter((tok) => tok.kind !== 'comment' && tok.kind !== 'newline');
	for (let i = 0; i + 1 < toks.length; i++) {
		const word = tokenText(toks[i]);
		if ((word === 'sub' || word === 'function') && tokenText(toks[i - 1]) !== 'end' && tokenText(toks[i - 1]) !== 'exit') {
			const name = tokenName(toks[i + 1])?.toLowerCase();
			if (name) {
				out.add(name);
			}
		}
	}
	return out;
}

/** A statement, and each statement a single-line If runs after Then or Else. */
function segmentsOf(statement: readonly VbaToken[]): VbaToken[][] {
	const out: VbaToken[][] = [];
	let current: VbaToken[] = [];
	for (const tok of statement) {
		const word = tokenText(tok);
		if (word === 'then' || word === 'else') {
			out.push(current);
			current = [];
			continue;
		}
		current.push(tok);
	}
	out.push(current);
	return out;
}

function markWrites(segment: readonly VbaToken[], procedures: ReadonlySet<string>, written: Set<string>): void {
	// A line number or label leads some statements.
	const start = segment[0]?.kind === 'integerLiteral' ? 1 : 0;
	const toks = segment.slice(start);
	const head = tokenText(toks[0]);
	if (!head || DECLARING_HEADS.has(head)) {
		return;
	}
	const markAll = (from: number, to: number): void => {
		for (let i = from; i < to; i++) {
			const name = tokenName(toks[i])?.toLowerCase();
			if (name) {
				written.add(name);
			}
		}
	};
	if (WRITING_HEADS.has(head)) {
		markAll(0, toks.length);
		return;
	}
	const equals = topLevelEqualsIndex(toks);
	const assignment = equals > 0 && !COMPARING_HEADS.has(head);
	if (assignment) {
		markAll(0, equals);
	}
	// A whole name passed to a call may come back changed (ByRef).
	for (let i = assignment ? equals + 1 : 1; i < toks.length; i++) {
		const name = tokenName(toks[i])?.toLowerCase();
		if (!name || toks[i - 1]?.rawText === '.' || toks[i - 1]?.rawText === '!') {
			continue;
		}
		const next = toks[i + 1]?.rawText;
		if (next !== undefined && next !== ',' && next !== ')') {
			continue;
		}
		const previous = toks[i - 1];
		const callee = calleeOf(toks, i);
		if (callee === undefined) {
			// At the top level: an argument of a call statement, `Fill s`.
			const callStatement = !assignment && !COMPARING_HEADS.has(head)
				&& (previous?.rawText === ',' || (i === 1 && tokenName(previous) !== undefined) || (previous?.kind === 'identifier' && toks[i - 2]?.rawText === '.'));
			if (callStatement) {
				written.add(name);
			}
			continue;
		}
		if (previous?.rawText !== '(' && previous?.rawText !== ',') {
			continue;
		}
		const library = !procedures.has(callee) && resolveRuntimeFunction(callee)?.kind === 'function';
		if (!library) {
			written.add(name);
		}
	}
}

/** The name before the parenthesis that encloses `toks[at]`, or undefined at the top level. */
function calleeOf(toks: readonly VbaToken[], at: number): string | undefined {
	let depth = 0;
	for (let j = at - 1; j >= 0; j--) {
		if (toks[j].rawText === ')') {
			depth++;
		} else if (toks[j].rawText === '(') {
			if (depth === 0) {
				return tokenName(toks[j - 1])?.toLowerCase() ?? '';
			}
			depth--;
		}
	}
	return undefined;
}

const PROJECT_WRITES = new WeakMap<ModuleSymbols, ReadonlySet<string>>();
const MODULE_WRITES = new WeakMap<ModuleSymbols, ReadonlySet<string>>();

/**
 * Records, for one analysis of a module, what the rest of the project may
 * write. Without it a Public variable is never taken as unchanged.
 */
export function rememberProjectWrittenNames(symbols: ModuleSymbols, names: ReadonlySet<string> | undefined): void {
	if (names) {
		PROJECT_WRITES.set(symbols, names);
	}
}

/**
 * The module's variables that nothing writes, by lowercased name. A Private
 * or Dim variable needs only this module; a Public or Global one needs the
 * whole project. `As New` and fixed-length Strings are left out.
 */
export function untouchedModuleVariables(source: string, symbols: ModuleSymbols): ReadonlyMap<string, VbaSymbol> {
	let writes = MODULE_WRITES.get(symbols);
	if (!writes) {
		writes = writtenNamesIn(source);
		MODULE_WRITES.set(symbols, writes);
	}
	const project = PROJECT_WRITES.get(symbols);
	const out = new Map<string, VbaSymbol>();
	for (const child of symbols.root.children ?? []) {
		if (child.kind !== 'moduleVariable' || child.isAutoInstantiated || child.fixedLength !== undefined) {
			continue;
		}
		const lower = child.name.toLowerCase();
		const shared = child.visibility === 'Public' || child.visibility === 'Global';
		if (writes.has(lower) || (shared && (!project || project.has(lower)))) {
			continue;
		}
		out.set(lower, child);
	}
	return out;
}

/** {@link untouchedModuleVariables} less those a procedure's own local or parameter hides. */
export function untouchedModuleVariablesIn(source: string, symbols: ModuleSymbols, proc: ProcedureNode): ReadonlyMap<string, VbaSymbol> {
	const all = untouchedModuleVariables(source, symbols);
	if (all.size === 0) {
		return all;
	}
	const hidden = new Set((procedureSymbolFor(symbols, proc)?.children ?? []).map((child) => child.name.toLowerCase()));
	for (const param of proc.params) {
		hidden.add(param.name.toLowerCase());
	}
	hidden.add(proc.name.toLowerCase());
	const out = new Map(all);
	for (const lower of hidden) {
		out.delete(lower);
	}
	return out;
}
