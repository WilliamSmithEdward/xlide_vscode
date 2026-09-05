import type { Span } from '../parser/nodes';
import { findIdentifierOccurrences, stripVba } from '../../vbaSourceScan';

/**
 * Where a procedure is called, and where an argument would go.
 *
 * VBA has two call syntaxes and they take arguments differently: `Go 1, 2`
 * takes a bare list, `Call Go(1, 2)` and `x = Go(1)` take a bracketed one, and
 * `Go (1)` is a THIRD thing - a bare list whose single argument is
 * parenthesised, which passes it by value. A refactoring that adds an argument
 * has to write the form the call site already uses, so this reports which one
 * it found rather than rewriting every call into one shape.
 */

export interface CallSite {
	/** Offset of the called name. */
	offset: number;
	/** Where an added argument is inserted. */
	argumentInsert: Span;
	/** The text to insert there for a given argument, brackets and commas included. */
	argumentText: (value: string) => string;
	/** True when the call already brackets its arguments. */
	bracketed: boolean;
	/** True when the call passes no arguments at all. */
	empty: boolean;
}

export interface CallSiteOptions {
	/** A span to ignore - the procedure's own body, so its header is not a call. */
	skip?: Span;
	/**
	 * The module the procedure lives in. A qualified call (`Module.Go`) is
	 * still a call; an unqualified one in another module is too, since VBA
	 * resolves public procedures project-wide.
	 */
	qualifier?: string;
}

export function callSitesOf(
	source: string,
	procedureName: string,
	options: CallSiteOptions = {},
): CallSite[] {
	const out: CallSite[] = [];
	for (const occurrence of findIdentifierOccurrences(source, procedureName)) {
		const { offset } = occurrence;
		if (options.skip && offset >= options.skip.start && offset <= options.skip.end) {
			continue;
		}
		if (isDeclaration(source, offset)) {
			continue;
		}
		const site = siteAt(source, offset, procedureName.length);
		if (site) {
			out.push(site);
		}
	}
	return out;
}

/** The line declares the procedure rather than calling it. */
function isDeclaration(source: string, offset: number): boolean {
	const lineStart = source.lastIndexOf('\n', Math.max(offset - 1, 0)) + 1;
	const before = stripVba(source.slice(lineStart, offset));
	return /\b(?:Sub|Function|Property\s+(?:Get|Let|Set)|Declare)\s+$/i.test(before)
		|| /^\s*(?:Public|Private|Friend|Static)?\s*(?:Static\s+)?(?:Sub|Function|Property)\b/i.test(
			source.slice(lineStart, offset),
		);
}

function siteAt(source: string, offset: number, nameLength: number): CallSite | undefined {
	const after = offset + nameLength;
	const rest = source.slice(after);
	const lead = /^[ \t]*/.exec(rest)?.[0] ?? '';
	const at = after + lead.length;
	const next = source[at];

	// `x = Go` or `Go` on its own: an empty call, bracketed or not.
	if (next === '(') {
		const close = matchingBracket(source, at);
		if (close === -1) {
			return undefined;
		}
		const inner = source.slice(at + 1, close).trim();
		return {
			offset,
			argumentInsert: { start: close, end: close },
			argumentText: (value) => (inner === '' ? value : `, ${value}`),
			bracketed: true,
			empty: inner === '',
		};
	}
	// A member access or a declaration of something else: not this call.
	if (next === '.' || next === '=' && source[at + 1] !== '=') {
		// `Go = 1` inside Go assigns the return value, not a call.
		return undefined;
	}
	const lineEnd = endOfLine(source, at);
	const trailing = stripVba(source.slice(at, lineEnd)).trimEnd();
	return {
		offset,
		argumentInsert: { start: at, end: at + trailing.length },
		argumentText: (value) => (trailing === '' ? ` ${value}` : `${trailing}, ${value}`),
		bracketed: false,
		empty: trailing === '',
	};
}

function matchingBracket(source: string, open: number): number {
	let depth = 0;
	let inString = false;
	for (let i = open; i < source.length; i += 1) {
		const ch = source[i];
		if (ch === '"') { inString = !inString; continue; }
		if (inString) { continue; }
		if (ch === '\n') { return -1; }
		if (ch === '(') { depth += 1; }
		if (ch === ')') {
			depth -= 1;
			if (depth === 0) { return i; }
		}
	}
	return -1;
}

function endOfLine(source: string, offset: number): number {
	const at = source.indexOf('\n', offset);
	if (at === -1) { return source.length; }
	return source[at - 1] === '\r' ? at - 1 : at;
}
