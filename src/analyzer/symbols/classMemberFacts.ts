// What a class's own code shows a member always gives (issue #414, each
// measured in Excel 16.0 through `Dim c As New Class1`):
//
//   Public M As Object, nothing in the class assigns it    Nothing
//   Function M() As Object that only sets it to Nothing    Nothing
//   Public M As Variant, nothing in the class assigns it   Empty
//   Property Get M() As Variant: M = 1, and nothing else   a scalar
//   Function M() As Variant: M = 1, and nothing else       a scalar
//   Function or Get As Variant that never assigns M        Empty
//
// Code outside the class can still assign a field through the instance; the
// rule that reads these facts follows the instance's own uses.

import { tokenizeCached } from '../lexer/tokenize';
import { firstTokenAtOrAfter, tokenName } from '../lexer/tokenHelpers';
import type { VbaToken } from '../lexer/tokenKinds';
import type { VbaSymbol } from './symbolModel';

const SCALAR_TYPES: ReadonlySet<string> = new Set([
	'string', 'boolean', 'date', 'byte', 'integer', 'long', 'longlong', 'longptr', 'single', 'double', 'currency', 'decimal',
]);

function normalizeType(type: string | undefined): string | undefined {
	return type?.trim().replace(/\s*\(\s*\)\s*$/, '').replace(/^vba\./i, '').toLowerCase() || undefined;
}

function isKnownScalarType(type: string): boolean {
	return SCALAR_TYPES.has(type);
}

export type ClassMemberValue = 'nothing' | 'empty' | 'scalar';

/** Words that make a body's flow more than one straight run. */
const FLOW_WORDS: ReadonlySet<string> = new Set(['if', 'select', 'for', 'do', 'while', 'with', 'goto', 'gosub', 'on', 'exit', 'resume', 'end']);

/** The member values the class's code decides, by lowercased member name. */
export function classMemberValues(source: string, children: readonly VbaSymbol[]): Map<string, ClassMemberValue> {
	const out = new Map<string, ClassMemberValue>();
	const toks = tokenizeCached(source).filter((tok) => tok.kind !== 'comment');
	const word = (tok: VbaToken | undefined): string => (tok?.rawText ?? '').toLowerCase();
	const name = (tok: VbaToken | undefined): string | undefined => tokenName(tok)?.toLowerCase();
	// Only the earliest/latest mention is needed to decide whether a name
	// occurs outside its declaration. Build that index once, not per field.
	const mentionsByName = new Map<string, { first: number; last: number }>();
	for (const tok of toks) {
		if (tok.kind !== 'identifier' && tok.kind !== 'bracketedIdentifier') { continue; }
		const lower = name(tok)!;
		const mentions = mentionsByName.get(lower);
		if (mentions) {
			mentions.last = tok.start;
		} else {
			mentionsByName.set(lower, { first: tok.start, last: tok.start });
		}
	}
	const moduleValues = new Map<string, VbaSymbol>();
	for (const child of children) {
		if (child.kind === 'moduleVariable' || child.kind === 'constant') { moduleValues.set(child.name.toLowerCase(), child); }
	}
	for (const symbol of children) {
		const lower = symbol.name.toLowerCase();
		const type = normalizeType(symbol.asType);
		if (symbol.kind === 'moduleVariable') {
			if (symbol.isArray || symbol.isAutoInstantiated) {
				continue;
			}
			// Any mention past the declaration, Me.M and a ByRef pass included,
			// may assign it.
			const mentions = mentionsByName.get(lower);
			const named = mentions !== undefined && (mentions.first < symbol.fullSpan.start || mentions.last >= symbol.fullSpan.end);
			if (named) {
				continue;
			}
			if (type === undefined || type === 'variant') {
				out.set(lower, 'empty');
			} else if (!isKnownScalarType(type)) {
				out.set(lower, 'nothing');
			}
			continue;
		}
		if (symbol.kind !== 'function' && symbol.kind !== 'propertyGet') {
			continue;
		}
		if (/\(\s*\)\s*$/.test(symbol.asType ?? '')) { continue; } // An array result is not an object reference.
		// The body's statements, the header line left out.
		const body: VbaToken[] = [];
		for (let i = firstTokenAtOrAfter(toks, symbol.nameSpan.end + 1); i < toks.length; i++) {
			const tok = toks[i];
			if (tok.start > symbol.fullSpan.end) { break; }
			if (tok.end <= symbol.fullSpan.end) { body.push(tok); }
		}
		const headerEnd = body.findIndex((tok) => tok.kind === 'newline');
		const statements = splitStatements(body.slice(headerEnd + 1));
		// The last statement is End Function or End Property.
		const inner = statements.filter((stmt) => !(word(stmt[0]) === 'end' && ['function', 'property'].includes(word(stmt[1]))));
		const mentions = inner.filter((stmt) => stmt.some((tok, i) => (tok.kind === 'identifier' || tok.kind === 'bracketedIdentifier') && name(tok) === lower && stmt[i - 1]?.rawText !== '.'));
		if (type !== undefined && type !== 'variant' && !isKnownScalarType(type)) {
			// An object result never set, or set only to Nothing, is Nothing.
			if (mentions.every((stmt) => stmt.length === 4 && word(stmt[0]) === 'set' && name(stmt[1]) === lower && stmt[2].rawText === '=' && word(stmt[3]) === 'nothing')) {
				out.set(lower, 'nothing');
			}
			continue;
		}
		// A Variant result nothing assigns is Empty (issue #414, measured).
		if ((type === undefined || type === 'variant') && mentions.length === 0) {
			out.set(lower, 'empty');
			continue;
		}
		if ((type === undefined || type === 'variant') && mentions.length === 1
			&& !inner.some((stmt) => FLOW_WORDS.has(word(stmt[0])))) {
			const stmt = mentions[0];
			const value = stmt.slice(2);
			const literal = value.length === 1 || (value.length === 2 && value[0].rawText === '-') ? value[value.length - 1] : undefined;
			const referenced = value.length === 1 ? name(value[0]) : undefined;
			const bound = referenced ? (symbol.children ?? []).find(child => child.name.toLowerCase() === referenced) ?? moduleValues.get(referenced) : undefined;
			const scalarVariable = bound && !bound.isArray && !/\(\s*\)\s*$/.test(bound.asType ?? '')
				&& isKnownScalarType(normalizeType(bound.asType) ?? '');
			if (name(stmt[0]) === lower && stmt[1]?.rawText === '=' && ((literal && ['integerLiteral', 'floatLiteral', 'stringLiteral'].includes(literal.kind)) || scalarVariable)) {
				out.set(lower, 'scalar');
			}
		}
	}
	return out;
}

/** Tokens split at line ends and colons into statements, empty ones dropped. */
function splitStatements(toks: readonly VbaToken[]): VbaToken[][] {
	const out: VbaToken[][] = [];
	let current: VbaToken[] = [];
	for (const tok of toks) {
		if (tok.kind === 'newline' || tok.rawText === ':') {
			if (current.length > 0) {
				out.push(current);
			}
			current = [];
			continue;
		}
		current.push(tok);
	}
	if (current.length > 0) {
		out.push(current);
	}
	return out;
}
