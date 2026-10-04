import type { Span } from '../parser/nodes';
import { tokenize } from '../lexer/tokenize';
import { firstTokenAtOrAfter, isDecimalLineNumber, tokenName } from '../lexer/tokenHelpers';
import { findIdentifierOccurrences, lineStartAtAnyBreak, lineEndAtOrAfter, stripVba, VBA_IDENTIFIER_PATTERN } from '../../vbaSourceScan';

/** A call's insertion point, preserving its original argument syntax and text. */
export interface CallSite {
	offset: number;
	argumentInsert: Span;
	/** Supply the new parameter name when extending a named-argument list. */
	argumentText: (value: string, parameterName?: string) => string;
	/** Arguments use a bracketed call-list form rather than a bare list. */
	bracketed: boolean;
	empty: boolean;
}
export interface CallSiteOptions {
	/** Ignore the procedure's own body and header. */
	skip?: Span;
	/** Optional source-binding filter; true call syntax bypasses function result variables. */
	accept?: (offset: number, call: boolean, qualifier: string | undefined, nameSpan: Span) => boolean;
}
const QUALIFIERS = new RegExp('(?:' + VBA_IDENTIFIER_PATTERN + '\\s*\\.\\s*)+$', 'u');
const DECLARATION = /\b(?:Sub|Function|Property\s+(?:Get|Let|Set)|Declare)\s+$|^\s*(?:Public|Private|Friend|Static)?\s*(?:Static\s+)?(?:Sub|Function|Property)\b/i;

/** Query-owned metadata over one immutable, bounded logical-line token stream.
	* Fragment lexing stays outside the shared module lexer cache. */
class CallLine {
	readonly tokens;
	readonly roots: number[] = [];
	readonly ends: number[] = [];
	readonly depths: number[] = [];
	readonly pairs = new Map<number, number>();
	readonly named = new Map<number, number[]>();
	readonly elses = new Map<number, number[]>();
	readonly typeNames = new Set<number>();
	constructor(readonly source: string, readonly start: number, readonly end: number) {
		this.tokens = tokenize(source.slice(start, end));
		let root = 0;
		const stack: number[] = [];
		const pendingTypeOf = new Map<number, number>();
		for (let i = 0; i < this.tokens.length; i++) {
			const token = this.tokens[i];
			this.roots[i] = root;
			this.depths[i] = stack.length;
			if (token.kind === 'colon' || token.kind === 'comment' || token.kind === 'newline') { root = i + 1; stack.length = 0; pendingTypeOf.clear(); }
			else if (token.rawText === '(') { stack.push(i); }
			else if (token.rawText === ')') { const open = stack.pop(); if (open !== undefined) { this.pairs.set(open, i); } }
			else if (token.kind === 'keyword' && /^typeof$/i.test(token.rawText)) { pendingTypeOf.set(stack.length, (pendingTypeOf.get(stack.length) ?? 0) + 1); }
			else if (token.kind === 'keyword' && /^is$/i.test(token.rawText) && (pendingTypeOf.get(stack.length) ?? 0) > 0) {
				pendingTypeOf.set(stack.length, pendingTypeOf.get(stack.length)! - 1);
				for (let name = i + 1; tokenName(this.tokens[name]) !== undefined; name += 2) {
					this.typeNames.add(name);
					if (this.tokens[name + 1]?.rawText !== '.') { break; }
				}
			}
			else if (token.kind === 'operator' && token.rawText === ':=') { this.add(this.named, stack.length, i); }
			else if (token.kind === 'keyword' && /^else$/i.test(token.rawText)
				&& !['.', '!'].includes(this.tokens[i - 1]?.rawText)) { this.add(this.elses, stack.length, i); }
		}
		let endAt = this.tokens.length;
		for (let i = this.tokens.length - 1; i >= 0; i--) {
			if (this.tokens[i].kind === 'colon' || this.tokens[i].kind === 'comment' || this.tokens[i].kind === 'newline') { endAt = i; }
			this.ends[i] = endAt;
		}
	}
	private add(map: Map<number, number[]>, depth: number, index: number): void {
		const values = map.get(depth);
		if (values) { values.push(index); } else { map.set(depth, [index]); }
	}
	first(map: ReadonlyMap<number, readonly number[]>, depth: number, start: number): number {
		const values = map.get(depth) ?? [];
		let low = 0, high = values.length;
		while (low < high) { const mid = low + Math.floor((high - low) / 2); if (values[mid] < start) { low = mid + 1; } else { high = mid; } }
		return values[low] ?? this.tokens.length;
	}
	site(offset: number, wanted: string, accept?: CallSiteOptions['accept']): CallSite | undefined {
		let index = firstTokenAtOrAfter(this.tokens, offset - this.start);
		const previous = this.tokens[index - 1];
		if (previous?.kind === 'bracketedIdentifier' && previous.start < offset - this.start && previous.end > offset - this.start) { index--; }
		const token = this.tokens[index];
		if (!token || token.start > offset - this.start || tokenName(token)?.toLowerCase() !== wanted) { return undefined; }
		let root = this.roots[index];
		if (index > root && isDecimalLineNumber(this.tokens[root])) { root++; }
		let header = root;
		while (/^(Public|Private|Friend|Static)$/i.test(this.tokens[header]?.rawText ?? '')) { header++; }
		if (/^(Sub|Function|Property|Declare)$/i.test(this.tokens[header]?.rawText ?? '')) { return undefined; }
		let first = index;
		while (first >= root + 2 && this.tokens[first - 1].rawText === '.' && tokenName(this.tokens[first - 2]) !== undefined
			&& (!/^(Then|Else|Call)$/i.test(this.tokens[first - 2].rawText) || this.tokens[first - 3]?.rawText === '.')) { first -= 2; }
		const leadingDot = first > root && this.tokens[first - 1].rawText === '.';
		const beforeIndex = first - (leadingDot ? 2 : 1);
		const before = beforeIndex >= root ? this.tokens[beforeIndex] : undefined;
		const explicit = /^Call$/i.test(before?.rawText ?? '');
		const bare = before === undefined || /^(Then|Else)$/i.test(before.rawText);
		const next = this.tokens[index + 1];
		if (this.typeNames.has(index) || next?.kind === 'colon' && index === root || next?.rawText === ':=' || before?.rawText === '!' || /^(GoTo|GoSub|Resume|AddressOf|New|As|Implements)$/i.test(before?.rawText ?? '')) { return undefined; }
		const qualifier = leadingDot ? '.' : first === index ? undefined : this.source.slice(this.start + this.tokens[first].start, this.start + this.tokens[index - 1].start).trim();
		if (accept && !accept(offset, bare || explicit || next?.rawText === '(', qualifier, {start: this.start + token.start, end: this.start + token.end})) { return undefined; }
		if (next?.rawText === '.' || next?.rawText === '=' && (bare || explicit || /^(Set|Let|For|LSet|RSet)$/i.test(before?.rawText ?? ''))) { return undefined; }
		const after = this.start + token.end;
		if (next?.rawText === '(') {
			const close = this.pairs.get(index + 1);
			if (close === undefined) { return undefined; }
			const empty = close === index + 2;
			if (!bare || empty) {
				const insert = this.start + this.tokens[close].start;
				const named = this.first(this.named, this.depths[index + 1] + 1, index + 2) < close;
				return { offset, argumentInsert: { start: insert, end: insert }, argumentText: (value, name) => (empty ? '' : ', ') + (named && name ? name + ':=' : '') + value, bracketed: true, empty };
			}
		}
		if (!bare && !explicit) { return bracketedEmpty(offset, after); }
		let end = this.ends[index];
		end = Math.min(end, this.first(this.elses, this.depths[index], index + 1));
		const empty = index + 1 >= end;
		if (explicit) { return empty ? bracketedEmpty(offset, after) : undefined; }
		const insert = empty ? after : this.start + this.tokens[end - 1].end;
		const named = this.first(this.named, this.depths[index], index + 1) < end;
		return { offset, argumentInsert: { start: insert, end: insert }, argumentText: (value, name) => (empty ? ' ' : ', ') + (named && name ? name + ':=' : '') + value, bracketed: false, empty };
	}
}
function bracketedEmpty(offset: number, after: number): CallSite {
	return { offset, argumentInsert: { start: after, end: after }, argumentText: value => '(' + value + ')', bracketed: true, empty: true };
}

/** A possible continuation marker; the lexer decides whether it belongs to trivia. */
function continuesAt(source: string, end: number): boolean {
	let at = end - 1;
	while (source[at] === ' ' || source[at] === '\t') { at--; }
	return source[at] === '_' && (source[at - 1] === ' ' || source[at - 1] === '\t');
}

export function callSitesOf(source: string, procedureName: string, options: CallSiteOptions = {}): CallSite[] {
	const out: CallSite[] = [];
	const wanted = procedureName.toLowerCase();
	let start = -1, end = -1, line: CallLine | undefined, seen = false;
	for (const occurrence of findIdentifierOccurrences(source, procedureName)) {
		const { offset } = occurrence;
		if (options.skip && offset >= options.skip.start && offset <= options.skip.end) { continue; }
		if (offset < start || offset >= end) {
			start = offset - occurrence.column;
			while (start > 0) {
				let previousEnd = start - 1;
				if (source[previousEnd] === '\n' && source[previousEnd - 1] === '\r') { previousEnd--; }
				if (!continuesAt(source, previousEnd)) { break; }
				start = lineStartAtAnyBreak(source, previousEnd);
			}
			end = lineEndAtOrAfter(source, offset);
			while (end < source.length && continuesAt(source, end)) {
				end = lineEndAtOrAfter(source, end + (source[end] === '\r' && source[end + 1] === '\n' ? 2 : 1));
			}
			// Before any string delimiter, an apostrophe starts an actual comment.
			// Keep numeric/comment lines on the common path without lexing their notes.
			if (start === offset - occurrence.column && !/[\r\n]/.test(source.slice(start, end))) {
				const text = source.slice(start, end), comment = text.indexOf("'");
				if (comment >= 0 && !text.slice(0, comment).includes('"')) { end = start + comment; }
			}
			line = undefined; seen = false;
			if (start !== offset - occurrence.column || /[\r\n]/.test(source.slice(start, end)) || /[:'#\[]|\b(?:Rem|Else|TypeOf)\b/i.test(source.slice(start, end))) { line = new CallLine(source, start, end); }
		} else if (seen && !line) { line = new CallLine(source, start, end); }
		seen = true;
		if (line) { const site = line.site(offset, wanted, options.accept); if (site) { out.push(site); } continue; }
		// The common single-call line needs no lexer or token-index allocation.
		const prefix = stripVba(source.slice(start, offset));
		if (DECLARATION.test(prefix)) { continue; }
		if (/(?:^|\bThen|\bElse|\bCall)\s*\.$/i.test(prefix.trimEnd())) { line = new CallLine(source, start, end); const site = line.site(offset, wanted, options.accept); if (site) { out.push(site); } continue; }
		const after = offset + occurrence.text.length;
		const rest = source.slice(after, end);
		const lead = /^[ \t]*/.exec(rest)?.[0] ?? '';
		const next = rest[lead.length];
		if (next === '.') { continue; }
		const context = prefix.replace(/^\s*\d+\s+/, '').replace(QUALIFIERS, '').trim();
		const bare = context === '' || /\b(?:Then|Else)$/i.test(context);
		if (/\b(?:GoTo|GoSub|Resume|AddressOf|New|As|Implements)$/i.test(context) || prefix.trimEnd().endsWith('!')) { continue; }
		const qualifier = QUALIFIERS.exec(prefix)?.[0].replace(/\s*\.\s*$/, '').trim() ?? (prefix.trimEnd().endsWith('.') ? '.' : undefined);
		if (options.accept && !options.accept(offset, bare || /^Call$/i.test(context) || next === '(', qualifier, {start: offset, end: after})) { continue; }
		if (next === '(') {
			// Numeric flat lists need no token index; complex arguments use the lexer.
			const flat = /^\(([ \t0-9,+\-*/\\.^&%!=<>]*)\)/.exec(rest.slice(lead.length));
			if (!flat) { line = new CallLine(source, start, end); const site = line.site(offset, wanted, options.accept); if (site) { out.push(site); } continue; }
			const empty = flat[1].trim() === '';
			if (!bare || empty) {
				const insert = after + lead.length + flat[0].length - 1;
				out.push({ offset, argumentInsert: { start: insert, end: insert }, argumentText: value => (empty ? '' : ', ') + value, bracketed: true, empty });
				continue;
			}
		}
		const empty = rest.trimEnd().length === 0;
		if (next === '=' && (bare || /^(Call|Set|Let|For|LSet|RSet)$/i.test(context))) { continue; }
		if (!bare) { if (!/^Call$/i.test(context) || empty) { out.push(bracketedEmpty(offset, after)); } continue; }
		const insert = empty ? after : after + rest.trimEnd().length;
		out.push({ offset, argumentInsert: { start: insert, end: insert }, argumentText: value => (empty ? ' ' : ', ') + value, bracketed: false, empty });
	}
	return out;
}
