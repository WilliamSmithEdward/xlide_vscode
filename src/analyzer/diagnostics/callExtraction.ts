// Call-statement extraction and arity validation shared by diagnostics rules.
//
// Extracted verbatim from `analyzeModule.ts`: the CallArguments model, the
// bare/qualified call-statement extractors, top-level argument-slot splitting,
// and the arity validator with its quick-fix placeholder data builders. The
// callable signature shapes (CallableTypeSignature et al.) live here because
// extraction and validation define the call contract; the signature *tables*
// are built by the type-inference module.

import { tokenize } from '../lexer/tokenize';
import type { VbaToken } from '../lexer/tokenKinds';
import type { Span } from '../parser/nodes';
import { bareCallStatementTarget as callStatementTarget } from '../call/callContext';
import { qualifiedProcedureKey } from '../symbols/symbolModel';
import type { PushFn, VbaDiagnosticData } from './analysisContext';
import {
	firstExecutableTokenIndex,
	stripHeaderBrackets,
	tokenName,
	tokenText,
} from './walker';

/** A resolved call statement: callee plus its top-level argument slots. */
export interface CallArguments {
	/** Callee identifier text. */
	name: string;
	/** Optional module qualifier for `ModuleName.MemberName` calls. */
	qualifier?: string;
	/** Lowercased signature lookup key; defaults to lowercased `name`. */
	lookupKey?: string;
	/** Absolute span of the callee identifier. */
	nameSpan: Span;
	/** True for the explicit `Call name...` form. */
	explicitCall?: boolean;
	/**
	 * Top-level, comma-separated argument groups. An empty list means no
	 * arguments were supplied; an empty inner array is an omitted positional
	 * argument (`Foo 1, , 3`).
	 */
	slots: VbaToken[][];
	/** Absolute spans for each argument slot; empty slots use the separator span. */
	slotSpans?: Span[];
	/** Absolute offset of the statement slice the slot tokens are relative to. */
	sliceStart: number;
}

export interface CallableParamType {
	name: string;
	type?: string;
	optional: boolean;
	paramArray: boolean;
	/** True when the parameter is declared as an array (`value() As T`). */
	isArray?: boolean;
	byRef?: boolean;
}

export interface CallableTypeSignature {
	name: string;
	params: CallableParamType[];
	returnType?: string;
}

export interface InferredArgumentType {
	type: string;
	label: string;
	span: Span;
	stringValue?: string;
	numericValue?: number;
	numericText?: string;
}

/**
 * If the statement spanning `span` is a bare call statement, returns the callee
 * and its top-level argument slots; otherwise undefined. Reuses
 * {@link callStatementTarget} for the safe call-detection gating, then peels off
 * the argument region (the parenless tail, or the contents of the `Call`
 * statement's parentheses).
 */
export function extractCall(source: string, span: Span): CallArguments | undefined {
	const hit = callStatementTarget(source, span);
	if (!hit) {
		return undefined;
	}
	const sliceStart = span.start;
	const toks = tokenize(source.slice(span.start, span.end)).filter(
		(t) => t.kind !== 'comment' && t.kind !== 'newline',
	);
	const startIndex = firstExecutableTokenIndex(toks);
	const relCalleeStart = hit.span.start - sliceStart;
	const calleeIdx = toks.findIndex((t) => t.start === relCalleeStart);
	if (calleeIdx < 0) {
		return undefined;
	}

	const explicitCall = tokenText(toks[startIndex]) === 'call';
	const next = toks[calleeIdx + 1];
	let argToks: VbaToken[];
	if (explicitCall) {
		if (next && next.kind === 'punctuation' && next.rawText === '(') {
			// Collect the tokens strictly inside the call's parentheses.
			let depth = 0;
			let closed = false;
			const inner: VbaToken[] = [];
			for (let k = calleeIdx + 1; k < toks.length; k++) {
				const t = toks[k];
				if (t.kind === 'punctuation' && t.rawText === '(') {
					depth++;
					if (depth === 1) {
						continue; // skip the opening paren itself
					}
				} else if (t.kind === 'punctuation' && t.rawText === ')') {
					depth--;
					if (depth === 0) {
						closed = true;
						break;
					}
				}
				if (depth >= 1) {
					inner.push(t);
				}
			}
			if (!closed) {
				return undefined; // unbalanced - the parentheses rule reports this
			}
			argToks = inner;
		} else {
			argToks = []; // `Call Foo` with no parameter list
		}
	} else {
		argToks = toks.slice(calleeIdx + 1);
	}

	const split = argToks.length === 0 ? emptyArgSplit() : splitArgSlots(argToks, sliceStart);
	return {
		name: hit.name,
		nameSpan: hit.span,
		explicitCall,
		slots: split.slots,
		slotSpans: split.spans,
		sliceStart,
	};
}

/**
 * Extracts a module-qualified call statement (`ModuleName.Procedure ...`) only
 * when the project signature map proves that `ModuleName.Procedure` is an
 * exported project procedure. This keeps host/object member calls out of the
 * arity/type validator.
 */
export function extractQualifiedCall(
	source: string,
	span: Span,
	moduleSignatures: ReadonlyMap<string, CallableTypeSignature>,
): CallArguments | undefined {
	const sliceStart = span.start;
	const toks = tokenize(source.slice(span.start, span.end)).filter(
		(t) => t.kind !== 'comment' && t.kind !== 'newline',
	);
	if (toks.length === 0) {
		return undefined;
	}

	let qualifierIdx = firstExecutableTokenIndex(toks);
	const explicitCall = tokenText(toks[qualifierIdx]) === 'call';
	if (explicitCall) {
		qualifierIdx += 1;
	}
	const qualifier = tokenName(toks[qualifierIdx]);
	const dot = toks[qualifierIdx + 1];
	const callee = toks[qualifierIdx + 2];
	const name = callee ? tokenName(callee) : undefined;
	if (!qualifier || dot?.rawText !== '.' || !name) {
		return undefined;
	}
	const lookupKey = qualifiedProcedureKey(qualifier, name);
	if (!moduleSignatures.has(lookupKey)) {
		return undefined;
	}

	const next = toks[qualifierIdx + 3];
	let argToks: VbaToken[];
	if (explicitCall) {
		if (next && next.kind === 'punctuation' && next.rawText === '(') {
			let depth = 0;
			let closed = false;
			const inner: VbaToken[] = [];
			for (let k = qualifierIdx + 3; k < toks.length; k++) {
				const t = toks[k];
				if (t.kind === 'punctuation' && t.rawText === '(') {
					depth++;
					if (depth === 1) {
						continue;
					}
				} else if (t.kind === 'punctuation' && t.rawText === ')') {
					depth--;
					if (depth === 0) {
						closed = true;
						break;
					}
				}
				if (depth >= 1) {
					inner.push(t);
				}
			}
			if (!closed) {
				return undefined;
			}
			argToks = inner;
		} else {
			argToks = [];
		}
	} else {
		if (next?.rawText === '(') {
			return undefined; // expressionCalls handles parenthesized forms.
		}
		if (next) {
			const gap = source.slice(span.start + callee.end, span.start + next.start);
			if (!/\s/.test(gap)) {
				return undefined;
			}
		}
		argToks = toks.slice(qualifierIdx + 3);
	}

	let depth = 0;
	for (let k = qualifierIdx + 3; k < toks.length; k++) {
		const raw = toks[k].rawText;
		if (raw === '(' || raw === '[') {
			depth++;
		} else if (raw === ')' || raw === ']') {
			depth--;
		} else if (depth === 0 && raw === '=') {
			return undefined;
		}
	}

	const split = argToks.length === 0 ? emptyArgSplit() : splitArgSlots(argToks, sliceStart);
	return {
		name,
		qualifier,
		lookupKey,
		nameSpan: { start: span.start + callee.start, end: span.start + callee.end },
		explicitCall,
		slots: split.slots,
		slotSpans: split.spans,
		sliceStart,
	};
}

export interface ArgSplit {
	slots: VbaToken[][];
	spans: Span[];
}

/** Splits an argument token run into top-level (depth-0) comma-separated slots. */
export function splitArgSlots(toks: VbaToken[], sliceStart: number): ArgSplit {
	const slots: VbaToken[][] = [[]];
	const spans: Span[] = [];
	let depth = 0;
	let emptyMarker: VbaToken | undefined;
	const finishSlot = (nextSeparator?: VbaToken): void => {
		const slot = slots[slots.length - 1];
		spans.push(argumentSlotSpan(slot, emptyMarker, nextSeparator, sliceStart));
	};
	for (const t of toks) {
		if (t.kind === 'punctuation' && t.rawText === '(') {
			depth++;
		} else if (t.kind === 'punctuation' && t.rawText === ')') {
			depth--;
		}
		if (t.kind === 'punctuation' && t.rawText === ',' && depth === 0) {
			finishSlot(t);
			slots.push([]);
			emptyMarker = t;
		} else {
			slots[slots.length - 1].push(t);
			emptyMarker = undefined;
		}
	}
	finishSlot();
	return { slots, spans };
}

export function emptyArgSplit(): ArgSplit {
	return { slots: [], spans: [] };
}

function argumentSlotSpan(
	slot: VbaToken[],
	emptyMarker: VbaToken | undefined,
	nextSeparator: VbaToken | undefined,
	sliceStart: number,
): Span {
	if (slot.length > 0) {
		return {
			start: sliceStart + slot[0].start,
			end: sliceStart + slot[slot.length - 1].end,
		};
	}
	if (emptyMarker) {
		return { start: sliceStart + emptyMarker.start, end: sliceStart + emptyMarker.end };
	}
	if (nextSeparator) {
		return { start: sliceStart + nextSeparator.start, end: sliceStart + nextSeparator.end };
	}
	return { start: sliceStart, end: sliceStart };
}

/** True if a slot is a named argument (`name := value`). */
export function isNamedSlot(slot: VbaToken[]): boolean {
	return (
		slot.length >= 2 &&
		(slot[0].kind === 'identifier' || slot[0].kind === 'bracketedIdentifier') &&
		slot[1].kind === 'operator' &&
		slot[1].rawText === ':='
	);
}

export function callableAcceptsZeroArguments(sig: CallableTypeSignature): boolean {
	return sig.params.every((param) => param.optional || param.paramArray);
}

/** Describes a procedure's acceptable argument-count range for a message. */
export function describeArity(required: number, max: number): string {
	if (max === Infinity) {
		return `at least ${required} argument${required === 1 ? '' : 's'}`;
	}
	if (required === max) {
		return `${required} argument${required === 1 ? '' : 's'}`;
	}
	return `between ${required} and ${max} arguments`;
}

/**
 * Validates one call's argument list against a procedure's parameters. When the
 * call uses named arguments, each name is checked against the parameter names
 * and the positional count check is skipped (positional/named mixing is too
 * subtle to count safely); otherwise the supplied slot count is checked against
 * the required minimum and the maximum implied by `Optional`/`ParamArray`.
 */
export function validateArity(
	source: string,
	sig: CallableTypeSignature,
	call: CallArguments,
	push: PushFn,
): void {
	const displayName = callDisplayName(sig, call);
	const params = sig.params;
	let required = params.length;
	for (let k = 0; k < params.length; k++) {
		if (params[k].optional || params[k].paramArray) {
			required = k;
			break;
		}
	}
	const hasParamArray = params.some((p) => p.paramArray);
	const max = hasParamArray ? Infinity : params.length;

	const named = call.slots.filter(isNamedSlot);
	if (named.length > 0) {
		// PCEC_008 + omitted-after-named: nothing positional may follow a named
		// argument. VBE rejects at compile (regardless of the signature) both a
		// positional VALUE after a named arg (oracle positional_after_named_argument_compile
		// — `f(a:=1, 2)` fires on the `2`) and an OMITTED/empty slot after a named arg
		// (oracle omit_trailing_after_named_compile `f(a:=1, )` and
		// omit_middle_between_named_compile `f(a:=1, , c:=3)`). Pure slot-order syntax:
		// once a named slot is seen, the first subsequent NON-named slot is the
		// violation. Omissions BEFORE the first named arg stay legal (oracle
		// omit_leading_before_named_compile `f(, b:=2)`, omit_positional_then_omit_then_named_compile
		// `f(1, , c:=3)`), as does the `f(1, b:=2)` positional-then-named ordering.
		let sawNamed = false;
		for (let i = 0; i < call.slots.length; i++) {
			const slot = call.slots[i];
			if (isNamedSlot(slot)) {
				sawNamed = true;
				continue;
			}
			if (!sawNamed) {
				continue;
			}
			if (slot.length > 0) {
				push(
					'argumentCount',
					`A positional argument may not follow a named argument in the call to '${displayName}'.`,
					{
						start: call.sliceStart + slot[0].start,
						end: call.sliceStart + slot[slot.length - 1].end,
					},
				);
			} else {
				push(
					'argumentCount',
					`An omitted argument may not follow a named argument in the call to '${displayName}'.`,
					call.slotSpans?.[i] ?? call.nameSpan,
				);
			}
			break; // one syntax error per call, matching VBE
		}
		const paramNames = new Set(
			params.map((p) => stripHeaderBrackets(p.name).toLowerCase()),
		);
		const seen = new Set<string>();
		for (const slot of named) {
			const raw = stripHeaderBrackets(slot[0].rawText);
			const lower = raw.toLowerCase();
			if (!paramNames.has(lower)) {
				push(
					'argumentCount',
					`Named argument not found: '${raw}' is not a parameter of '${displayName}'.`,
					{
						start: call.sliceStart + slot[0].start,
						end: call.sliceStart + slot[0].end,
					},
				);
				continue;
			}
			if (seen.has(lower)) {
				push(
					'argumentCount',
					`Named argument already specified: '${raw}' is supplied more than once to '${displayName}'.`,
					{
						start: call.sliceStart + slot[0].start,
						end: call.sliceStart + slot[0].end,
					},
				);
				continue;
			}
			seen.add(lower);
		}
		return; // positional count is not validated alongside named arguments
	}

	for (let i = 0; i < Math.min(call.slots.length, params.length); i++) {
		const param = params[i];
		if (call.slots[i].length === 0 && !param.optional && !param.paramArray) {
			const name = stripHeaderBrackets(param.name);
			const placeholder = omittedArgumentPlaceholderData(source, call, param, i);
			push(
				'argumentCount',
				`Argument not optional: '${name}' is required by '${displayName}'.`,
				call.slotSpans?.[i] ?? call.nameSpan,
				placeholder,
			);
		}
	}

	const n = call.slots.length;
	if (n < required || n > max) {
		const missingParam = n < required ? params[n] : undefined;
		const placeholder = missingParam
			? trailingMissingArgumentPlaceholderData(source, call, missingParam)
			: undefined;
		push(
			'argumentCount',
			`Wrong number of arguments to '${displayName}': expected ${describeArity(required, max)}, but got ${n}.`,
			call.nameSpan,
			placeholder,
		);
	}
}

function callDisplayName(sig: CallableTypeSignature, call: CallArguments): string {
	return call.qualifier ? `${call.qualifier}.${sig.name}` : sig.name;
}

function omittedArgumentPlaceholderData(
	source: string,
	call: CallArguments,
	param: CallableParamType,
	slotIndex: number,
): VbaDiagnosticData | undefined {
	const separator = call.slotSpans?.[slotIndex];
	if (!separator || source.slice(separator.start, separator.end) !== ',') {
		return undefined;
	}
	const parameterName = stripHeaderBrackets(param.name);
	const placeholder = placeholderNameForParameter(parameterName);
	return {
		missingRequiredArgumentPlaceholder: {
			parameterName,
			edit: {
				span: separatorWithFollowingHorizontalSpace(source, separator),
				newText: slotIndex === 0 ? `${placeholder}, ` : `, ${placeholder}`,
			},
		},
	};
}

function trailingMissingArgumentPlaceholderData(
	source: string,
	call: CallArguments,
	param: CallableParamType,
): VbaDiagnosticData | undefined {
	const parameterName = stripHeaderBrackets(param.name);
	const placeholder = placeholderNameForParameter(parameterName);
	if (call.slots.length === 0) {
		const innerSpan = emptyParenthesizedArgumentSpan(source, call);
		if (innerSpan) {
			return missingArgumentPlaceholderData(parameterName, innerSpan, placeholder);
		}
		const newText = call.explicitCall ? `(${placeholder})` : ` ${placeholder}`;
		return missingArgumentPlaceholderData(
			parameterName,
			{ start: call.nameSpan.end, end: call.nameSpan.end },
			newText,
		);
	}
	const lastSpan = call.slotSpans?.[call.slotSpans.length - 1];
	if (!lastSpan) {
		return undefined;
	}
	return missingArgumentPlaceholderData(
		parameterName,
		{ start: lastSpan.end, end: lastSpan.end },
		`, ${placeholder}`,
	);
}

function missingArgumentPlaceholderData(
	parameterName: string,
	span: Span,
	newText: string,
): VbaDiagnosticData {
	return {
		missingRequiredArgumentPlaceholder: {
			parameterName,
			edit: { span, newText },
		},
	};
}

function placeholderNameForParameter(name: string): string {
	const safe = name.replace(/[^A-Za-z0-9_]/g, '_').replace(/^[^A-Za-z_]+/, '');
	return `TODO_${safe || 'Argument'}`;
}

function separatorWithFollowingHorizontalSpace(source: string, separator: Span): Span {
	let end = separator.end;
	while (end < source.length && (source[end] === ' ' || source[end] === '\t')) {
		end++;
	}
	return { start: separator.start, end };
}

function emptyParenthesizedArgumentSpan(source: string, call: CallArguments): Span | undefined {
	let open = call.nameSpan.end;
	while (open < source.length && (source[open] === ' ' || source[open] === '\t')) {
		open++;
	}
	if (source[open] !== '(') {
		return undefined;
	}
	let close = open + 1;
	while (close < source.length && (source[close] === ' ' || source[close] === '\t')) {
		close++;
	}
	if (source[close] !== ')') {
		return undefined;
	}
	return { start: open + 1, end: close };
}
