// Rule: a Variant local whose value the code makes plain, used as something
// that value is not (issue #121). Measured in Excel 16.0 (build 20326,
// 2026-09-26); each compiles and raises every time it runs.
//
//  - A scalar used as an object: `v = 5` or `v = "abc"` then `v.Foo` -> 424,
//    Object required. An array used as one: `v = Array(1, 2)` then `v.Foo`
//    -> 424.
//  - A scalar used as an array: `v = 5` or `v = "abc"` then `UBound(v)` -> 13.
//  - An array used as a scalar: `v = Array(1, 2)` then `v + 1`, `v - 1`,
//    `v & "x"`, `If v = 1 Then` -> 13, Type mismatch.
//  - An array a call returns, used the same way (issue #239):
//    `Array(1) + 1`, `Split("a") + 1`, `-Array(1)`, `Not Array(1)`.
//  - A scalar where only an array will do (issue #219): `v = 5` then
//    `Erase v`, `ReDim Preserve v(2)` or `For Each x In v` -> 13. For Each
//    over a Variant nothing assigns, which is Empty, raises 13 too.
//
// The values come from the same analysis the division and subscript rules use:
// a literal, or an array from Array(), Split() on literals or a Range
// literal's Value, that either every assignment in the procedure gives, or
// the last assignment before the statement gives with nothing between able to
// change it (issue #180).

import type { ConditionalActivityTracker } from '../../conditional/conditionalCompilation';
import type { VbaToken } from '../../lexer/tokenKinds';
import type { BodyNode, ForBlockNode, ModuleNode } from '../../parser/nodes';
import { buildModuleSymbols } from '../../symbols/buildModuleSymbols';
import { procedureSymbolFor, type PushFn } from '../analysisContext';
import type { VbaSymbol } from '../../symbols/symbolModel';
import {
	knownLocalLiteralValuesAt,
	normalizeType,
	runtimeCallableSourceShadowed,
	sourceNameScopeFor,
	typeEnvironmentFor,
	type KnownLocalValue,
	type SourceNameScope,
} from '../typeInference';
import {
	activeModuleMembers,
	bareAssignmentTarget,
	forEachStatement,
	isInactiveNode,
	matchParenFrom,
	statementAndBranchSpans,
	statementTokens,
	tokenName,
	tokenText,
} from '../walker';
import { isBareOrVbaQualifiedIntrinsicCall, nameMentions } from './shared';
import { knownArrayShapesAt, moduleOptionBase, type FixedArrayBound } from './arrays';

const SCALAR_OPERATORS: ReadonlySet<string> = new Set(['=', '<', '>', '<=', '>=', '<>', '+', '-', '*', '/', '\\', '&', '^']);

export function checkVariantValueMisuse(
	source: string,
	mod: ModuleNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
	projectVisibleSymbols?: readonly VbaSymbol[],
): void {
	const optionBase = moduleOptionBase(mod, activity);
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		const env = typeEnvironmentFor(symbols, member);
		const isVariant = (lower: string): boolean => {
			const type = normalizeType(env.get(lower));
			return type === undefined || type === 'variant';
		};
		// What each Variant holds at a statement (issue #180): the last
		// assignment to reach it, or the one value the procedure agrees on.
		const valuesAt = knownLocalLiteralValuesAt(source, member, symbols, activity);
		const shapesAt = knownArrayShapesAt(source, symbols, member, activity, optionBase);
		const scalarsFor = memoByIdentity((values: ReadonlyMap<string, KnownLocalValue>) => {
			const scalars = new Map<string, string>();
			for (const [lower, value] of values) {
				if (value.origin === 'literal' && isVariant(lower)) {
					scalars.set(lower, value.kind === 'string' ? `the string "${value.value}"` : `the number ${value.value}`);
				}
			}
			return scalars;
		});
		const arraysFor = memoByIdentity((shapes: ReadonlyMap<string, FixedArrayBound>) => {
			const arrays = new Map<string, string>();
			for (const [lower, shape] of shapes) {
				if (isVariant(lower)) {
					arrays.set(lower, shape.origin);
				}
			}
			return arrays;
		});
		// A Variant local is still Empty at a statement that names it first,
		// though Erase and ReDim end what the value analysis knows of it: the
		// only statement to name it, or the first in a procedure with no GoTo,
		// GoSub or Resume to come back to an earlier line. Inside a block too:
		// the first pass through a loop, or the arm that runs, reaches it
		// before anything else names it (issue #237).
		let mentions: Map<string, number> | undefined;
		let procedureTokens: readonly VbaToken[] | undefined;
		const emptyHere = (lower: string, offset: number): boolean => {
			const local = procedureSymbolFor(symbols, member)?.children?.find((child) => child.name.toLowerCase() === lower);
			if (local?.kind !== 'localVariable' || local.isArray || local.visibility === 'Static' || !isVariant(lower)) {
				return false;
			}
			if ((mentions ??= nameMentions(source, member, activity)).get(lower) === 1) {
				return true;
			}
			procedureTokens ??= statementTokens(source, member.span).map((tok) => ({ ...tok, start: tok.start + member.span.start, end: tok.end + member.span.start }));
			if (procedureTokens.some((tok) => ['goto', 'gosub', 'resume'].includes(tokenText(tok)))) {
				return false;
			}
			// The first use after the declaration (a Dim names it too).
			const uses = procedureTokens.filter((tok) => tokenName(tok)?.toLowerCase() === lower && !isInDeclaration(source, tok.start));
			return uses[0]?.start === offset;
		};
		const sourceNames = sourceNameScopeFor(symbols, member, projectVisibleSymbols);
		forEachStatement(member.body, (stmt) => {
			for (const span of statementAndBranchSpans(stmt)) {
				for (const hit of arrayCallOperands(statementTokens(source, span), sourceNames)) {
					push('variantValueMisuse', hit.message, { start: span.start + hit.start, end: span.start + hit.end });
				}
			}
			for (const span of statementAndBranchSpans(stmt)) {
				const toks = statementTokens(source, span);
				const head = tokenText(toks[0]);
				if (head !== 'erase' && head !== 'redim') {
					continue;
				}
				for (let i = 1; i < toks.length; i++) {
					const lower = tokenName(toks[i])?.toLowerCase();
					const statement = lower ? arrayStatementTarget(toks, i) : undefined;
					if (statement && emptyHere(lower!, span.start + toks[i].start)) {
						push('variantValueMisuse', `'${toks[i].rawText}' is never assigned, so it is Empty here, which is not an array for ${statement} to act on. This will raise Run-time error '13': Type mismatch.`, { start: span.start + toks[i].start, end: span.start + toks[i].end });
					}
				}
			}
			const scalars = scalarsFor(valuesAt(stmt));
			const arrays = arraysFor(shapesAt(stmt));
			if (scalars.size === 0 && arrays.size === 0) {
				return;
			}
			for (const span of statementAndBranchSpans(stmt)) {
				const toks = statementTokens(source, span);
				const target = bareAssignmentTarget(source, span);
				const targetIndex = target ? toks.findIndex((tok) => tok.rawText === '=') - 1 : -1;
				for (let i = 0; i < toks.length; i++) {
					if (i === targetIndex || toks[i - 1]?.rawText === '.') {
						continue;
					}
					const lower = tokenName(toks[i])?.toLowerCase();
					if (!lower) {
						continue;
					}
					const at = { start: span.start + toks[i].start, end: span.start + toks[i].end };
					const scalar = scalars.get(lower);
					const array = arrays.get(lower);
					if (!scalar && !array) {
						continue;
					}
					const next = toks[i + 1];
					if (next?.rawText === '.' && tokenName(toks[i + 2])) {
						const holds = scalar ?? `an array from ${array}`;
						push('variantValueMisuse', `'${toks[i].rawText}' holds ${holds} here, which has no members. This will raise Run-time error '424': Object required.`, at);
						continue;
					}
					if (scalar && isBoundArgument(toks, i)) {
						push('variantValueMisuse', `'${toks[i].rawText}' holds ${scalar} here, which is not an array. This will raise Run-time error '13': Type mismatch.`, at);
						continue;
					}
					const arrayStatement = scalar ? arrayStatementTarget(toks, i) : undefined;
					if (scalar && arrayStatement) {
						push('variantValueMisuse', `'${toks[i].rawText}' holds ${scalar} here, which is not an array for ${arrayStatement} to act on. This will raise Run-time error '13': Type mismatch.`, at);
						continue;
					}
					const len = array ? lenCallAround(toks, i, i, sourceNames) : undefined;
					if (len) {
						push('variantValueMisuse', `'${toks[i].rawText}' holds an array from ${array} here, which ${len} cannot measure. This will raise Run-time error '13': Type mismatch.`, at);
						continue;
					}
					if (array && next?.rawText !== '(') {
						// The operator on either side, never the assignment's own `=`.
						const previous = i - 1 === targetIndex + 1 ? undefined : toks[i - 1];
						const operator = [next, previous].find((tok) => tok && ((tok.kind === 'operator' && SCALAR_OPERATORS.has(tok.rawText)) || tokenText(tok) === 'mod'));
						if (operator) {
							push('variantValueMisuse', `'${toks[i].rawText}' holds an array from ${array} here, which '${operator.rawText}' cannot combine with a scalar. This will raise Run-time error '13': Type mismatch.`, at);
						}
					}
				}
			}
		}, activity);
		// `For Each x In v` with v a scalar or Empty (issue #219).
		forEachLoopOver(member.body, activity, (loop) => {
			const lower = loop.sourceExpression?.trim().toLowerCase();
			if (!lower || !/^[a-z_][a-z0-9_]*$/.test(lower) || !isVariant(lower) || !loop.sourceExpressionSpan) {
				return;
			}
			const value = valuesAt(loop).get(lower);
			const holds = value?.kind === 'empty'
				? 'nothing (it is never assigned, so it is Empty)'
				: value?.origin === 'literal'
					? (value.kind === 'string' ? `the string "${value.value}"` : `the number ${value.value}`)
					: undefined;
			if (holds) {
				push('variantValueMisuse', `'${loop.sourceExpression!.trim()}' holds ${holds} here, which For Each cannot step through. This will raise Run-time error '13': Type mismatch.`, loop.sourceExpressionSpan);
			}
		});
	}
}

/** The calls that return an array whatever their arguments. */
const ARRAY_FUNCTIONS: ReadonlySet<string> = new Set(['array', 'split']);

/**
 * `Array(1) + 1`, `Split("a") & "x"`, `-Array(1)` and `Not Array(1)`: an
 * array a call returns, as the operand of a scalar operator. Each raises
 * 13 (issue #239, measured in Excel 16.0). Offsets are the statement's.
 */
function arrayCallOperands(
	toks: readonly VbaToken[],
	sourceNames: SourceNameScope,
): Array<{ start: number; end: number; message: string }> {
	const out: Array<{ start: number; end: number; message: string }> = [];
	// A single-line If's own line is its condition; each branch comes as a span of its own.
	const condition = tokenText(toks[0]) === 'if';
	const end = condition ? toks.findIndex((tok) => tokenText(tok) === 'then') : toks.length;
	const assignment = condition ? -1 : topLevelEquals(toks);
	for (let i = 0; i < end - 1; i++) {
		const name = tokenText(toks[i]);
		if (!ARRAY_FUNCTIONS.has(name) || toks[i + 1].rawText !== '(' || !isBareOrVbaQualifiedIntrinsicCall(toks, i)) {
			continue;
		}
		const qualified = toks[i - 1]?.rawText === '.';
		if (!qualified && runtimeCallableSourceShadowed(toks[i].rawText, sourceNames)) {
			continue;
		}
		const close = matchParenFrom(toks, i + 1);
		// `Split(s)(0)` indexes the array, and its element is a scalar.
		if (close < 0 || toks[close + 1]?.rawText === '(') {
			continue;
		}
		const first = qualified ? i - 2 : i;
		// A statement's own `=` assigns; any other `=` compares.
		const before = first - 1 === assignment ? undefined : toks[first - 1];
		const after = toks[close + 1];
		const operator = [after, before].find((tok) => tok && ((tok.kind === 'operator' && SCALAR_OPERATORS.has(tok.rawText)) || ['mod', 'not', 'and', 'or', 'xor', 'like'].includes(tokenText(tok))));
		const call = toks.slice(first, close + 1).map((tok) => tok.rawText).join('');
		const len = operator ? undefined : lenCallAround(toks, first, close, sourceNames);
		if (len) {
			out.push({
				start: toks[first].start,
				end: toks[close].end,
				message: `${call} returns an array, which ${len} cannot measure. This will raise Run-time error '13': Type mismatch.`,
			});
			continue;
		}
		if (!operator) {
			continue;
		}
		out.push({
			start: toks[first].start,
			end: toks[close].end,
			message: `${call} returns an array, which '${operator.rawText}' cannot use as a scalar. This will raise Run-time error '13': Type mismatch.`,
		});
	}
	return out;
}

/**
 * The Len or LenB whose one argument runs from `first` to `last`: an array
 * value there raises 13 (issue #248, measured in Excel 16.0).
 */
function lenCallAround(toks: readonly VbaToken[], first: number, last: number, sourceNames: SourceNameScope): string | undefined {
	const callee = first - 2;
	const word = tokenText(toks[callee]);
	if ((word !== 'len' && word !== 'lenb') || toks[first - 1]?.rawText !== '(' || toks[last + 1]?.rawText !== ')' || !isBareOrVbaQualifiedIntrinsicCall(toks, callee)) {
		return undefined;
	}
	if (toks[callee - 1]?.rawText === '.') {
		return `${toks[callee - 2].rawText}.${toks[callee].rawText}`;
	}
	return runtimeCallableSourceShadowed(toks[callee].rawText, sourceNames) ? undefined : toks[callee].rawText;
}

/** The first `=` outside parentheses, which is the assignment's own; -1 when none. */
function topLevelEquals(toks: readonly VbaToken[]): number {
	let depth = 0;
	for (let i = 0; i < toks.length; i++) {
		const raw = toks[i].rawText;
		if (raw === '(') {
			depth++;
		} else if (raw === ')') {
			depth--;
		} else if (raw === '=' && depth === 0) {
			return i;
		}
	}
	return -1;
}

/** Whether the offset is on a Dim, Static or Const line, which declares rather than uses. */
function isInDeclaration(source: string, offset: number): boolean {
	const lineStart = source.lastIndexOf(String.fromCharCode(10), offset - 1) + 1;
	return /^\s*(?:dim|static|const)(?![a-z0-9_])/i.test(source.slice(lineStart, offset));
}

/** Every For Each loop in a body, nested ones included. */
function forEachLoopOver(body: readonly BodyNode[], activity: ConditionalActivityTracker | undefined, visit: (loop: ForBlockNode) => void): void {
	for (const node of body) {
		if (isInactiveNode(activity, node)) {
			continue;
		}
		if (node.kind === 'ForBlock' && node.each) {
			visit(node);
		}
		// An If block's body holds every arm's statements.
		if ('body' in node && Array.isArray(node.body)) {
			forEachLoopOver(node.body as BodyNode[], activity, visit);
		}
	}
}

/**
 * The statement a name is the target of, where only an array will do:
 * `Erase v` and `ReDim Preserve v(2)` raise 13 on a scalar (issue #219,
 * measured in Excel 16.0). A plain ReDim makes v an array and runs.
 */
function arrayStatementTarget(toks: readonly VbaToken[], i: number): string | undefined {
	const head = tokenText(toks[0]);
	const previous = toks[i - 1]?.rawText;
	if (head === 'erase' && (i === 1 || previous === ',')) {
		return 'Erase';
	}
	if (head === 'redim' && tokenText(toks[1]) === 'preserve' && (i === 2 || previous === ',') && toks[i + 1]?.rawText === '(') {
		return 'ReDim Preserve';
	}
	return undefined;
}

/** `derive` run once per distinct input object. */
function memoByIdentity<K extends object, V>(derive: (key: K) => V): (key: K) => V {
	const cache = new Map<K, V>();
	return (key) => {
		let value = cache.get(key);
		if (value === undefined) {
			value = derive(key);
			cache.set(key, value);
		}
		return value;
	};
}

/** True when `toks[i]` is the whole first argument of UBound, LBound or Join (issue #239). */
function isBoundArgument(toks: readonly VbaToken[], i: number): boolean {
	const name = tokenText(toks[i - 2]);
	return toks[i - 1]?.rawText === '(' && (name === 'ubound' || name === 'lbound' || name === 'join') && toks[i - 3]?.rawText !== '.'
		&& (toks[i + 1]?.rawText === ')' || toks[i + 1]?.rawText === ',');
}
