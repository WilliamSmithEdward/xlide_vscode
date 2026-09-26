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
//
// The values come from the same analysis the division and subscript rules use:
// a local every assignment gives the same literal, or an array from Array(),
// Split() on literals or a Range literal's Value, with nothing else able to
// change it.

import type { ConditionalActivityTracker } from '../../conditional/conditionalCompilation';
import type { VbaToken } from '../../lexer/tokenKinds';
import type { ModuleNode } from '../../parser/nodes';
import { buildModuleSymbols } from '../../symbols/buildModuleSymbols';
import type { PushFn } from '../analysisContext';
import { knownLocalLiteralValues, normalizeType, typeEnvironmentFor } from '../typeInference';
import {
	activeModuleMembers,
	bareAssignmentTarget,
	forEachStatement,
	statementAndBranchSpans,
	statementTokens,
	tokenName,
	tokenText,
} from '../walker';
import { knownArrayShapes, moduleOptionBase } from './arrays';

const SCALAR_OPERATORS: ReadonlySet<string> = new Set(['=', '<', '>', '<=', '>=', '<>', '+', '-', '*', '/', '\\', '&', '^']);

export function checkVariantValueMisuse(
	source: string,
	mod: ModuleNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
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
		const scalars = new Map<string, string>();
		for (const [lower, value] of knownLocalLiteralValues(source, member, symbols, activity)) {
			if (value.origin === 'literal' && isVariant(lower)) {
				scalars.set(lower, value.kind === 'string' ? `the string "${value.value}"` : `the number ${value.value}`);
			}
		}
		const arrays = new Map<string, string>();
		for (const [lower, shape] of knownArrayShapes(source, member.body, symbols, member, activity, optionBase)) {
			if (isVariant(lower)) {
				arrays.set(lower, shape.origin);
			}
		}
		if (scalars.size === 0 && arrays.size === 0) {
			continue;
		}
		forEachStatement(member.body, (stmt) => {
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
	}
}

/** True when `toks[i]` is the whole first argument of UBound or LBound. */
function isBoundArgument(toks: readonly VbaToken[], i: number): boolean {
	const name = tokenText(toks[i - 2]);
	return toks[i - 1]?.rawText === '(' && (name === 'ubound' || name === 'lbound') && toks[i - 3]?.rawText !== '.'
		&& (toks[i + 1]?.rawText === ')' || toks[i + 1]?.rawText === ',');
}
