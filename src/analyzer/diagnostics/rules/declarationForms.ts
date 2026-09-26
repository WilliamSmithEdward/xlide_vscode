// Rule family: declaration forms the VBE refuses while compiling (issue
// #124). Each was measured in Excel 16.0 (build 20326, 2026-09-25/26) with
// the message quoted.
//
//  - array-parameter-form: `ByVal a() As Long` -> "Array argument must be
//    ByRef"; `Optional a() As Long` -> "Optional argument must be Variant or
//    intrinsic type with a default value".
//  - parameter-default-type-mismatch: `Optional ByVal i As Integer = 40000`,
//    `Optional ByVal b As Byte = 256` -> "Overflow".
//  - const-overflow: an Enum member `eA = 3000000000#` -> "Overflow" (Enum
//    members are Long).
//  - duplicate-deftype: `DefLng A-Z` followed by `DefStr S` -> "Duplicate
//    Deftype statement": a letter may be given a default type once.
//  - bracketed-variable-name: `Dim [my var] As Long` -> "Syntax error".
//    Brackets make a foreign name for an Enum member (`[Two Words] = 2`
//    compiles) but not for a variable.

import type { ConditionalActivityTracker } from '../../conditional/conditionalCompilation';
import { parseVbaIntegerLiteral } from '../../constants/integerConstantExpression';
import type { VbaToken } from '../../lexer/tokenKinds';
import type { ModuleNode, ParameterNode, Span } from '../../parser/nodes';
import type { PushFn } from '../analysisContext';
import { normalizeType, numericLiteralBounds } from '../typeInference';
import {
	absoluteSpan,
	activeModuleMembers,
	declaredNameSpan,
	forEachVariableGroup,
	isInactiveNode,
	statementTokensAfterLeadingLabel,
	tokenName,
	tokenText,
} from '../walker';
import { DEFTYPE_KEYWORDS } from './shared';

const LONG_RANGE = { min: -2147483648, max: 2147483647 };

export function checkDeclarationForms(
	source: string,
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	const claimedLetters = new Map<string, string>();
	for (const member of activeModuleMembers(mod, activity)) {
		switch (member.kind) {
			case 'Procedure':
				for (const param of member.params) {
					checkParameter(source, param, push);
				}
				forEachVariableGroup(member.body, (group) => {
					for (const decl of group.declarations) {
						checkBracketedName(source, decl.name, decl.nameSpan, push);
					}
				}, activity);
				break;
			case 'VariableGroup':
				for (const decl of member.declarations) {
					checkBracketedName(source, decl.name, decl.nameSpan, push);
				}
				break;
			case 'Enum':
				for (const item of member.members) {
					if (item.valueRaw === undefined || isInactiveNode(activity, item)) {
						continue;
					}
					const value = literalNumber(statementTokensAfterLeadingLabel(item.valueRaw, { start: 0, end: item.valueRaw.length }));
					if (value !== undefined && (value < LONG_RANGE.min || value > LONG_RANGE.max)) {
						push(
							'constOverflow',
							`Enum member '${item.name}' is ${item.valueRaw.trim()}, outside the Long range ${LONG_RANGE.min} to ${LONG_RANGE.max} an Enum member holds. This is a VBE compile error: Overflow.`,
							declaredNameSpan(source, item.span, item.name),
						);
					}
				}
				break;
			case 'Statement':
				checkDeftype(source, member.span, claimedLetters, push);
				break;
			default:
				break;
		}
	}
}

function checkParameter(source: string, param: ParameterNode, push: PushFn): void {
	const nameSpan = param.nameSpan ?? param.span;
	if (param.isArray && param.byVal && !param.paramArray) {
		push('arrayParameterForm', `Array parameter '${param.name}' cannot be ByVal: an array argument must be ByRef.`, nameSpan);
	}
	if (param.isArray && param.optional) {
		push('arrayParameterForm', `Optional parameter '${param.name}' cannot be an array: an Optional argument must be Variant or an intrinsic type with a default value.`, nameSpan);
	}
	if (!param.optional || param.defaultRaw === undefined || param.isArray) {
		return;
	}
	const type = normalizeType(param.asType);
	const bounds = type ? numericLiteralBounds(type) : undefined;
	if (!bounds) {
		return;
	}
	const value = literalNumber(statementTokensAfterLeadingLabel(param.defaultRaw, { start: 0, end: param.defaultRaw.length }));
	if (value === undefined) {
		return;
	}
	const stored = type === 'single' || type === 'double' || type === 'currency' ? value : Math.round(value);
	if (stored < bounds.min || stored > bounds.max) {
		const at = source.indexOf(param.defaultRaw.trim(), param.span.start);
		const span: Span = at >= 0 && at < param.span.end ? { start: at, end: at + param.defaultRaw.trim().length } : param.span;
		push(
			'parameterDefaultTypeMismatch',
			`Optional parameter '${param.name}' is declared As ${param.asType}, whose range is ${bounds.min} to ${bounds.max}; its default ${param.defaultRaw.trim()} does not fit. This is a VBE compile error: Overflow.`,
			span,
		);
	}
}

/** The number a plain signed literal denotes, or undefined for anything else. */
function literalNumber(toks: readonly VbaToken[]): number | undefined {
	let sign = 1;
	let rest = toks.filter((tok) => tok.kind !== 'comment');
	if (rest[0]?.rawText === '-' || rest[0]?.rawText === '+') {
		sign = rest[0].rawText === '-' ? -1 : 1;
		rest = rest.slice(1);
	}
	if (rest.length !== 1) {
		return undefined;
	}
	if (rest[0].kind === 'integerLiteral') {
		const parsed = parseVbaIntegerLiteral(rest[0].rawText);
		return parsed === undefined ? undefined : sign * parsed;
	}
	if (rest[0].kind === 'floatLiteral') {
		const parsed = Number(rest[0].rawText.replace(/[dD]/g, 'E').replace(/[!#@]$/, ''));
		return Number.isFinite(parsed) ? sign * parsed : undefined;
	}
	return undefined;
}

function checkBracketedName(source: string, name: string, nameSpan: Span | undefined, push: PushFn): void {
	if (nameSpan && source[nameSpan.start] === '[') {
		push(
			'bracketedVariableName',
			`'[${name}]' is not a variable name: brackets make a foreign name only for an Enum member or a member of another object. This is a VBE compile error: Syntax error.`,
			nameSpan,
		);
	}
}

/**
 * `DefLng A-Z` then `DefStr S`: every letter a Deftype names must be new.
 * Letters are tracked across the module's Deftype statements in order.
 */
function checkDeftype(source: string, span: Span, claimed: Map<string, string>, push: PushFn): void {
	const toks = statementTokensAfterLeadingLabel(source, span);
	const head = tokenText(toks[0]);
	if (!DEFTYPE_KEYWORDS.has(head)) {
		return;
	}
	const statement = toks[0].canonicalText ?? toks[0].rawText;
	const letters: string[] = [];
	for (let i = 1; i < toks.length; i++) {
		const name = tokenName(toks[i]);
		if (!name || name.length !== 1) {
			continue;
		}
		const upper = name.toUpperCase();
		if (toks[i - 1]?.rawText === '-' && letters.length > 0) {
			const from = letters[letters.length - 1].charCodeAt(0);
			for (let code = from + 1; code <= upper.charCodeAt(0); code++) {
				letters.push(String.fromCharCode(code));
			}
			continue;
		}
		letters.push(upper);
	}
	for (const letter of letters) {
		const earlier = claimed.get(letter);
		if (earlier) {
			push(
				'duplicateDeftype',
				`Letter '${letter}' already has a default type from '${earlier}' above; a letter takes one Deftype statement. This is a VBE compile error: Duplicate Deftype statement.`,
				absoluteSpan(span, toks[0]),
			);
			return;
		}
	}
	for (const letter of letters) {
		claimed.set(letter, statement);
	}
}
