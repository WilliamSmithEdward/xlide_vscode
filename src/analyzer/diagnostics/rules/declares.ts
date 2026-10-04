// Rule family: Declare statements (issue #254). Every case was measured in
// Excel 16.0 (build 20326, 2026-10-01).
//
//  - invalid-proc-header: what the VBE refuses in a Declare's own line. A
//    Declare Sub given an As clause, a Declare Function given both a type
//    character and an As clause, or one returning `String * n`, is
//    "Expected: end of statement"; one returning `As Any` is "Expected: type
//    name"; a Lib or Alias that is not a string literal, a Const included, is
//    "Expected: string constant". A parameter `As String * n`, of a Declare
//    or of a procedure, is "Expected array".
//  - unusable-declare: a Declare that compiles and fails on every call. CDecl
//    raises 49, Bad DLL calling convention; `Lib ""` or a Lib of spaces 48,
//    File not found; an Alias of nothing or spaces 453, and `Alias "#0"` 452,
//    Can't find DLL entry point. A Declare never called runs, so each call is
//    reported, not the Declare. A real library or entry point the machine
//    lacks (53, 453) depends on the machine and is left alone.

import type { ConditionalActivityTracker } from '../../conditional/conditionalCompilation';
import type { VbaToken } from '../../lexer/tokenKinds';
import type { DeclareNode, ModuleNode, ParameterNode, Span } from '../../parser/nodes';
import type { buildModuleSymbols } from '../../symbols/buildModuleSymbols';
import { procedureSymbolFor, type PushFn } from '../analysisContext';
import { stringLiteralValue } from '../typeInference';
import { activeModuleMembers, forEachStatement, matchParenFrom, statementTokens, tokenName, tokenText } from '../walker';

export function checkDeclareStatements(
	source: string,
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind === 'Declare') {
			checkDeclareLine(source, member, push);
		}
		if (member.kind === 'Declare' || member.kind === 'Procedure') {
			for (const param of member.params) {
				checkFixedLengthParameter(source, param, push);
			}
		}
	}
}

function checkDeclareLine(source: string, declare: DeclareNode, push: PushFn): void {
	const toks = statementTokens(source, declare.span);
	const at = (tok: VbaToken): Span => ({ start: declare.span.start + tok.start, end: declare.span.start + tok.end });
	const refuse = (tok: VbaToken, what: string, error: string): void => {
		push('invalidProcedureHeader', `Declare '${declare.name}': ${what}. This is a VBE compile error: ${error}.`, at(tok));
	};
	let open = -1;
	for (let i = 0; i < toks.length; i++) {
		const word = tokenText(toks[i]);
		if ((word === 'lib' || word === 'alias') && toks[i + 1] && toks[i + 1].kind !== 'stringLiteral') {
			refuse(toks[i + 1], `${toks[i].rawText} takes a string literal, and '${toks[i + 1].rawText}' is none`, 'Expected: string constant');
			return;
		}
		if (toks[i].rawText === '(' && open < 0) {
			open = i;
		}
	}
	const close = open < 0 ? -1 : matchParenFrom(toks, open);
	const as = close < 0 ? undefined : toks[close + 1];
	if (!as || tokenText(as) !== 'as') {
		return;
	}
	if (!declare.isFunction) {
		refuse(as, 'a Sub returns nothing, so it takes no As clause', 'Expected: end of statement');
	} else if (declare.typeSuffix) {
		refuse(as, `the type character '${declare.typeSuffix}' already gives the return type`, 'Expected: end of statement');
	} else if (tokenText(toks[close + 2]) === 'any') {
		refuse(toks[close + 2], 'As Any is for a parameter, not a return type', 'Expected: type name');
	} else if (toks[close + 3]?.rawText === '*') {
		refuse(toks[close + 3], 'a Declare Function cannot return a fixed-length String', 'Expected: end of statement');
	}
}

/** `ByVal s As String * 4`: no parameter takes a fixed-length String. */
function checkFixedLengthParameter(source: string, param: ParameterNode, push: PushFn): void {
	const toks = statementTokens(source, param.span);
	const star = toks.findIndex((tok, i) => tok.rawText === '*' && tokenText(toks[i - 1]) === 'string' && tokenText(toks[i - 2]) === 'as');
	if (star < 0) {
		return;
	}
	push(
		'invalidProcedureHeader',
		`Parameter '${param.name}' is declared As String * ${toks[star + 1]?.rawText ?? 'n'}, and no parameter takes a fixed-length String. This is a VBE compile error: Expected array.`,
		{ start: param.span.start + toks[star - 1].start, end: param.span.start + (toks[star + 1] ?? toks[star]).end },
	);
}

/** Why every call of a Declare fails, or undefined. */
function failure(source: string, declare: DeclareNode): string | undefined {
	const toks = statementTokens(source, declare.span);
	const valueAfter = (word: string): string | undefined => {
		const i = toks.findIndex((tok) => tokenText(tok) === word);
		return i >= 0 && toks[i + 1]?.kind === 'stringLiteral' ? stringLiteralValue(toks[i + 1].rawText) : undefined;
	};
	if (toks.some((tok) => tokenText(tok) === 'cdecl')) {
		return `is declared CDecl, which Office on Windows refuses: every call to it raises Run-time error '49': Bad DLL calling convention`;
	}
	const lib = valueAfter('lib');
	if (lib !== undefined && lib.trim() === '') {
		return `names no library, Lib "${lib}": every call to it raises Run-time error '48': File not found`;
	}
	const alias = valueAfter('alias');
	if (alias !== undefined && alias.trim() === '') {
		return `names no entry point, Alias "${alias}": every call to it raises Run-time error '453': Can't find DLL entry point`;
	}
	if (alias !== undefined && /^#0+$/.test(alias)) {
		return `names entry point ${alias}, an ordinal no library has: every call to it raises Run-time error '452': Can't find DLL entry point 0`;
	}
	return undefined;
}

export function checkUnusableDeclareCalls(
	source: string,
	mod: ModuleNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	const failing = new Map<string, string>();
	for (const member of activeModuleMembers(mod, activity)) {
		const why = member.kind === 'Declare' ? failure(source, member) : undefined;
		if (why && member.kind === 'Declare') {
			failing.set(member.name.toLowerCase(), why);
		}
	}
	if (failing.size === 0) {
		return;
	}
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		// A parameter or local of the same name hides the Declare.
		const hidden = new Set((procedureSymbolFor(symbols, member)?.children ?? []).map((child) => child.name.toLowerCase()));
		// A single-line If's statement holds its branches.
		forEachStatement(member.body, (stmt) => {
			const toks = statementTokens(source, stmt.span);
			for (let i = 0; i < toks.length; i++) {
				const lower = tokenName(toks[i])?.toLowerCase();
				const why = lower && !hidden.has(lower) ? failing.get(lower) : undefined;
				if (!why || toks[i - 1]?.rawText === '.' || toks[i - 1]?.rawText === '!') {
					continue;
				}
				push('unusableDeclare', `'${toks[i].rawText}' ${why}.`, { start: stmt.span.start + toks[i].start, end: stmt.span.start + toks[i].end });
			}
		}, activity);
	}
}
