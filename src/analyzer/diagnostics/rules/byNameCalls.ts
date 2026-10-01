// Rule family: a call made by a name in a string (issue #243). Measured in
// Excel 16.0 (build 20326, 2026-09-30); each compiles and raises every time.
//
//  - `CallByName(c, "NoSuch", VbMethod)` with c a project class that has no
//    Public member of that name, a Private one included -> 438. On a
//    Collection, any name but Add, Count, Item and Remove -> 438.
//  - `CallByName c, "Hello", VbLet, 1` with Hello a Function, or VbGet on a
//    Sub or Function -> 450: the call type does not fit the member.
//  - In Excel, `Application.Run "NoSuch"` with no Sub or Function of that name in a
//    standard or document module of the project -> 1004. Private ones count;
//    a class module's members do not.
//
// Names compare without case. A name built at run time is not judged.

import type { MemberCompletionContext } from '../../completion/memberAccess';
import type { ConditionalActivityTracker } from '../../conditional/conditionalCompilation';
import type { VbaToken } from '../../lexer/tokenKinds';
import type { ModuleNode, Span } from '../../parser/nodes';
import type { buildModuleSymbols } from '../../symbols/buildModuleSymbols';
import type { PushFn } from '../analysisContext';
import { normalizeType, stringLiteralValue, typeEnvironmentFor } from '../typeInference';
import {
	activeModuleMembers,
	forEachStatement,
	matchParenFrom,
	statementAndBranchSpans,
	statementTokens,
	tokenName,
	tokenText,
} from '../walker';

const COLLECTION_MEMBERS: ReadonlySet<string> = new Set(['add', 'count', 'item', 'remove']);

/** VbCallType, by its name and its value. */
const CALL_TYPES: Readonly<Record<string, string>> = {
	vbmethod: 'method', vbget: 'get', vblet: 'let', vbset: 'set', '1': 'method', '2': 'get', '4': 'let', '8': 'set',
};

export function checkByNameCalls(
	source: string,
	mod: ModuleNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
	memberCtx: MemberCompletionContext,
	runnable: ReadonlySet<string> | undefined,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	const moduleNames = new Set((symbols.root.children ?? []).map((child) => child.name.toLowerCase()));
	// Word's Run also reaches the global templates, Normal.dotm among them,
	// which the project cannot see; only Excel's is judged.
	const excel = (memberCtx.model?.hostName ?? 'Excel') === 'Excel';
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		const env = typeEnvironmentFor(symbols, member);
		forEachStatement(member.body, (stmt) => {
			for (const span of statementAndBranchSpans(stmt)) {
				const toks = statementTokens(source, span).filter((tok) => tok.kind !== 'comment');
				for (let i = 0; i < toks.length; i++) {
					const word = tokenText(toks[i]);
					if (word === 'callbyname' && toks[i - 1]?.rawText !== '.' && !moduleNames.has('callbyname')) {
						checkCallByName(span, toks, i, env, memberCtx, push);
					} else if (word === 'run' && runnable && excel && toks[i - 1]?.rawText === '.' && tokenText(toks[i - 2]) === 'application' && toks[i - 3]?.rawText !== '.') {
						checkApplicationRun(span, toks, i, runnable, push);
					}
				}
			}
		}, activity);
	}
}

/** The arguments after `toks[name]`: in parentheses, or to the end of a call statement. */
function argumentsAfter(toks: readonly VbaToken[], name: number): VbaToken[][] {
	const paren = toks[name + 1]?.rawText === '(';
	const close = paren ? matchParenFrom(toks, name + 1) : toks.length;
	const out: VbaToken[][] = [];
	let current: VbaToken[] = [];
	let depth = 0;
	for (let i = name + (paren ? 2 : 1); i < close; i++) {
		const raw = toks[i].rawText;
		depth += raw === '(' ? 1 : raw === ')' ? -1 : 0;
		if (raw === ',' && depth === 0) {
			out.push(current);
			current = [];
			continue;
		}
		current.push(toks[i]);
	}
	out.push(current);
	return out;
}

function spanOf(base: Span, arg: readonly VbaToken[]): Span {
	return { start: base.start + arg[0].start, end: base.start + arg[arg.length - 1].end };
}

function checkCallByName(
	span: Span,
	toks: readonly VbaToken[],
	at: number,
	env: ReadonlyMap<string, string>,
	memberCtx: MemberCompletionContext,
	push: PushFn,
): void {
	const args = argumentsAfter(toks, at);
	const [object, procName, callType] = args;
	if (!object || object.length !== 1 || procName?.length !== 1 || procName[0].kind !== 'stringLiteral' || callType?.length !== 1) {
		return;
	}
	const type = env.get(tokenName(object[0])?.toLowerCase() ?? '');
	const name = stringLiteralValue(procName[0].rawText);
	const lower = name.toLowerCase();
	const kind = CALL_TYPES[tokenText(callType[0])];
	if (!type || !kind) {
		return;
	}
	if (normalizeType(type) === 'collection') {
		if (!COLLECTION_MEMBERS.has(lower)) {
			push('runtimeMemberNotFound', `CallByName asks Collection '${object[0].rawText}' for '${name}', which it does not have. This will raise Run-time error '438': Object doesn't support this property or method.`, spanOf(span, procName));
		}
		return;
	}
	const surface = (memberCtx.projectClassMembers ?? []).find((candidate) => candidate.kind === 'class' && candidate.name.toLowerCase() === type.toLowerCase());
	if (!surface) {
		return;
	}
	const found = surface.members.find((candidate) => candidate.name.toLowerCase() === lower);
	if (!found) {
		push('runtimeMemberNotFound', `CallByName asks '${object[0].rawText}', a ${surface.name}, for '${name}', which is no Public member of ${surface.name}. This will raise Run-time error '438': Object doesn't support this property or method.`, spanOf(span, procName));
		return;
	}
	if (found.kind === 'method' && kind !== 'method') {
		push('runtimeMemberNotFound', `CallByName reaches ${surface.name}.${found.name}, a ${found.returns ? 'Function' : 'Sub'}, with ${tokenText(callType[0]).startsWith('vb') ? callType[0].rawText : `call type ${callType[0].rawText}`}, which only a property takes. This will raise Run-time error '450': Wrong number of arguments or invalid property assignment.`, spanOf(span, callType));
	}
}

function checkApplicationRun(
	span: Span,
	toks: readonly VbaToken[],
	at: number,
	runnable: ReadonlySet<string>,
	push: PushFn,
): void {
	const macro = argumentsAfter(toks, at)[0];
	if (macro?.length !== 1 || macro[0].kind !== 'stringLiteral') {
		return;
	}
	const name = stringLiteralValue(macro[0].rawText);
	// Another workbook's macro, `Book1.xlsm!Macro`, or a quoted name, is not this project's to judge.
	if (!/^[A-Za-z][A-Za-z0-9_]*(?:\.[A-Za-z][A-Za-z0-9_]*)?$/.test(name) || runnable.has(name.toLowerCase())) {
		return;
	}
	push('runtimeMemberNotFound', `Application.Run names '${name}', and no standard or document module of the project has a Sub or Function of that name. This will raise Run-time error '1004': Cannot run the macro '${name}'.`, spanOf(span, macro));
}
