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
// Issue #408, measured in Excel 16.0 on 2026-10-02: the bare `Run` is
// Application.Run; too many arguments for the procedure named raise 450
// and too few 449, through Run and CallByName alike; VbLet on a Property
// Get alone raises 451; and a Collection's four members are methods, so
// any call type but VbMethod raises 438.
//
// Names compare without case. A name built at run time is not judged.

import type { MemberCompletionContext } from '../../completion/memberAccess';
import type { ConditionalActivityTracker } from '../../conditional/conditionalCompilation';
import type { VbaToken } from '../../lexer/tokenKinds';
import type { ModuleNode, Span } from '../../parser/nodes';
import type { buildModuleSymbols } from '../../symbols/buildModuleSymbols';
import type { PushFn } from '../analysisContext';
import { normalizeType, runtimeSignatureParameterText, splitSignatureTopLevel, stringLiteralValue, typeEnvironmentFor } from '../typeInference';
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
				const toks = statementTokens(source, span);
				for (let i = 0; i < toks.length; i++) {
					const word = tokenText(toks[i]);
					if (word === 'callbyname' && toks[i - 1]?.rawText !== '.' && !moduleNames.has('callbyname')) {
						checkCallByName(span, toks, i, env, memberCtx, push);
					} else if (word === 'run' && runnable && excel && ((toks[i - 1]?.rawText === '.' && tokenText(toks[i - 2]) === 'application' && toks[i - 3]?.rawText !== '.')
						// Excel's global Run is Application.Run (issue #408).
						|| (toks[i - 1]?.rawText !== '.' && !moduleNames.has('run') && !env.has('run')))) {
						checkApplicationRun(span, toks, i, runnable, memberCtx, push);
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
		} else if (kind !== 'method') {
			// Add, Count, Item and Remove are all methods (issue #408, measured).
			push('runtimeMemberNotFound', `CallByName asks Collection '${object[0].rawText}' for '${name}' with ${callType[0].rawText}, but ${name} is a method, which only VbMethod reaches. This will raise Run-time error '438': Object doesn't support this property or method.`, spanOf(span, callType));
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
		return;
	}
	// `CallByName c, "P", VbLet, 5` with P a Property Get alone (issue #408).
	if (found.kind === 'property' && kind === 'let' && !found.letAccessor && !found.setAccessor && found.writable !== true && found.signature !== undefined) {
		push('runtimeMemberNotFound', `CallByName assigns ${surface.name}.${found.name}, which has a Property Get and no Property Let. This will raise Run-time error '451': Property let procedure not defined and property get procedure did not return an object.`, spanOf(span, callType));
		return;
	}
	if (found.kind === 'method' && found.signature) {
		const problem = argumentCountProblem(found.signature, args.length - 3, `${surface.name}.${found.name}`);
		if (problem) {
			push('runtimeMemberNotFound', `CallByName ${problem}`, spanOf(span, procName));
		}
	}
}

/** The parameters a member signature lists: how many a call must pass, and may. */
function parameterCounts(signature: string): { required: number; max: number } | undefined {
	const list = runtimeSignatureParameterText(signature)?.trim();
	if (list === undefined) {
		return undefined;
	}
	const params = list === '' ? [] : splitSignatureTopLevel(list).map((param) => param.trim());
	const required = params.filter((param) => !param.startsWith('[') && !/^paramarray\b/i.test(param)).length;
	return { required, max: params.some((param) => /paramarray/i.test(param)) ? Infinity : params.length };
}

/**
 * Too many arguments for a procedure called by name raise 450, too few 449
 * (issue #408, measured in Excel 16.0 through Application.Run and
 * CallByName). The message reads after "Application.Run " or "CallByName ".
 */
function argumentCountProblem(signature: string, given: number, shown: string): string | undefined {
	const counts = parameterCounts(signature);
	if (!counts) {
		return undefined;
	}
	if (given > counts.max) {
		return `passes ${given} argument${given === 1 ? '' : 's'} to ${shown}, which takes ${counts.max}. This will raise Run-time error '450': Wrong number of arguments or invalid property assignment.`;
	}
	if (given < counts.required) {
		return `passes ${given} argument${given === 1 ? '' : 's'} to ${shown}, which needs ${counts.required}. This will raise Run-time error '449': Argument not optional.`;
	}
	return undefined;
}

/** Whether a name is a cell address in A1 (`Pub2`, `XFD1`) or R1C1 (`R1C1`, `R2`, `C3`) style. */
function readsAsCellAddress(name: string): boolean {
	const a1 = /^([A-Za-z]{1,3})(\d+)$/.exec(name);
	if (a1) {
		const column = [...a1[1].toUpperCase()].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0);
		const row = Number(a1[2]);
		if (column <= 16384 && row >= 1 && row <= 1048576) {
			return true;
		}
	}
	const r1c1 = /^(?:R(\d+)C(\d+)|R(\d+)|C(\d+))$/i.exec(name);
	if (!r1c1) {
		return false;
	}
	const row = Number(r1c1[1] ?? r1c1[3] ?? 1);
	const column = Number(r1c1[2] ?? r1c1[4] ?? 1);
	return row >= 1 && row <= 1048576 && column >= 1 && column <= 16384;
}

function checkApplicationRun(
	span: Span,
	toks: readonly VbaToken[],
	at: number,
	runnable: ReadonlySet<string>,
	memberCtx: MemberCompletionContext,
	push: PushFn,
): void {
	const args = argumentsAfter(toks, at);
	const macro = args[0];
	if (macro?.length !== 1 || macro[0].kind !== 'stringLiteral') {
		return;
	}
	const name = stringLiteralValue(macro[0].rawText);
	// Another workbook's macro, `Book1.xlsm!Macro`, or a quoted name, is not this project's to judge.
	if (!/^[A-Za-z][A-Za-z0-9_]*(?:\.[A-Za-z][A-Za-z0-9_]*)?$/.test(name)) {
		return;
	}
	if (runnable.has(name.toLowerCase())) {
		// The arguments against the one Public procedure of that name.
		const [moduleName, procedure] = name.includes('.') ? name.toLowerCase().split('.') : [undefined, name.toLowerCase()];
		const candidates = (memberCtx.projectClassMembers ?? [])
			.filter((type) => type.kind === 'standardModule' && (moduleName === undefined || type.name.toLowerCase() === moduleName))
			.flatMap((type) => type.members.filter((member) => member.kind === 'method' && member.name.toLowerCase() === procedure));
		// A bare name that reads as a cell address, `Pub2` or `R1C1`, is taken
		// as the address (issue #468, measured in Excel 16.0).
		if (moduleName === undefined && readsAsCellAddress(name)) {
			const owner = candidates.length === 1 ? candidates[0].moduleName : undefined;
			push('runtimeMemberNotFound', `Application.Run reads '${name}' as a cell address, not as the procedure of that name. This will raise Run-time error '1004': Cannot run the macro '${name}'.${owner ? ` Name it with its module: "${owner}.${name}".` : ''}`, spanOf(span, macro));
			return;
		}
		const problem = candidates.length === 1 && candidates[0].signature ? argumentCountProblem(candidates[0].signature, args.length - 1, name) : undefined;
		if (problem) {
			push('runtimeMemberNotFound', `Application.Run ${problem}`, spanOf(span, macro));
		}
		return;
	}
	push('runtimeMemberNotFound', `Application.Run names '${name}', and no standard or document module of the project has a Sub or Function of that name. This will raise Run-time error '1004': Cannot run the macro '${name}'.`, spanOf(span, macro));
}
