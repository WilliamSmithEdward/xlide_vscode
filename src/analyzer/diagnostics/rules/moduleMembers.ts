// Rule family: a standard module's member used against its kind, through the
// module's name or bare (issue #423). Each is a compile error, measured in
// Excel 16.0 (build 20430, 2026-10-02) with M in Module2 and the use in
// Module1:
//
//  - `Module2.M = 9` with M a Sub: "Expected Function or variable"; with M a
//    Function or Declare returning a type of VBA's own: "Function call on
//    left-hand side of assignment must return Variant or Object".
//  - `Main = Module2.M` with M a Sub: "Expected Function or variable".
//  - `Module2.M` or `Call Module2.M` with M a variable, a Const or an Enum
//    member: "Expected procedure, not variable".
//  - `M`, `Call M`, `Module2.M` or `Call Module2.M` with M a Property Get and
//    no Let or Set: "Invalid use of property"; `M = 9`: "Can't assign to
//    read-only property".
//  - `M = 9` or `Main = M` with M a Type: "Variable not defined", under
//    Option Explicit.
//
// The bare forms of the first three are judged elsewhere already.

import type { MemberCompletionContext } from '../../completion/memberAccess';
import type { VbaToken } from '../../lexer/tokenKinds';
import type { Span } from '../../parser/nodes';
import type { buildModuleSymbols } from '../../symbols/buildModuleSymbols';
import type { VbaSymbol } from '../../symbols/symbolModel';
import { procedureSymbolFor, type PushFn } from '../analysisContext';
import { defTypeOf, getterMayReturnObject, isKnownScalarType, normalizeType, sourceIdentifierBinding } from '../typeInference';
import { statementAndBranchSpans, statementTokensAfterLeadingLabel, tokenName, tokenText, type ProcedureStatementVisitor } from '../walker';

/** The kinds of symbol a bare name reads as a value. */
const VALUE_KINDS: ReadonlySet<string> = new Set(['moduleVariable', 'constant', 'function', 'propertyGet', 'declare']);

const COMPARING_HEADS: ReadonlySet<string> =new Set(['if', 'elseif', 'do', 'loop', 'while', 'select', 'case', 'for']);

/** The symbols named `lower` that a standard module declares, Public or by default. */
function moduleMember(moduleName: string, lower: string, symbols: ReturnType<typeof buildModuleSymbols>, visible: readonly VbaSymbol[]): VbaSymbol[] {
	const own = symbols.moduleName.toLowerCase() === moduleName.toLowerCase();
	const pool = own ? symbols.all.filter((sym) => !sym.containerName || sym.kind === 'enumMember') : visible.filter((sym) => sym.moduleName.toLowerCase() === moduleName.toLowerCase());
	return pool.filter((sym) => sym.name.toLowerCase() === lower && sym.kind !== 'localVariable' && sym.kind !== 'parameter'
		&& (own || (sym.visibility !== 'Private' && sym.visibility !== 'Dim')));
}

/** Whether every symbol of the name is a Property Get, Let or Set. */
function propertyOnly(found: readonly VbaSymbol[]): boolean {
	return found.length > 0 && found.every((sym) => sym.kind === 'propertyGet' || sym.kind === 'propertyLet' || sym.kind === 'propertySet');
}

function getOnly(found: readonly VbaSymbol[]): boolean {
	return found.length > 0 && found.some((sym) => sym.kind === 'propertyGet') && found.every((sym) => sym.kind === 'propertyGet');
}

/** Why a Sub, or a Function of a type of VBA's own, cannot be assigned. */
function procedureAssignment(found: readonly VbaSymbol[], name: string): string | undefined {
	if (found.length !== 1) {
		return undefined;
	}
	const [sym] = found;
	if (sym.kind === 'sub' || (sym.kind === 'declare' && sym.declareKind === 'Sub')) {
		return `'${name}' is a Sub, which has no value to assign. This is a VBE compile error: Expected Function or variable.`;
	}
	const returns = normalizeType(sym.asType);
	const isFunction = sym.kind === 'function' || (sym.kind === 'declare' && sym.declareKind === 'Function');
	return isFunction && returns && returns !== 'variant' && isKnownScalarType(returns) && !sym.isArray
		? `'${name}' is a Function returning ${sym.asType}, and a call cannot be assigned to. This is a VBE compile error: Function call on left-hand side of assignment must return Variant or Object.`
		: undefined;
}

export function checkModuleMemberForms(
	source: string,
	symbols: ReturnType<typeof buildModuleSymbols>,
	memberCtx: MemberCompletionContext,
	projectVisibleSymbols: readonly VbaSymbol[] | undefined,
	explicit: boolean,
	push: PushFn,
): ProcedureStatementVisitor {
	const modules = new Set((memberCtx.projectClassMembers ?? []).filter((type) => type.kind === 'standardModule').map((type) => type.name.toLowerCase()));
	if (symbols.moduleKind === 'standard') {
		modules.add(symbols.moduleName.toLowerCase());
	}
	const visible = projectVisibleSymbols ?? [];
	// These tests depend on declaration names, not on the statement. Index the
	// candidate sets once; most assignments cannot name a Type or property.
	const typeNames = new Set<string>();
	const propertyNames = new Set<string>();
	for (const sym of [...(symbols.root.children ?? []), ...visible]) {
		const lower = sym.name.toLowerCase();
		if (sym.kind === 'type') { typeNames.add(lower); }
		if (modules.has(sym.moduleName.toLowerCase())
			&& (sym.kind === 'propertyGet' || sym.kind === 'propertyLet' || sym.kind === 'propertySet')) {
			propertyNames.add(lower);
		}
	}
	const externalValueNames = new Set(visible.filter(sym => VALUE_KINDS.has(sym.kind)
		&& sym.visibility !== 'Private' && sym.moduleName.toLowerCase() !== symbols.moduleName.toLowerCase())
		.map(sym => sym.name.toLowerCase()));
	return (member) => {
		const procSym = procedureSymbolFor(symbols, member);
		const locals = new Set([member.name, ...member.params.map((param) => param.name), ...(procSym?.children ?? []).map((child) => child.name)].map((name) => name.toLowerCase()));
		const isModule = (tok: VbaToken | undefined): boolean => {
			const lower = tokenName(tok)?.toLowerCase();
			return lower !== undefined && modules.has(lower) && !locals.has(lower);
		};
		const at = (span: Span, from: VbaToken, to: VbaToken): Span => ({ start: span.start + from.start, end: span.start + to.end });
		const check = (span: Span): void => {
			const toks = statementTokensAfterLeadingLabel(source, span);
			const head = tokenText(toks[0]);
			const first = head === 'call' ? 1 : 0;
			const assignAt = COMPARING_HEADS.has(head) || head === 'set' ? -1 : toks.findIndex((tok) => tok.rawText === '=' && tok.kind === 'operator');
			// `Module2.M` at the statement's start: called, or assigned.
			if (isModule(toks[first]) && toks[first + 1]?.rawText === '.' && tokenName(toks[first + 2]) !== undefined) {
				const name = toks[first + 2];
				const found = moduleMember(tokenName(toks[first])!, tokenName(name)!.toLowerCase(), symbols, visible);
				const label = `${toks[first].rawText}.${name.rawText}`;
				const where = at(span, toks[first], name);
				if (first === 0 && assignAt === first + 3) {
					const why = procedureAssignment(found, label);
					if (why) {
						push('assignmentToProcedureName', why, where);
					}
					return;
				}
				const statementCall = assignAt < 0 && toks[first + 3]?.rawText !== '.' && toks[first + 3]?.rawText !== '!';
				if (statementCall && found.length === 1 && ['moduleVariable', 'constant', 'enumMember'].includes(found[0].kind)) {
					const kind = found[0].kind === 'moduleVariable' ? 'a module variable' : found[0].kind === 'constant' ? 'a constant' : 'an Enum member';
					push('nonCallableCallStatement', `Cannot call '${label}' because it resolves to ${kind}, not a Sub or Function. This is a VBE compile error: Expected procedure, not variable.`, where);
					return;
				}
				if (statementCall && propertyOnly(found)) {
					push('invalidPropertyUse', `'${label}' is a property, and a statement cannot call one. This is a VBE compile error: Invalid use of property.`, where);
				}
				return;
			}
			// `Main = Module2.M` with M a Sub: the whole value.
			if (assignAt > 0 && toks.length === assignAt + 4 && isModule(toks[assignAt + 1]) && toks[assignAt + 2]?.rawText === '.') {
				const name = toks[assignAt + 3];
				const found = moduleMember(tokenName(toks[assignAt + 1])!, tokenName(name)?.toLowerCase() ?? '', symbols, visible);
				if (found.length === 1 && found[0].kind === 'sub') {
					push('subUsedAsValue', `'${toks[assignAt + 1].rawText}.${name.rawText}' is a Sub, which returns nothing, so it cannot be used as a value. This is a VBE compile error: Expected Function or variable.`, at(span, toks[assignAt + 1], name));
				}
				return;
			}
			// Bare `M`, `Call M` and `M = 9`, with M a standard module's Property Get.
			const bare = tokenName(toks[first]);
			if (!bare || toks[first + 1]?.rawText === '.' || toks[first + 1]?.rawText === '(') {
				return;
			}
			if (propertyNames.has(bare.toLowerCase())) {
				const binding = sourceIdentifierBinding(symbols, procSym, projectVisibleSymbols, bare, 'call');
				const inModule = binding.definitions.filter((sym) => modules.has(sym.moduleName.toLowerCase()));
				// Inside the property, its name is its value: `M = 1` in Get M.
				const own = bare.toLowerCase() === member.name.toLowerCase();
				if (!own && binding.scope !== 'ambiguous' && inModule.length === binding.definitions.length && propertyOnly(inModule)) {
					const where = at(span, toks[first], toks[first]);
					const getter = inModule.find(sym => sym.kind === 'propertyGet');
					const getterType = getter?.asType ?? (getter?.moduleName.toLowerCase() === symbols.moduleName.toLowerCase() ? defTypeOf(symbols, getter.name) : undefined);
					if (assignAt === first + 1 && first === 0 && getOnly(inModule) && !getterMayReturnObject(getterType, memberCtx)) {
						push('readonlyMemberAssignment', `'${bare}' has a Property Get and no Property Let, so it cannot be assigned. This is a VBE compile error: Can't assign to read-only property.`, where);
					} else if (assignAt < 0 && head !== 'set') {
						push('invalidPropertyUse', `'${bare}' is a property, and a statement cannot call one. This is a VBE compile error: Invalid use of property.`, where);
					}
					return;
				}
			}
			// `M = 9` and `Main = M` with M a Type: no variable of that name.
			if (!explicit || assignAt < 0) {
				return;
			}
			// `TypeName(M)` reads M as a value too (issue #639).
			const typeNameArguments = toks.flatMap((_, k) => (tokenText(toks[k - 2]) === 'typename' && toks[k - 1]?.rawText === '(' && toks[k + 1]?.rawText === ')' ? [k] : []));
			for (const index of [0, assignAt + 1, ...typeNameArguments]) {
				const tok = toks[index];
				const name = tokenName(tok);
				const asArgument = typeNameArguments.includes(index);
				if (!name || (!asArgument && ((index === 0 && assignAt !== 1) || (index > 0 && toks.length !== assignAt + 2)))) {
					continue;
				}
				// A variable, Const or Function of the name elsewhere in the
				// project is what the name reads: `Public Zq As Long` in Module2
				// beside a Private Type Zq here runs (issue #639, measured in
				// Excel 16.0).
				const lower = name.toLowerCase();
				if (!typeNames.has(lower) || externalValueNames.has(lower)) {
					continue;
				}
				const found = sourceIdentifierBinding(symbols, procSym, projectVisibleSymbols, name, index === 0 ? 'assignmentTarget' : 'expression');
				if (found.scope === 'unresolved' || (found.definitions.length > 0 && found.definitions.every((sym) => sym.kind === 'type'))) {
					push('undeclaredVariable', `Variable not defined: '${name}'. It names a user-defined type, not a variable. This is a VBE compile error.`, at(span, tok, tok));
				}
			}
		};
		return (stmt) => {
			for (const span of statementAndBranchSpans(stmt)) {
				check(span);
			}
		};
	};
}
