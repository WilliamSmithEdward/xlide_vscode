import { signatureDeclaresParameters } from '../../completion/memberAccess';
// Rule: a member of a project class instance used where what it holds, or
// what it is, cannot serve (issue #414). Each case measured in Excel 16.0 on
// an instance the procedure itself makes, `Dim c As New Class1` or
// `Set c = New Class1`:
//
//   c.M.Add 1, c.M(1), c.M & "x" with M an object field never set   91
//   Main = c.M with M a Function that returns Nothing               91
//   Main = c.M(1), c.M(1) = 2 with M a Get returning 1              13
//   Set o = c.M with the same Get                                   424
//   c.M.Count with M a Variant field never assigned                 424
//
// Late-bound misuse of a member, a Sub assigned or a Set-only property
// read, is runtime-member-not-found's.
//
// The instance is followed only while the procedure keeps it to itself:
// any use of it other than `c.Member`, or a second Set, ends the rule's
// reading of it. A field assigned through it, `Set c.M = ...`, is not
// judged.

import { topLevelEqualsIndex } from '../../lexer/tokenHelpers';
import type { MemberCompletionContext } from '../../completion/memberAccess';
import type { ConditionalActivityTracker } from '../../conditional/conditionalCompilation';
import type { VbaToken } from '../../lexer/tokenKinds';
import type { ModuleNode, Span , ProcedureNode } from '../../parser/nodes';
import type { buildModuleSymbols } from '../../symbols/buildModuleSymbols';
import type { VbaProjectClassMember, VbaProjectClassMembers, VbaSymbol } from '../../symbols/symbolModel';
import { procedureSymbolFor, type PushFn } from '../analysisContext';
import { normalizeType } from '../typeInference';
import { activeModuleMembers, forEachStatement, matchParenFrom, statementAndBranchSpans, statementTokensAfterLeadingLabel, tokenName, tokenText } from '../walker';

/** Operators that read a value: beside one, an object member gives its default. */
const VALUE_OPERATORS: ReadonlySet<string> = new Set(['&', '+', '-', '*', '/', '\\', '^', 'mod', '<', '>', '<=', '>=', '<>']);

interface Instance {
	type: VbaProjectClassMembers;
}

export function checkClassInstanceValues(
	source: string,
	mod: ModuleNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
	memberCtx: MemberCompletionContext,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
	procedureFilter?: (member: ProcedureNode) => boolean,
): void {
	const classes = new Map((memberCtx.projectClassMembers ?? []).filter((type) => type.kind === 'class').map((type) => [type.name.toLowerCase(), type]));
	if (classes.size === 0) {
		return;
	}
	// Only consulted class surfaces need an index. Keep it within this query so
	// later project metadata changes cannot reuse stale members.
	const memberIndexes = new Map<VbaProjectClassMembers, ReadonlyMap<string, VbaProjectClassMember>>();
	const findMember = (type: VbaProjectClassMembers, name: string): VbaProjectClassMember | undefined => {
		// A bounded linear scan avoids allocating an index for tiny classes.
		if (type.members.length <= 8) {
			return type.members.find((candidate) => candidate.name.toLowerCase() === name);
		}
		let index = memberIndexes.get(type);
		if (!index) {
			const built = new Map<string, VbaProjectClassMember>();
			for (const candidate of type.members) {
				const key = candidate.name.toLowerCase();
				if (!built.has(key)) { built.set(key, candidate); }
			}
			index = built;
			memberIndexes.set(type, index);
		}
		return index.get(name);
	};
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind === 'Procedure' && procedureFilter && !procedureFilter(member)) { continue; }
		if (member.kind !== 'Procedure') {
			continue;
		}
		const locals: Array<{ local: VbaSymbol; declared: string | undefined }> = [];
		for (const local of procedureSymbolFor(symbols, member)?.children ?? []) {
			if (local.kind !== 'localVariable' || local.isArray || local.visibility === 'Static') { continue; }
			const declared = normalizeType(local.asType);
			if (declared !== undefined && declared !== 'object' && declared !== 'variant' && !classes.has(declared)) { continue; }
			locals.push({ local, declared });
		}
		if (locals.length === 0) {
			continue;
		}
		// Every statement of the procedure, its single-line If arms apart.
		const statements: Array<{ span: Span; toks: VbaToken[] }> = [];
		forEachStatement(member.body, (stmt) => {
			for (const span of statementAndBranchSpans(stmt)) {
				statements.push({ span, toks: statementTokensAfterLeadingLabel(source, span) });
			}
		}, activity);
		// Set targets and non-member uses are properties of each statement, not
		// of each local. Index them once for all candidate instances.
		const names = new Set(locals.map(({ local }) => local.name.toLowerCase()));
		const setsByName = new Map<string, typeof statements>();
		const escaped = new Set<string>();
		for (const statement of statements) {
			const { toks } = statement;
			const head = tokenText(toks[0]);
			const target = head === 'set' && toks[2]?.rawText === '=' ? tokenName(toks[1])?.toLowerCase() : undefined;
			if (target && names.has(target)) {
				const sets = setsByName.get(target);
				if (sets) { sets.push(statement); } else { setsByName.set(target, [statement]); }
			}
			if (head === 'dim') { continue; }
			for (let i = 0; i < toks.length; i++) {
				const tok = toks[i];
				if ((tok.kind !== 'identifier' && tok.kind !== 'bracketedIdentifier') || toks[i - 1]?.rawText === '.') { continue; }
				const name = tokenName(tok)!.toLowerCase();
				if (names.has(name) && toks[i + 1]?.rawText !== '.' && !(i === 1 && target === name)) { escaped.add(name); }
			}
		}
		const instances = new Map<string, Instance>();
		for (const { local, declared } of locals) {
			const lower = local.name.toLowerCase();
			const sets = setsByName.get(lower);
			const set = sets?.length === 1 ? sets[0] : undefined;
			let type: VbaProjectClassMembers | undefined;
			if (local.isAutoInstantiated && declared && classes.has(declared) && !sets?.length) {
				type = classes.get(declared);
			} else if (set && set.toks.length === 5 && tokenText(set.toks[3]) === 'new' && classes.has(tokenText(set.toks[4]))
				&& (declared === undefined || declared === 'object' || declared === 'variant' || declared === tokenText(set.toks[4]))) {
				type = classes.get(tokenText(set.toks[4]));
			}
			if (!type) {
				continue;
			}
			if (!escaped.has(lower)) {
				instances.set(lower, { type });
			}
		}
		if (instances.size === 0) {
			continue;
		}
		// Fields the procedure assigns through the instance hold what it gave.
		const assigned = new Set<string>();
		for (const { toks } of statements) {
			const first = tokenText(toks[0]) === 'set' || tokenText(toks[0]) === 'let' ? 1 : 0;
			const lower = tokenName(toks[first])?.toLowerCase();
			if (lower && instances.has(lower) && toks[first + 1]?.rawText === '.' && tokenName(toks[first + 2])) {
				const close = toks[first + 3]?.rawText === '(' ? matchParenFrom(toks, first + 3) : first + 2;
				if (toks[close + 1]?.rawText === '=') {
					assigned.add(`${lower}.${tokenName(toks[first + 2])!.toLowerCase()}`);
				}
			}
		}
		for (const { span, toks } of statements) {
			checkStatement(span, toks, instances, assigned, findMember, push);
		}
	}
}

function checkStatement(span: Span, toks: readonly VbaToken[], instances: ReadonlyMap<string, Instance>, assigned: ReadonlySet<string>, findMember: (type: VbaProjectClassMembers, name: string) => VbaProjectClassMember | undefined, push: PushFn): void {
	const head = tokenText(toks[0]);
	for (let i = 0; i + 2 < toks.length; i++) {
		const lower = tokenName(toks[i])?.toLowerCase();
		const instance = lower ? instances.get(lower) : undefined;
		if (!instance || toks[i - 1]?.rawText === '.' || toks[i + 1].rawText !== '.' || !tokenName(toks[i + 2])) {
			continue;
		}
		const name = tokenName(toks[i + 2])!.toLowerCase();
		const member = findMember(instance.type, name);
		if (!member) {
			continue;
		}
		const indexed = toks[i + 3]?.rawText === '(';
		const indexesResult = indexed && !signatureDeclaresParameters(member.signature);
		const close = indexed ? matchParenFrom(toks, i + 3) : i + 2;
		if (close < 0) {
			continue;
		}
		const after = toks[close + 1];
		const before = toks[i - 1];
		const first = head === 'set' || head === 'let' ? 1 : 0;
		const target = i === first && after?.rawText === '=';
		const setRead = head === 'set' && toks[2]?.rawText === '=' && i === 3 && close === toks.length - 1;
		const plainRead = head !== 'set' && i >= 2 && toks[i - 1].rawText === '=' && i - 1 === topLevelEqualsIndex(toks) && close === toks.length - 1;
		const memberOf = after?.rawText === '.' && !indexed;
		const label = toks.slice(i, close + 1).map((tok) => tok.rawText).join('');
		const at: Span = { start: span.start + toks[i].start, end: span.start + toks[close].end };
		const shown = `'${label}'`;
		const fieldAssigned = assigned.has(`${lower}.${name}`);
		// What it holds, through any binding.
		if (member.knownValue === 'nothing' && !fieldAssigned && !target) {
			const operand = (after && VALUE_OPERATORS.has(tokenText(after) || after.rawText)) || (before && VALUE_OPERATORS.has(tokenText(before) || before.rawText));
			if (memberOf || indexesResult || (operand && isObjectType(member, 'object')) || (plainRead && member.kind === 'method')) {
				push('objectVariableNotSet', `${shown} is Nothing here: ${member.kind === 'method' ? `the Function ${member.name} returns nothing else` : `nothing in ${instance.type.name} sets ${member.name}`}. This will raise Run-time error '91': Object variable or With block variable not set.`, at);
				continue;
			}
		}
		if (member.knownValue === 'empty' && member.signature === undefined && !fieldAssigned && memberOf) {
			push('variantValueMisuse', `${shown} is Empty here: nothing in ${instance.type.name} assigns ${member.name}, so it has no members. This will raise Run-time error '424': Object required.`, at);
			continue;
		}
		if (member.knownValue === 'scalar') {
			if (indexesResult && (target || !memberOf)) {
				push('variantValueMisuse', `${member.name} gives a single value, so ${shown} has no element to ${target ? 'assign' : 'read'}. This will raise Run-time error '13': Type mismatch.`, at);
				continue;
			}
			if (setRead && !indexesResult) {
				push('variantValueMisuse', `${member.name} gives a single value, not an object, so Set has nothing to assign. This will raise Run-time error '424': Object required.`, at);
				continue;
			}
		}
	}
}

function isObjectType(member: VbaProjectClassMember, type: string): boolean {
	return normalizeType(member.returns) === type;
}
