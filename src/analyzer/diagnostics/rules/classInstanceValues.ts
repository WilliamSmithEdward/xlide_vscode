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
import type { VbaProjectClassMember, VbaProjectClassMembers } from '../../symbols/symbolModel';
import { procedureSymbolFor, type PushFn } from '../analysisContext';
import { normalizeType } from '../typeInference';
import { activeModuleMembers, forEachStatement, matchParenFrom, statementAndBranchSpans, statementTokensAfterLeadingLabel, tokenName, tokenText } from '../walker';

/** Operators that read a value: beside one, an object member gives its default. */
const VALUE_OPERATORS: ReadonlySet<string> = new Set(['&', '+', '-', '*', '/', '\\', '^', 'mod', '<', '>', '<=', '>=', '<>']);

interface Instance {
	name: string;
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
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind === 'Procedure' && procedureFilter && !procedureFilter(member)) { continue; }
		if (member.kind !== 'Procedure') {
			continue;
		}
		const locals = (procedureSymbolFor(symbols, member)?.children ?? []).filter((child) => child.kind === 'localVariable' && !child.isArray && child.visibility !== 'Static');
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
		const instances = new Map<string, Instance>();
		for (const local of locals) {
			const lower = local.name.toLowerCase();
			const declared = normalizeType(local.asType);
			const sets = statements.filter(({ toks }) => tokenText(toks[0]) === 'set' && tokenName(toks[1])?.toLowerCase() === lower && toks[2]?.rawText === '=');
			let type: VbaProjectClassMembers | undefined;
			if (local.isAutoInstantiated && declared && classes.has(declared) && sets.length === 0) {
				type = classes.get(declared);
			} else if (sets.length === 1 && sets[0].toks.length === 5 && tokenText(sets[0].toks[3]) === 'new' && classes.has(tokenText(sets[0].toks[4]))
				&& (declared === undefined || declared === 'object' || declared === 'variant' || declared === tokenText(sets[0].toks[4]))) {
				type = classes.get(tokenText(sets[0].toks[4]));
			}
			if (!type) {
				continue;
			}
			// Kept to itself: every other mention is `c.Member`.
			const own = statements.every(({ toks }) => toks.every((tok, i) => {
				if (tok.kind !== 'identifier' || tok.rawText.toLowerCase() !== lower || toks[i - 1]?.rawText === '.') {
					return true;
				}
				return toks[i + 1]?.rawText === '.' || (sets.length === 1 && toks === sets[0].toks && i === 1) || tokenText(toks[0]) === 'dim';
			}));
			if (own) {
				instances.set(lower, { name: local.name, type });
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
					assigned.add(`${lower}.${tokenText(toks[first + 2])}`);
				}
			}
		}
		for (const { span, toks } of statements) {
			checkStatement(span, toks, instances, assigned, push);
		}
	}
}

function checkStatement(span: Span, toks: readonly VbaToken[], instances: ReadonlyMap<string, Instance>, assigned: ReadonlySet<string>, push: PushFn): void {
	const head = tokenText(toks[0]);
	for (let i = 0; i + 2 < toks.length; i++) {
		const lower = tokenName(toks[i])?.toLowerCase();
		const instance = lower ? instances.get(lower) : undefined;
		if (!instance || toks[i - 1]?.rawText === '.' || toks[i + 1].rawText !== '.' || !tokenName(toks[i + 2])) {
			continue;
		}
		const name = tokenText(toks[i + 2]);
		const member = instance.type.members.find((candidate) => candidate.name.toLowerCase() === name);
		if (!member) {
			continue;
		}
		const indexed = toks[i + 3]?.rawText === '(';
		const close = indexed ? matchParenFrom([...toks], i + 3) : i + 2;
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
			if (memberOf || indexed || (operand && isObjectType(member, 'object')) || (plainRead && member.kind === 'method')) {
				push('objectVariableNotSet', `${shown} is Nothing here: ${member.kind === 'method' ? `the Function ${member.name} returns nothing else` : `nothing in ${instance.type.name} sets ${member.name}`}. This will raise Run-time error '91': Object variable or With block variable not set.`, at);
				continue;
			}
		}
		if (member.knownValue === 'empty' && !fieldAssigned && memberOf) {
			push('variantValueMisuse', `${shown} is Empty here: nothing in ${instance.type.name} assigns ${member.name}, so it has no members. This will raise Run-time error '424': Object required.`, at);
			continue;
		}
		if (member.knownValue === 'scalar') {
			if (indexed && (target || !memberOf)) {
				push('variantValueMisuse', `${member.name} gives a single value, so ${shown} has no element to ${target ? 'assign' : 'read'}. This will raise Run-time error '13': Type mismatch.`, at);
				continue;
			}
			if (setRead && !indexed) {
				push('variantValueMisuse', `${member.name} gives a single value, not an object, so Set has nothing to assign. This will raise Run-time error '424': Object required.`, at);
				continue;
			}
		}
	}
}

function isObjectType(member: VbaProjectClassMember, type: string): boolean {
	return normalizeType(member.returns) === type;
}
