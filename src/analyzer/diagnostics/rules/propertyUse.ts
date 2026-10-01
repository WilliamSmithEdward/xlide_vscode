// Rule: a property of a project class used in a way its procedures do not
// allow (issue #266). Measured in Excel 16.0 (build 20326, 2026-10-01), on a
// variable declared As the class, each is the compile error "Invalid use of
// property":
//
//  - reading a property that has a Property Let or Set and no Property Get:
//    `Main = c.P`, `Set o = c.P`;
//  - calling a property as a statement: `c.P` with P a Property Get, and
//    `c.P 5` with P a Property Let.
//
// A variable As Object or Variant is late bound: the same read compiles and
// raises 450, which runtime-member-not-found reports.

import type { MemberCompletionContext } from '../../completion/memberAccess';
import type { VbaToken } from '../../lexer/tokenKinds';
import type { Span } from '../../parser/nodes';
import type { PushFn } from '../analysisContext';
import { resolveExactMemberCompletion } from '../typeInference';
import {
	matchParenFrom,
	statementAndBranchSpans,
	statementTokensAfterLeadingLabel,
	tokenName,
	tokenText,
	type ProcedureStatementVisitor,
} from '../walker';

/** Heads whose `=` compares; any other statement's first top-level `=` assigns. */
const COMPARING_HEADS: ReadonlySet<string> = new Set(['if', 'elseif', 'do', 'loop', 'while', 'select', 'case', 'for']);

export function checkInvalidPropertyUse(
	source: string,
	memberCtx: MemberCompletionContext,
	push: PushFn,
): ProcedureStatementVisitor {
	const classes = new Set((memberCtx.projectClassMembers ?? []).filter((type) => type.kind === 'class').map((type) => type.name.toLowerCase()));
	if (classes.size === 0) {
		return () => undefined;
	}
	const check = (span: Span): void => {
		const toks = statementTokensAfterLeadingLabel(source, span).filter((tok) => tok.kind !== 'comment');
		const head = tokenText(toks[0]);
		const first = head === 'call' || head === 'set' || head === 'let' ? 1 : 0;
		const assignAt = COMPARING_HEADS.has(head) ? -1 : topLevelEquals(toks);
		for (let i = first + 1; i + 1 < toks.length; i++) {
			if (toks[i].rawText !== '.' || tokenName(toks[i + 1]) === undefined || tokenName(toks[i - 1]) === undefined) {
				continue;
			}
			const name = toks[i + 1];
			const at: Span = { start: span.start + name.start, end: span.start + name.end };
			const member = resolveExactMemberCompletion(source, name.rawText, at.end, memberCtx);
			// A receiver As Object or Variant resolves no member, so only an
			// early-bound one gets this far.
			if (member?.kind !== 'property' || !classes.has(member.owner.toLowerCase())) {
				continue;
			}
			// `c.P.X` or `c.P(1).X`: a member of what it returns, which the
			// property itself does not decide.
			const close = toks[i + 2]?.rawText === '(' ? matchParenFrom(toks, i + 2) : i + 1;
			const after = toks[close + 1]?.rawText;
			if (close < 0 || after === '.' || after === '!' || (close > i + 1 && after === '(')) {
				continue;
			}
			// `c.P` or `a.b.P` from the statement's start: names and dots in turn.
			const headChain = toks.slice(first, i).every((tok, k) => (k % 2 === 0 ? tokenName(tok) !== undefined : tok.rawText === '.'));
			if (headChain && assignAt > i) {
				continue; // the assignment's target, `c.P = 5` or `c.P(1) = 5`, which the assignment rules judge
			}
			let chainStart = i - 1;
			while (chainStart >= 2 && toks[chainStart - 1].rawText === '.' && tokenName(toks[chainStart - 2]) !== undefined) {
				chainStart -= 2;
			}
			const label = toks.slice(chainStart, i + 2).map((tok) => tok.rawText).join('');
			if (headChain && assignAt < 0 && head !== 'set' && head !== 'let') {
				push('invalidPropertyUse', `'${label}' is a property, and a statement cannot call one. This is a VBE compile error: Invalid use of property.`, at);
				continue;
			}
			if (member.signature === undefined && (member.letAccessor || member.setAccessor)) {
				const setter = member.letAccessor ? 'a Property Let' : 'a Property Set';
				push('invalidPropertyUse', `'${label}' has ${setter} and no Property Get, so it has no value to read. This is a VBE compile error: Invalid use of property.`, at);
			}
		}
	};
	return () => (stmt) => {
		for (const span of statementAndBranchSpans(stmt)) {
			check(span);
		}
	};
}

function topLevelEquals(toks: readonly VbaToken[]): number {
	let depth = 0;
	for (let i = 0; i < toks.length; i++) {
		const raw = toks[i].rawText;
		if (raw === '(') {
			depth++;
		} else if (raw === ')') {
			depth--;
		} else if (raw === '=' && depth === 0 && toks[i].kind === 'operator') {
			return i;
		}
	}
	return -1;
}
