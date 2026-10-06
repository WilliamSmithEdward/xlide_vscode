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

import type { MemberCompletion, MemberCompletionContext } from '../../completion/memberAccess';
import type { VbaToken } from '../../lexer/tokenKinds';
import type { Span } from '../../parser/nodes';
import type { PushFn } from '../analysisContext';
import type { DiagnosticRuleName } from '../ruleMetadata';
import { isKnownScalarType, normalizeType, resolveExactMemberCompletion, runtimeSignatureParameterText, splitSignatureTopLevel } from '../typeInference';
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
	const check = (span: Span, conditionOnly: boolean): void => {
		const all = statementTokensAfterLeadingLabel(source, span);
		// A single-line If is read up to Then: its branches are statements of
		// their own, checked as such.
		const toks = conditionOnly ? all.slice(0, all.findIndex((tok) => tokenText(tok) === 'then') + 1) : all;
		const head = tokenText(toks[0]);
		const first = head === 'call' || head === 'set' || head === 'let' ? 1 : 0;
		const assignAt = COMPARING_HEADS.has(head) ? -1 : topLevelEquals(toks);
		// A With member assigned, `.M = 1`, stands at the statement's start.
		const withStart = toks[first]?.rawText === '.' && assignAt > first;
		for (let i = withStart ? first : first + 1; i + 1 < toks.length; i++) {
			if (toks[i].rawText !== '.' || tokenName(toks[i + 1]) === undefined || (i > first && tokenName(toks[i - 1]) === undefined) || (i === first && !withStart)) {
				continue;
			}
			// `If .State("q") > 3`, `Case .Count`, `Debug.Print .Count`: a With
			// member after a keyword, which is no receiver (issue #413). Me is.
			if (i > first && toks[i - 1].kind === 'keyword' && tokenText(toks[i - 1]) !== 'me') {
				continue;
			}
			const name = toks[i + 1];
			const at: Span = { start: span.start + name.start, end: span.start + name.end };
			const member = resolveExactMemberCompletion(source, name.rawText, at.end, memberCtx);
			// A receiver As Object or Variant resolves no member, so only an
			// early-bound one gets this far.
			if (!member || !classes.has(member.owner.toLowerCase())) {
				continue;
			}
			const close = toks[i + 2]?.rawText === '(' ? matchParenFrom(toks, i + 2) : i + 1;
			const after = toks[close + 1]?.rawText;
			// `a.b.M 1` or `.b.M 1`: names and dots from the statement's start.
			// `a.b.M 1`, `.b(2).M 1`: names, each maybe indexed, and dots from `from` to the member.
			const chainFrom = (from: number): boolean => {
				let k = from;
				while (k < i && tokenName(toks[k]) !== undefined) {
					k++;
					if (toks[k]?.rawText === '(') {
						const shut = matchParenFrom(toks, k);
						if (shut < 0) {
							return false;
						}
						k = shut + 1;
					}
					if (k === i) {
						return true;
					}
					if (toks[k]?.rawText !== '.') {
						return false;
					}
					k++;
				}
				return false;
			};
			const headCall = assignAt < 0 && head !== 'set' && head !== 'let' && (chainFrom(first) || (toks[first]?.rawText === '.' && chainFrom(first + 1)));
			// `o.M .Left + 1`: in a call statement, a dot after a space starts
			// an argument, a With member, and takes no member of M.
			const argumentDot = headCall && after === '.' && close >= 0 && toks[close + 1].start > toks[close].end;
			const misuse = close < 0 ? undefined : memberMisuse(member, {
				indexed: close > i + 1,
				after: argumentDot ? undefined : after,
				target: tokenText(toks[first]) !== '.' && toks.slice(first, i).every((tok, k) => (k % 2 === 0 ? tokenName(tok) !== undefined : tok.rawText === '.')) && assignAt === close + 1,
				setTarget: head === 'set' && assignAt === close + 1,
				statementCall: headCall && (argumentDot || (after !== '.' && after !== '!' && after !== '(')),
				setRead: head === 'set' && assignAt >= 0 && assignAt < i,
				withTarget: i === first && assignAt === close + 1,
			});
			if (misuse) {
				push(misuse.rule, `'${toks.slice(Math.max(first, i - 1), i + 2).map((tok) => tok.rawText).join('')}' ${misuse.message}`, at);
				continue;
			}
			if (member.kind !== 'property' || i === first) {
				continue;
			}
			// `c.P.X` or `c.P(1).X`: a member of what it returns, which the
			// property itself does not decide.
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
			check(span, span === stmt.span && stmt.kind === 'Statement' && stmt.singleLineIfBranches !== undefined);
		}
	};
}

interface MemberUse {
	/** Written with an argument list: `c.M(1)`. */
	indexed: boolean;
	/** The token after the member and its argument list. */
	after: string | undefined;
	/** The statement assigns to it: `c.M = 5`, `c.M(1) = 2`. */
	target: boolean;
	/** The target of a Set: `Set c.M = x`. */
	setTarget: boolean;
	/** Read whole by a Set: `Set o = c.M`. */
	setRead: boolean;
	/** Called as a statement: `c.M`, `Call c.M(1)`. */
	statementCall: boolean;
	/** A With member assigned: `.M = 1`. */
	withTarget: boolean;
}

/**
 * What the VBE refuses in a use of a project class's member through a
 * variable declared As the class (issue #414, each measured in Excel 16.0):
 * a Sub used as a value or assigned; a member of a scalar's value, or Is
 * Nothing on it; an argument list on a Property Get that takes none; a
 * Function or Get whose argument is left out; an element of a String Get
 * assigned; a member of a property with no Get; and a
 * Collection field beside an operator.
 */
function memberMisuse(member: MemberCompletion, use: MemberUse): { rule: DiagnosticRuleName; message: string } | undefined {
	const writes = use.target || use.withTarget;
	if (member.kind === 'method' && member.sub) {
		// A plain read, `Main = c.M`, is the value rules' (issue #369).
		if (writes || use.setTarget || use.setRead || use.after === '.') {
			return { rule: 'subUsedAsValue', message: 'is a Sub, which gives no value and takes no assignment. This is a VBE compile error: Expected Function or variable.' };
		}
		return undefined;
	}
	const params = parameterCounts(member.signature);
	const type = normalizeType(member.returns ?? member.declaredType);
	const scalar = type !== undefined && type !== 'variant' && isKnownScalarType(type);
	// `c.M = 9` with M a Function returning Long (issue #423).
	if (member.kind === 'method' && scalar && use.target && !use.indexed) {
		return { rule: 'assignmentToProcedureName', message: `is a Function returning ${capitalized(type!)}, and a call cannot be assigned to. This is a VBE compile error: Function call on left-hand side of assignment must return Variant or Object.` };
	}
	// `c.M = 5` or `.M = 1` calls M and assigns to what it returns: a
	// Variant holding a value raises 424, and a Collection, whose Item needs
	// an index, does not compile (issue #414, measured in Excel 16.0).
	// Known to return Empty or a value: one holding an object raises 438
	// instead, and is left alone.
	const holdsValue = member.knownValue === 'empty' || member.knownValue === 'scalar';
	if (member.kind === 'method' && writes && !use.indexed && params.required === 0) {
		if ((type === undefined || type === 'variant') && holdsValue) {
			return { rule: 'variantValueMisuse', message: "is a Function, so the assignment calls it and assigns to the Variant it returns, which holds no object. This will raise Run-time error '424': Object required." };
		}
		if (type === 'collection') {
			return { rule: 'argumentCount', message: 'is a Function returning a Collection, so the assignment reaches its default member Item, which needs an index. This is a VBE compile error: Argument not optional.' };
		}
	}
	if (member.kind === 'property' && member.signature === undefined && (member.letAccessor || member.setAccessor)) {
		// `c.M(1) = 2` with M a Property Set and no Let (issue #414).
		if (writes && !use.setTarget && use.indexed && member.setAccessor && !member.letAccessor) {
			return { rule: 'invalidPropertyUse', message: 'has a Property Set and no Property Let, so a value cannot be assigned to it. This is a VBE compile error: Invalid use of property.' };
		}
		if (use.after === '.') {
			return { rule: 'invalidPropertyUse', message: `has ${member.letAccessor ? 'a Property Let' : 'a Property Set'} and no Property Get, so it has no value to take a member of. This is a VBE compile error: Invalid use of property.` };
		}
		// `c.M(1) = 2` with a Let that takes only the value (issue #414,
		// measured in Excel 16.0).
		if (use.target && use.indexed && member.letAccessor && member.letParamCount === 1) {
			return { rule: 'invalidPropertyUse', message: 'has a Property Let that takes no index and no Property Get, so an element of it cannot be assigned. This is a VBE compile error: Invalid use of property.' };
		}
		return undefined;
	}
	// A plain read of a property is argument-count's (issue #224).
	if (params.required > 0 && !use.indexed && !writes && !use.setTarget && !use.statementCall && (member.kind === 'method' || use.setRead || use.after === '.')) {
		return { rule: 'argumentCount', message: `needs ${params.required === 1 ? 'an argument' : `${params.required} arguments`}, and is read here with none. This is a VBE compile error: Argument not optional.` };
	}
	if (scalar && !use.indexed && use.after === '.') {
		return { rule: 'scalarMemberAccess', message: `holds ${article(type!)} ${capitalized(type!)}, which has no members. This is a VBE compile error: Invalid qualifier.` };
	}
	if (scalar && !use.indexed && use.after?.toLowerCase() === 'is') {
		return { rule: 'isOperatorNonObject', message: `holds ${article(type!)} ${capitalized(type!)}, which Is cannot compare with an object. This is a VBE compile error: Type mismatch.` };
	}
	const getOnly = member.kind === 'property' && member.signature !== undefined && !member.letAccessor && !member.setAccessor && !member.writable;
	if (getOnly && scalar && use.indexed && params.total === 0) {
		if (use.target && type === 'string') {
			return { rule: 'readonlyMemberAssignment', message: 'has a Property Get and no Property Let, so an element of it cannot be assigned. This is a VBE compile error: Can\'t assign to read-only property.' };
		}
		if (!writes) {
			return { rule: 'argumentCount', message: 'is a Property Get that takes no argument, so it cannot be given one. This is a VBE compile error: Wrong number of arguments or invalid property assignment.' };
		}
	}
	if (member.kind === 'property' && member.signature === undefined && type === 'collection' && !use.indexed && use.after !== undefined && SCALAR_OPERATOR_TEXT.has(use.after)) {
		return { rule: 'collectionOperand', message: `is a Collection: its default member Item needs an index, so '${use.after}' has no value to work on. This is a VBE compile error: Argument not optional.` };
	}
	return undefined;
}

const SCALAR_OPERATOR_TEXT: ReadonlySet<string> = new Set(['&', '+', '-', '*', '/', '\\', '^']);

/** How many parameters a source signature declares, and how many a call must pass. */
function parameterCounts(signature: string | undefined): { total: number; required: number } {
	const inner = signature === undefined ? undefined : runtimeSignatureParameterText(signature)?.trim();
	if (!inner) {
		return { total: 0, required: 0 };
	}
	const parts = splitSignatureTopLevel(inner).map((part) => part.trim());
	return { total: parts.length, required: parts.filter((part) => !/^(Optional|ParamArray)\b/i.test(part) && !part.startsWith('[')).length };
}

function article(type: string): string {
	return /^[aeiou]/i.test(type) ? 'an' : 'a';
}

function capitalized(type: string): string {
	return type === 'longlong' ? 'LongLong' : type === 'longptr' ? 'LongPtr' : type.charAt(0).toUpperCase() + type.slice(1);
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
