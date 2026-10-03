// Rule: a ParamArray used where the VBE refuses it (issue #445). Each is a
// compile error, measured in Excel 16.0 (build 20430, 2026-10-02), inside
// `Function F(ParamArray p() As Variant)`:
//
//  - `ReDim p(3)` and `Erase p`: "Invalid ParamArray use".
//  - `G(p)`, `G((p))`, `Call S(p)` and `S p` with G's or S's parameter ByRef,
//    a Variant or an array: "Invalid ParamArray use". A ByVal parameter and
//    another ParamArray take it.
//  - `Set p = Nothing`: "Can't assign to array".
//  - `Function F(ParamArray p)`, declared with no parentheses: refused.
//
// `p = Array(1)`, `p(0) = 9`, `UBound(p)` and `For Each` over p compile.

import type { VbaToken } from '../../lexer/tokenKinds';
import type { ProcedureNode, Span } from '../../parser/nodes';
import type { PushFn } from '../analysisContext';
import type { CallableTypeSignature } from '../callExtraction';
import { matchParenFrom, statementAndBranchSpans, statementTokensAfterLeadingLabel, tokenName, tokenText, type ProcedureStatementVisitor } from '../walker';
import { splitTopLevelTokenGroups } from '../../lexer/tokenHelpers';

export function checkParamArrayUse(
	source: string,
	signatures: ReadonlyMap<string, CallableTypeSignature>,
	push: PushFn,
): ProcedureStatementVisitor {
	return (member: ProcedureNode) => {
		const param = member.params.find((p) => p.paramArray);
		if (!param) {
			return undefined;
		}
		const lower = param.name.toLowerCase();
		if (!param.isArray) {
			push('invalidParamArrayUse', `ParamArray '${param.name}' is declared with no parentheses; a ParamArray is an array of Variant, ${param.name}(). This is a VBE compile error.`, param.nameSpan ?? param.span);
		}
		const own = new Set(member.params.map((p) => p.name.toLowerCase()));
		const at = (span: Span, tok: VbaToken): Span => ({ start: span.start + tok.start, end: span.start + tok.end });
		return (stmt) => {
			for (const span of statementAndBranchSpans(stmt)) {
				const toks = statementTokensAfterLeadingLabel(source, span).filter((tok) => tok.kind !== 'comment');
				const head = tokenText(toks[0]);
				const named = (tok: VbaToken | undefined): boolean => tokenName(tok)?.toLowerCase() === lower;
				if (head === 'redim') {
					const target = tokenText(toks[1]) === 'preserve' ? 2 : 1;
					if (named(toks[target])) {
						push('invalidParamArrayUse', `ParamArray '${toks[target].rawText}' holds what the call passes, and ReDim cannot size it. This is a VBE compile error: Invalid ParamArray use.`, at(span, toks[target]));
					}
					continue;
				}
				if (head === 'erase') {
					for (const group of splitTopLevelTokenGroups(toks, 1, ',', toks.length)) {
						if (group.length === 1 && named(group[0])) {
							push('invalidParamArrayUse', `ParamArray '${group[0].rawText}' holds what the call passes, and Erase cannot clear it. This is a VBE compile error: Invalid ParamArray use.`, at(span, group[0]));
						}
					}
					continue;
				}
				if (head === 'set' && named(toks[1]) && toks[2]?.rawText === '=') {
					push('arrayTargetAssignment', `ParamArray '${toks[1].rawText}' is an array, which Set cannot assign to. This is a VBE compile error: Can't assign to array.`, at(span, toks[1]));
					continue;
				}
				// Passed whole to a ByRef parameter of a procedure the module knows.
				for (let i = 0; i < toks.length; i++) {
					const callee = tokenName(toks[i])?.toLowerCase();
					const signature = callee && !own.has(callee) && toks[i - 1]?.rawText !== '.' ? signatures.get(callee) : undefined;
					if (!signature) {
						continue;
					}
					const parenthesized = toks[i + 1]?.rawText === '(';
					const statementCall = i === 0 && !parenthesized && toks.length > 1 && toks[1].rawText !== '=';
					if (!parenthesized && !statementCall) {
						continue;
					}
					const close = parenthesized ? matchParenFrom(toks, i + 1) : toks.length;
					const args = close < 0 ? [] : splitTopLevelTokenGroups(toks, parenthesized ? i + 2 : i + 1, ',', close);
					args.forEach((arg, k) => {
						let value = arg.filter((tok) => tok.kind !== 'comment');
						while (value.length > 2 && value[0].rawText === '(' && matchParenFrom(value, 0) === value.length - 1) {
							value = value.slice(1, -1);
						}
						const target = signature.params[k];
						if (value.length === 1 && named(value[0]) && target && !target.paramArray && target.byRef !== false) {
							push('invalidParamArrayUse', `ParamArray '${value[0].rawText}' is passed whole to '${target.name}' of '${signature.name}', which takes it ByRef. This is a VBE compile error: Invalid ParamArray use.`, at(span, value[0]));
						}
					});
				}
			}
		};
	};
}
