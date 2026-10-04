// Rule: a ParamArray used where the VBE refuses it (issue #445). Each is a
// compile error, measured in Excel 16.0 (build 20430, 2026-10-02), inside
// `Function F(ParamArray p() As Variant)`:
//
//  - `ReDim p(3)` and `Erase p`: "Invalid ParamArray use".
//  - `G(p)`, `G((p))`, `Call S(p)` and `S p` with G's or S's parameter ByRef,
//    a Variant or an array: "Invalid ParamArray use". A ByVal parameter and
//    another ParamArray take it. So do a named argument, `G(v:=p)`,
//    `Module2.G(p)` and `k.G(p)` of a project class by their parameter, and
//    any member of a late-bound object, `o.G(p)`, whose parameter the
//    compiler cannot see (issue #685).
//  - `Set p = Nothing`: "Can't assign to array".
//  - `Function F(ParamArray p)`, declared with no parentheses: refused.
//
// `p = Array(1)`, `p(0) = 9`, `UBound(p)` and `For Each` over p compile.

import type { VbaToken } from '../../lexer/tokenKinds';
import type { ProcedureNode, Span } from '../../parser/nodes';
import type { PushFn } from '../analysisContext';
import type { CallableTypeSignature } from '../callExtraction';
import type { MemberCompletionContext } from '../../completion/memberAccess';
import type { buildModuleSymbols } from '../../symbols/buildModuleSymbols';
import type { VbaProcedureParam, VbaProjectClassMember } from '../../symbols/symbolModel';
import { normalizeType, typeEnvironmentFor } from '../typeInference';
import { matchParenFrom, statementAndBranchSpans, statementTokensAfterLeadingLabel, tokenName, tokenText, type ProcedureStatementVisitor } from '../walker';
import { splitTopLevelTokenGroups } from '../../lexer/tokenHelpers';

/** Whether the parameter at `index`, or named `named`, takes its argument ByRef: undefined where there is none. */
function takesByRef(params: readonly VbaProcedureParam[], index: number, named?: string): boolean | undefined {
	const param = named !== undefined ? params.find((p) => p.name.toLowerCase() === named.toLowerCase()) : params[index];
	return param ? !param.byVal && !param.paramArray : undefined;
}

export function checkParamArrayUse(
	source: string,
	signatures: ReadonlyMap<string, CallableTypeSignature>,
	push: PushFn,
	symbols?: ReturnType<typeof buildModuleSymbols>,
	memberCtx?: MemberCompletionContext,
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
		// What `receiver.callee` takes: 'late' for an Object or Variant, the
		// member's parameters for a project class or another module, else unknown.
		const env = symbols ? typeEnvironmentFor(symbols, member) : new Map<string, string>();
		const surfaces = memberCtx?.projectClassMembers ?? [];
		const paramsOf = (found: VbaProjectClassMember | undefined): readonly VbaProcedureParam[] | undefined =>
			found?.procedureParams?.function ?? found?.procedureParams?.sub ?? found?.procedureParams?.propertyGet;
		const receiverSignature = (receiver: string, callee: string): readonly VbaProcedureParam[] | 'late' | undefined => {
			if (env.has(receiver)) {
				const type = normalizeType(env.get(receiver)!.replace(/^new\s+/i, ''));
				if (type === undefined || type === 'object' || type === 'variant') {
					return 'late';
				}
				return paramsOf(surfaces.find((surface) => surface.kind === 'class' && surface.name.toLowerCase() === type)
					?.members.find((m) => m.name.toLowerCase() === callee));
			}
			return paramsOf(surfaces.find((surface) => surface.kind === 'standardModule' && surface.name.toLowerCase() === receiver)
				?.members.find((m) => m.name.toLowerCase() === callee));
		};
		return (stmt) => {
			for (const span of statementAndBranchSpans(stmt)) {
				const toks = statementTokensAfterLeadingLabel(source, span);
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
				// Passed whole to a ByRef parameter of a procedure the module knows,
				// by position or by name, or to a member of an object (issue #685,
				// measured in Excel 16.0): a late-bound one always, since the
				// compiler cannot tell it ByVal; a project class's or another
				// module's by its parameter.
				for (let i = 0; i < toks.length; i++) {
					const callee = tokenName(toks[i])?.toLowerCase();
					const qualified = toks[i - 1]?.rawText === '.';
					const receiver = qualified && toks[i - 3]?.rawText !== '.' ? tokenName(toks[i - 2])?.toLowerCase() : undefined;
					const signature = callee && !own.has(callee) && !qualified ? signatures.get(callee) : undefined;
					const via = callee && receiver ? receiverSignature(receiver, callee) : undefined;
					if (!signature && !via) {
						continue;
					}
					const parenthesized = toks[i + 1]?.rawText === '(';
					const statementCall = (i === 0 || (qualified && i === 2)) && !parenthesized && toks.length > i + 1 && toks[i + 1].rawText !== '=';
					if (!parenthesized && !statementCall) {
						continue;
					}
					const close = parenthesized ? matchParenFrom(toks, i + 1) : toks.length;
					const args = close < 0 ? [] : splitTopLevelTokenGroups(toks, parenthesized ? i + 2 : i + 1, ',', close);
					args.forEach((arg, k) => {
						let value = arg.filter((tok) => tok.kind !== 'comment');
						const nameAt = value.length > 2 && value[1].rawText === ':=' ? tokenName(value[0]) : undefined;
						if (nameAt !== undefined) {
							value = value.slice(2);
						}
						while (value.length > 2 && value[0].rawText === '(' && matchParenFrom(value, 0) === value.length - 1) {
							value = value.slice(1, -1);
						}
						if (value.length !== 1 || !named(value[0])) {
							return;
						}
						if (signature) {
							const target = nameAt !== undefined ? signature.params.find((p) => p.name.toLowerCase() === nameAt.toLowerCase()) : signature.params[k];
							if (target && !target.paramArray && target.byRef !== false) {
								push('invalidParamArrayUse', `ParamArray '${value[0].rawText}' is passed whole to '${target.name}' of '${signature.name}', which takes it ByRef. This is a VBE compile error: Invalid ParamArray use.`, at(span, value[0]));
							}
							return;
						}
						const byRef = via === 'late' ? true : takesByRef(via!, k, nameAt);
						if (byRef) {
							const what = via === 'late' ? `a member of late-bound '${toks[i - 2].rawText}', which may take it ByRef` : `'${toks[i - 2].rawText}.${toks[i].rawText}', which takes it ByRef`;
							push('invalidParamArrayUse', `ParamArray '${value[0].rawText}' is passed whole to ${what}. This is a VBE compile error: Invalid ParamArray use.`, at(span, value[0]));
						}
					});
				}
			}
		};
	};
}
