// Rule family: statement forms the VBE refuses while compiling (issue #125).
// Measured in Excel 16.0 (build 20326, 2026-09-25):
//
//  - collection-operand: `x = c + 1` with c As New Collection -> "Argument
//    not optional". A Collection's default member Item takes an index, so
//    the bare variable has no value for the operator.
//  - sub-used-as-value: `x = Foo` where Foo is a Sub -> "Expected Function
//    or variable".
//  - rem-after-then: `If x Then Rem note` -> "Syntax error". Rem starts a
//    comment only at the start of a statement.

import type { ConditionalActivityTracker } from '../../conditional/conditionalCompilation';
import type { ModuleNode } from '../../parser/nodes';
import { buildModuleSymbols } from '../../symbols/buildModuleSymbols';
import type { VbaProcedureSignature } from '../../symbols/symbolModel';
import { procedureSymbolFor, type PushFn } from '../analysisContext';
import { normalizeType, typeEnvironmentFor } from '../typeInference';
import {
	activeModuleMembers,
	bareAssignmentTarget,
	forEachStatement,
	setAssignmentTarget,
	statementAndBranchSpans,
	statementTokens,
	tokenName,
	tokenText,
} from '../walker';

const SCALAR_OPERATORS: ReadonlySet<string> = new Set(['=', '<', '>', '<=', '>=', '<>', '+', '-', '*', '/', '\\', '&', '^']);

export function checkStatementForms(
	source: string,
	mod: ModuleNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
	projectProcedures: ReadonlyMap<string, readonly VbaProcedureSignature[]> | undefined,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	// Subs of this module, and of the project's standard modules, by name;
	// a name that is also a Function or a module-level variable anywhere is
	// not judged.
	const subs = new Set<string>();
	const notSubs = new Set<string>();
	for (const symbol of symbols.root.children ?? []) {
		const lower = symbol.name.toLowerCase();
		if (symbol.kind === 'sub') {
			subs.add(lower);
		} else {
			notSubs.add(lower);
		}
	}
	for (const [lower, signatures] of projectProcedures ?? []) {
		for (const signature of signatures) {
			(signature.kind === 'sub' ? subs : notSubs).add(lower.toLowerCase());
		}
	}
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		const env = typeEnvironmentFor(symbols, member);
		const locals = new Set<string>();
		for (const child of procedureSymbolFor(symbols, member)?.children ?? []) {
			locals.add(child.name.toLowerCase());
		}
		forEachStatement(member.body, (stmt) => {
			for (const span of statementAndBranchSpans(stmt)) {
				const toks = statementTokens(source, span);
				const at = (i: number) => ({ start: span.start + toks[i].start, end: span.start + toks[i].end });
				if (tokenText(toks[0]) === 'if') {
					const then = toks.findIndex((tok) => tokenText(tok) === 'then');
					if (then > 0 && tokenText(toks[then + 1]) === 'rem') {
						push('remAfterThen', "'Rem' cannot follow 'Then' on one line: a Rem comment starts only at the start of a statement. This is a VBE compile error: Syntax error.", at(then + 1));
					}
				}
				const target = bareAssignmentTarget(source, span);
				// A Set's `=` is the assignment too: `Set c = New Collection` is no operand.
				const assigns = target !== undefined || setAssignmentTarget(source, span) !== undefined;
				const eq = assigns ? toks.findIndex((tok) => tok.rawText === '=') : -1;
				for (let i = 0; i < toks.length; i++) {
					const name = tokenName(toks[i]);
					if (!name || toks[i - 1]?.rawText === '.' || toks[i + 1]?.rawText === ':=' || i === eq - 1) {
						continue;
					}
					const lower = name.toLowerCase();
					if (normalizeType(env.get(lower)) === 'collection' && toks[i + 1]?.rawText !== '(' && toks[i + 1]?.rawText !== '.') {
						const previous = i - 1 === eq ? undefined : toks[i - 1];
						const operator = [toks[i + 1], previous].find((tok) => tok && ((tok.kind === 'operator' && SCALAR_OPERATORS.has(tok.rawText)) || tokenText(tok) === 'mod'));
						if (operator) {
							push('collectionOperand', `'${name}' is a Collection: its default member Item needs an index, so '${operator.rawText}' has no value to work on. This is a VBE compile error: Argument not optional.`, at(i));
							continue;
						}
					}
					if (target && i > eq && subs.has(lower) && !notSubs.has(lower) && !locals.has(lower) && !env.has(lower)) {
						push('subUsedAsValue', `'${name}' is a Sub, which returns nothing, so it cannot be used as a value. This is a VBE compile error: Expected Function or variable.`, at(i));
					}
				}
			}
		}, activity);
	}
}
