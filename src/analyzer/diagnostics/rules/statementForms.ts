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
import { isKnownScalarType, normalizeType, objectValueNeedsIndex, typeEnvironmentFor } from '../typeInference';
import type { MemberCompletionContext } from '../../completion/memberAccess';
import {
	activeModuleMembers,
	bareAssignmentTarget,
	firstExecutableTokenIndex,
	forEachStatement,
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
	memberCtx: MemberCompletionContext = {},
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
		// An object whose default member needs an index: a Collection, or
		// Excel's Hyperlinks, Areas, Borders, Windows, Workbooks and Shapes
		// (issue #221, measured in Excel 16.0).
		const needsIndex = new Map<string, boolean>();
		const indexed = (lower: string): boolean => {
			let answer = needsIndex.get(lower);
			if (answer === undefined) {
				const type = env.get(lower);
				answer = type !== undefined && objectValueNeedsIndex(type, memberCtx);
				needsIndex.set(lower, answer);
			}
			return answer;
		};
		const typedValue = (lower: string): boolean => {
			const type = normalizeType(lower === member.name.toLowerCase() ? member.returnType : env.get(lower));
			return type !== undefined && isKnownScalarType(type);
		};
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
				// A Set's `=` is the assignment too: `Set c = New Collection` is no
				// operand, and neither is `Set cols(1) = c`, whose target is
				// indexed (issue #140).
				const first = firstExecutableTokenIndex(toks);
				const assigns = target !== undefined || tokenText(toks[first]) === 'set';
				const eq = assigns ? toks.findIndex((tok) => tok.rawText === '=') : -1;
				// A one-line If is judged as its condition here; each branch is its
				// own span with its own assignment (issue #140: `If c Is Nothing
				// Then Set c = New Collection`).
				const then = tokenText(toks[first]) === 'if' && stmt.kind === 'Statement' && stmt.singleLineIfBranches
					? toks.findIndex((tok) => tokenText(tok) === 'then')
					: -1;
				const limit = then > 0 ? then : toks.length;
				for (let i = 0; i < limit; i++) {
					const name = tokenName(toks[i]);
					if (!name || toks[i - 1]?.rawText === '.' || toks[i + 1]?.rawText === ':=' || i === eq - 1) {
						continue;
					}
					// `AddressOf TimerProc` takes the procedure's address, not its value.
					if (tokenText(toks[i - 1]) === 'addressof') {
						continue;
					}
					const lower = name.toLowerCase();
					if (toks[i + 1]?.rawText !== '(' && toks[i + 1]?.rawText !== '.' && indexed(lower)) {
						const typeName = env.get(lower)!;
						const previous = i - 1 === eq ? undefined : toks[i - 1];
						const operator = [toks[i + 1], previous].find((tok) => tok && ((tok.kind === 'operator' && SCALAR_OPERATORS.has(tok.rawText)) || tokenText(tok) === 'mod'));
						if (operator) {
							push('collectionOperand', `'${name}' is ${/^[aeiou]/i.test(typeName) ? 'an' : 'a'} ${typeName}: its default member Item needs an index, so '${operator.rawText}' has no value to work on. This is a VBE compile error: Argument not optional.`, at(i));
							continue;
						}
						// `s = c` with s a String: the whole value of a Let into a
						// typed value (issue #221).
						if (target && i === eq + 1 && toks.length === eq + 2 && typedValue(target.name.toLowerCase())) {
							push('collectionOperand', `'${name}' is ${/^[aeiou]/i.test(typeName) ? 'an' : 'a'} ${typeName}: its default member Item needs an index, so it has no value for '${target.name}' to take. This is a VBE compile error: Argument not optional.`, at(i));
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
