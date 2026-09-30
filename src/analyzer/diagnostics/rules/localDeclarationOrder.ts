// Rule: a procedure's own Dim, Static or Const covers only the lines after it.
// Measured by full compile in 64-bit Excel 16.0 (build 20326, 2026-09-30):
//
//  - Under Option Explicit, a name used above its local declaration, with no
//    module or project declaration of that name, is "Variable not defined":
//    `Debug.Print K` then `Const K = 1`, `x = 1` then `Dim x`, `Const A = B
//    + 1` then `Const B = 1`, `Dim a(N)` then `Const N = 5`, a use inside an
//    earlier loop. VBA does not hoist a declaration.
//  - Where the earlier use found something else, a module-level declaration
//    of that name or, without Option Explicit, an implicit variable, the
//    local declaration is "Duplicate declaration in current scope". Without
//    Option Explicit, an earlier use inside a Const's value is "Constant
//    expression required" instead, since an implicit variable is no constant.
//  - Scope is the whole procedure: a Const inside an If block, used after
//    the block, compiles.

import type { ConditionalActivityTracker } from '../../conditional/conditionalCompilation';
import type { HostObjectModel } from '../../host/excelObjectModel';
import { resolveHostGlobal } from '../../host/hostModel';
import { tokenizeCached } from '../../lexer/tokenize';
import type { VbaToken } from '../../lexer/tokenKinds';
import type { ModuleNode, Span } from '../../parser/nodes';
import { resolveRuntimeFunction } from '../../runtime/vbaRuntime';
import type { buildModuleSymbols } from '../../symbols/buildModuleSymbols';
import type { VbaSymbol } from '../../symbols/symbolModel';
import type { PushFn } from '../analysisContext';
import { sourceIdentifierBinding } from '../typeInference';
import { activeModuleMembers, forEachVariableGroup, isInactiveNode, tokenName, tokenText } from '../walker';

interface LocalDeclaration {
	name: string;
	nameSpan: Span;
	/** Where the declaring statement begins; a use before it is too early. */
	start: number;
	isConst: boolean;
}

/** Words after which a name is a label, a type or an object, never a value read. */
const NOT_A_VALUE_AFTER = new Set(['as', 'new', 'goto', 'gosub', 'resume']);

export function checkLocalDeclarationOrder(
	source: string,
	mod: ModuleNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
	projectVisibleSymbols: readonly VbaSymbol[] | undefined,
	hostModel: HostObjectModel | undefined,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	const optionExplicit = activeModuleMembers(mod, activity).some(
		(member) => member.kind === 'Option' && /^explicit\b/i.test(member.optionText.trim()),
	);
	let tokens: readonly VbaToken[] | undefined;
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		const declarations = new Map<string, LocalDeclaration>();
		forEachVariableGroup(member.body, (group) => {
			for (const decl of group.declarations) {
				const key = decl.name.toLowerCase();
				if (!isInactiveNode(activity, decl) && !declarations.has(key)) {
					declarations.set(key, {
						name: decl.name,
						nameSpan: decl.nameSpan ?? decl.span,
						start: group.span.start,
						isConst: group.isConst,
					});
				}
			}
		}, activity);
		if (declarations.size === 0) {
			continue;
		}
		const params = new Set(member.params.map((param) => param.name.toLowerCase()));
		const headerEnd = source.indexOf('\n', member.span.start);
		const bodyStart = headerEnd < 0 ? member.span.end : headerEnd;
		const lastDeclaration = Math.max(...[...declarations.values()].map((decl) => decl.start));
		tokens ??= tokenizeCached(source);
		const reported = new Set<string>();
		for (let i = firstTokenAt(tokens, bodyStart); i < tokens.length && tokens[i].start < lastDeclaration; i++) {
			const tok = tokens[i];
			const key = tok.kind === 'identifier' ? tokenName(tok)?.toLowerCase() : undefined;
			const declaration = key ? declarations.get(key) : undefined;
			if (!key || !declaration || tok.start >= declaration.start || params.has(key) || reported.has(key)) {
				continue;
			}
			if (!isValueUse(tokens, i) || activity?.isInactive({ start: tok.start, end: tok.end })) {
				continue;
			}
			// A name VBA or the host already gives a meaning is not measured here.
			if (resolveRuntimeFunction(declaration.name) || (hostModel && resolveHostGlobal(declaration.name, hostModel))) {
				continue;
			}
			reported.add(key);
			const outer = sourceIdentifierBinding(symbols, undefined, projectVisibleSymbols, declaration.name, 'expression');
			const inConstValue = [...declarations.values()].some(
				(other) => other.isConst && other.start <= tok.start && constStatementEnd(source, other.start) > tok.start,
			);
			if (optionExplicit && outer.scope === 'unresolved') {
				push(
					'undeclaredVariable',
					`Variable not defined: '${tok.rawText}'. It is declared further down the procedure, and a declaration covers only the lines after it; move the ${declaration.isConst ? 'Const' : 'declaration'} above this line.`,
					{ start: tok.start, end: tok.end },
				);
			} else if (!optionExplicit && outer.scope === 'unresolved' && inConstValue) {
				push(
					'constValueNotConstant',
					`'${tok.rawText}' is declared further down the procedure, so here it is an implicit variable, which a Const cannot take its value from. This is a VBE compile error: Constant expression required.`,
					{ start: tok.start, end: tok.end },
				);
			} else {
				push(
					'duplicateDeclaration',
					`'${declaration.name}' is used above this declaration, where it ${outer.scope === 'unresolved' ? 'became an implicit variable' : 'named the module\'s declaration'}. This is a VBE compile error: Duplicate declaration in current scope.`,
					declaration.nameSpan,
				);
			}
		}
	}
}

function firstTokenAt(tokens: readonly VbaToken[], offset: number): number {
	let lo = 0;
	let hi = tokens.length;
	while (lo < hi) {
		const mid = (lo + hi) >> 1;
		if (tokens[mid].start < offset) {
			lo = mid + 1;
		} else {
			hi = mid;
		}
	}
	return lo;
}

/** Whether the name at `i` reads a variable, rather than naming a member, a label, a type or an argument. */
function isValueUse(tokens: readonly VbaToken[], i: number): boolean {
	const before = tokens[i - 1];
	const after = tokens[i + 1];
	if (before && (before.rawText === '.' || before.rawText === '!' || NOT_A_VALUE_AFTER.has(tokenText(before)))) {
		return false;
	}
	if (after?.rawText === ':=') {
		return false;
	}
	// `Label:` at the start of a line.
	const atLineStart = !before || before.kind === 'newline';
	return !(atLineStart && after?.kind === 'colon');
}

/** Where the Const statement starting at `start` ends: its line's end. */
function constStatementEnd(source: string, start: number): number {
	const end = source.slice(start).search(/\r|\n/);
	return end < 0 ? source.length : start + end;
}
