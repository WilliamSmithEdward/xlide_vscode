// Rule family: module-kind constraints (audit #0).
//
// Extracted verbatim from analyzeModule.ts: rules keyed to the containing
// module kind - object-module Public restrictions, Event/WithEvents/Friend/
// Implements placement, RaiseEvent targets, PtrSafe, and event-handler scope.

import {
	type EventHandlerDocumentType,
	eventHandlerDocumentTypeForContext,
	eventHandlerProcedureForName,
} from '../../completion/eventHandlers';
import {
	type ConditionalActivityTracker,
	type ConditionalCompilationEnvironment,
	conditionalCompilerConstants,
} from '../../conditional/conditionalCompilation';
import type { VbaToken } from '../../lexer/tokenKinds';
import type {
	ModuleNode,
	Span,
	VariableGroupNode,
} from '../../parser/nodes';
import type { ModuleSymbolKind } from '../../symbols/symbolModel';
import {
	isObjectModuleKind,
	type PushFn,
} from '../analysisContext';
import {
	absoluteSpan,
	activeModuleMembers,
	declaredNameSpan,
	firstTokenSpan,
	forEachProcedureBodyLine,
	forEachStatement,
	forEachVariableGroup,
	isInactiveNode,
	statementTokens,
	statementTokensAfterLeadingLabel,
	tokenName,
	tokenText,
	type ProcedureStatementVisitor,
} from '../walker';

/**
 * Rule: object modules cannot expose certain public declarations as object
 * members. VBE reports one compile error family for public constants,
 * fixed-length strings, arrays, user-defined types, and Declare statements in
 * class/document/UserForm modules.
 */
export function checkObjectModulePublicMembers(
	source: string,
	mod: ModuleNode,
	moduleKind: ModuleSymbolKind,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	if (!isObjectModuleKind(moduleKind)) {
		return;
	}

	const report = (kind: string, span: Span): void => {
		push(
			'objectModulePublicMember',
			`Public ${kind} are not allowed as Public members of object modules; VBE Compile rejects this declaration.`,
			span,
		);
	};

	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind === 'VariableGroup' && isPublicModifier(member.modifier)) {
			for (const decl of member.declarations) {
				const span = declaredNameSpan(source, decl.span, decl.name);
				if (member.isConst) {
					report('constants', span);
				} else if (decl.isArray) {
					report('arrays', span);
				} else if (decl.fixedLength !== undefined) {
					report('fixed-length strings', span);
				}
			}
			continue;
		}

		if (member.kind === 'Type' && isPublicModifier(member.visibility)) {
			report('user-defined types', declaredNameSpan(source, member.span, member.name));
			continue;
		}

		if (member.kind === 'Declare' && isPublicModifier(member.visibility)) {
			report('Declare statements', declaredNameSpan(source, member.span, member.name));
		}
	}
}

/**
 * Rule: `Event` declarations are object-module declarations. Standard modules
 * can contain ordinary procedures that look like handlers, but not `Event`
 * declarations themselves.
 */
export function checkEventDeclarationModuleKind(
	source: string,
	mod: ModuleNode,
	moduleKind: ModuleSymbolKind,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	if (isObjectModuleKind(moduleKind)) {
		return;
	}
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Event') {
			continue;
		}
		push(
			'eventDeclarationModuleKind',
			`Event declaration '${member.name}' is only valid in class, document, or UserForm modules.`,
			declaredNameSpan(source, member.span, member.name),
		);
	}
}

/**
 * Rule: the `Me` keyword refers to the current object instance and is only valid
 * in class, document, or UserForm modules. In a standard module any use of `Me`
 * is a compile error ("Invalid use of Me keyword"). Oracle-verified
 * (`corpus_me_004_compile`). A `Me` that follows `.` is a member name, not the
 * keyword, and is left alone (no false positive).
 */
export function checkMeOutsideObjectModule(
	moduleKind: ModuleSymbolKind,
	source: string,
	push: PushFn,
): ProcedureStatementVisitor {
	if (isObjectModuleKind(moduleKind)) {
		// `Me` is valid in an object module - skip every procedure (no per-statement cost).
		return () => undefined;
	}
	return () => (stmt) => {
		const toks = statementTokens(source, stmt.span);
		for (let i = 0; i < toks.length; i++) {
			if (tokenText(toks[i]) !== 'me') {
				continue;
			}
			if (i > 0 && toks[i - 1].rawText === '.') {
				continue; // a member named Me, not the Me keyword
			}
			push(
				'meOutsideObjectModule',
				"'Me' is only valid in a class, document, or UserForm module.",
				absoluteSpan(stmt.span, toks[i]),
			);
		}
	};
}

/**
 * Rule: `WithEvents` is a module-level object-module variable declarator. The
 * parser exposes the relevant settled facts directly: module kind, local vs
 * module declaration context, `As New`, and array declarators.
 */
export function checkWithEventsDeclarations(
	source: string,
	mod: ModuleNode,
	moduleKind: ModuleSymbolKind,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	const report = (message: string, span: Span): void => {
		push('withEventsDeclaration', message, span);
	};
	const inspect = (group: VariableGroupNode, insideProcedure: boolean): void => {
		if (!group.withEvents || isInactiveNode(activity, group)) {
			return;
		}
		for (const decl of group.declarations) {
			const nameSpan = declaredNameSpan(source, decl.span, decl.name);
			if (insideProcedure) {
				report(
					`WithEvents variable '${decl.name}' must be declared at module level.`,
					nameSpan,
				);
				continue;
			}
			if (!isObjectModuleKind(moduleKind)) {
				report(
					`WithEvents variable '${decl.name}' is only valid in class, document, or UserForm modules.`,
					nameSpan,
				);
				continue;
			}
			if (decl.isNew) {
				report(
					`WithEvents variable '${decl.name}' cannot be declared As New.`,
					nameSpan,
				);
			}
			if (decl.isArray) {
				report(
					`WithEvents variable '${decl.name}' cannot be an array.`,
					nameSpan,
				);
			}
		}
	};

	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind === 'VariableGroup') {
			inspect(member, false);
			continue;
		}
		if (member.kind === 'Procedure') {
			forEachVariableGroup(member.body, (group) => inspect(group, true), activity);
		}
	}
}

/**
 * Rule: `Friend` is procedure visibility for object modules. It is not a
 * module-variable modifier and it is not valid for standard-module procedures.
 */
export function checkFriendDeclarations(
	source: string,
	mod: ModuleNode,
	moduleKind: ModuleSymbolKind,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind === 'Procedure') {
			if (hasFriendModifier(member.modifiers) && !isObjectModuleKind(moduleKind)) {
				push(
					'friendDeclaration',
					`Friend procedure '${member.name}' is only valid in class, document, or UserForm modules.`,
					friendKeywordSpan(source, member.span),
				);
			}
			continue;
		}
		if (member.kind !== 'VariableGroup' || member.modifier.toLowerCase() !== 'friend') {
			continue;
		}
		push(
			'friendDeclaration',
			'Friend can only modify procedure declarations, not variables.',
			friendKeywordSpan(source, member.span),
		);
	}
}

function hasFriendModifier(modifiers: readonly string[]): boolean {
	return modifiers.some((modifier) => modifier.toLowerCase() === 'friend');
}

function friendKeywordSpan(source: string, span: Span): Span {
	const tok = statementTokensAfterLeadingLabel(source, span)
		.find((token) => tokenText(token) === 'friend');
	return tok ? absoluteSpan(span, tok) : firstTokenSpan(source, span);
}

/**
 * Rule: `Implements` is an object-module declaration-section statement. This
 * intentionally validates only placement/module-kind facts, leaving interface
 * member completeness to a later project-binder pass.
 */
export function checkImplementsStatementPlacement(
	source: string,
	mod: ModuleNode,
	moduleKind: ModuleSymbolKind,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	// Procedures above the Implements under test that could be compiled beside
	// it. One in the other arm of a `#If` chain never is (issues/58).
	const proceduresAbove: Span[] = [];
	const reportModuleKind = (hit: ImplementsStatementHit): void => {
		push(
			'implementsStatementPlacement',
			`Implements statement '${hit.name}' is only valid in class, document, or UserForm modules.`,
			hit.span,
		);
	};
	const reportProcedurePlacement = (hit: ImplementsStatementHit): void => {
		push(
			'implementsStatementPlacement',
			`Implements statement '${hit.name}' must appear in the module declaration section before any procedure.`,
			hit.span,
		);
	};

	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind === 'Procedure') {
			proceduresAbove.push(member.span);
			forEachStatement(member.body, (stmt) => {
				const hit = implementsStatementHit(source, stmt.span);
				if (hit) {
					reportProcedurePlacement(hit);
				}
			}, activity);
			continue;
		}
		if (member.kind !== 'Statement') {
			continue;
		}
		const hit = implementsStatementHit(source, member.span);
		if (!hit) {
			continue;
		}
		if (!isObjectModuleKind(moduleKind)) {
			reportModuleKind(hit);
			continue;
		}
		const compiledTogether = proceduresAbove.some(
			(prior) => !activity?.mutuallyExclusive(prior, member.span),
		);
		if (compiledTogether) {
			reportProcedurePlacement(hit);
		}
	}
}

interface ImplementsStatementHit {
	name: string;
	span: Span;
}

function implementsStatementHit(source: string, span: Span): ImplementsStatementHit | undefined {
	const toks = statementTokensAfterLeadingLabel(source, span);
	if (tokenText(toks[0]) !== 'implements') {
		return undefined;
	}
	const firstName = tokenName(toks[1]);
	if (!firstName) {
		return undefined;
	}
	let name = firstName;
	let endIndex = 1;
	while (toks[endIndex + 1]?.rawText === '.') {
		const part = tokenName(toks[endIndex + 2]);
		if (!part) {
			break;
		}
		name += `.${part}`;
		endIndex += 2;
	}
	return {
		name,
		span: {
			start: span.start + toks[1].start,
			end: span.start + toks[endIndex].end,
		},
	};
}

/**
 * Rule: `RaiseEvent` names an Event declared by the containing module. Event
 * signature/arity checks remain deferred to the richer event-binding slice.
 */
export function checkRaiseEventTargets(
	source: string,
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	const events = new Set<string>();
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind === 'Event' && member.name) {
			events.add(member.name.toLowerCase());
		}
	}

	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		forEachProcedureBodyLine(source, member, (lineSpan) => {
			if (activity?.isInactive(lineSpan)) {
				return;
			}
			// A single physical line can carry several `:`-separated statements
			// (e.g. `RaiseEvent A: RaiseEvent B`), so check every RaiseEvent on it.
			for (const hit of raiseEventTargetHits(source, lineSpan)) {
				if (events.has(hit.name.toLowerCase())) {
					continue;
				}
				push(
					'raiseEventUndeclaredEvent',
					`Event '${hit.name}' is not declared in this module, so it cannot be raised with RaiseEvent.`,
					hit.span,
				);
			}
		});
	}
}

function raiseEventTargetHits(
	source: string,
	span: Span,
): Array<{ name: string; span: Span }> {
	const toks = statementTokens(source, span);
	const hits: Array<{ name: string; span: Span }> = [];
	// Each `:`-separated statement segment on the line may be its own RaiseEvent.
	for (const segmentStart of statementSegmentStarts(toks)) {
		if (tokenText(toks[segmentStart]) !== 'raiseevent') {
			continue;
		}
		const nameTok = toks[segmentStart + 1];
		const name = nameTok ? tokenName(nameTok) : undefined;
		if (!name) {
			continue;
		}
		hits.push({
			name,
			span: {
				start: span.start + nameTok.start,
				end: span.start + nameTok.end,
			},
		});
	}
	return hits;
}

/**
 * Indices where each `:`-separated statement segment begins on a logical line.
 * The first segment skips a leading line-number or `Label:` prefix; subsequent
 * segments begin right after each top-level statement-separator colon. A colon
 * inside parentheses/brackets is not a separator.
 */
function statementSegmentStarts(toks: readonly VbaToken[]): number[] {
	if (toks.length === 0) {
		return [];
	}
	const starts: number[] = [];
	let depth = 0;
	let segmentStart = 0;
	// Skip a leading line-number or `Label:` on the first segment.
	if (toks.length > 1 && /^\d+$/.test(toks[0].rawText)) {
		segmentStart = 1;
	} else if (
		toks.length > 2 &&
		(toks[0].kind === 'identifier' || toks[0].kind === 'keyword') &&
		toks[1].rawText === ':'
	) {
		segmentStart = 2;
	}
	starts.push(segmentStart);
	for (let i = segmentStart; i < toks.length; i++) {
		const raw = toks[i].rawText;
		if (raw === '(' || raw === '[') {
			depth++;
		} else if (raw === ')' || raw === ']') {
			depth--;
		} else if (raw === ':' && depth === 0 && i + 1 < toks.length) {
			starts.push(i + 1);
		}
	}
	return starts;
}

export function checkDeclarePtrSafeForWin64(
	source: string,
	mod: ModuleNode,
	conditionalCompilation: ConditionalCompilationEnvironment | undefined,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	if (!conditionalValueTruthy(conditionalCompilerConstants(conditionalCompilation).get('win64'))) {
		return;
	}
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Declare' || member.ptrSafe) {
			continue;
		}
		push(
			'declareMissingPtrSafe',
			`Declare statement '${member.name}' must include PtrSafe when compiling for 64-bit Office.`,
			declaredNameSpan(source, member.span, member.name),
		);
	}
}

function conditionalValueTruthy(value: unknown): boolean {
	if (typeof value === 'boolean') {
		return value;
	}
	if (typeof value === 'number') {
		return value !== 0;
	}
	return typeof value === 'string' && value.length > 0;
}

export function checkEventHandlerModuleScope(
	source: string,
	mod: ModuleNode,
	moduleName: string,
	moduleKind: ModuleSymbolKind,
	documentType: EventHandlerDocumentType | undefined,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	const actualDocumentType = eventHandlerDocumentTypeForContext({
		moduleName,
		moduleKind,
		documentType,
	});
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure' || member.procKind !== 'Sub') {
			continue;
		}
		const event = eventHandlerProcedureForName(member.name);
		if (!event) {
			continue;
		}
		if (actualDocumentType === event.documentType) {
			continue;
		}
		const moduleDescription =
			moduleKind === 'document'
				? `${describeEventDocumentType(actualDocumentType)} document module`
				: `${moduleKind} module`;
		push(
			'eventHandlerWrongModule',
			`'${event.name}' matches a ${event.owner} event handler, but this ${moduleDescription} is not where Excel wires that event. It will behave like an ordinary procedure here.`,
			declaredNameSpan(source, member.span, member.name),
		);
	}
}

function describeEventDocumentType(
	documentType: EventHandlerDocumentType | undefined,
): string {
	switch (documentType) {
		case 'workbook':
			return 'workbook';
		case 'worksheet':
			return 'worksheet';
		case 'chart':
			return 'chart';
		default:
			return 'unknown';
	}
}

function isPublicModifier(value: string | undefined): boolean {
	return value?.toLowerCase() === 'public';
}
