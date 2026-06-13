// Rule family: control-flow rules (audit #0).
//
// Extracted verbatim from analyzeModule.ts: Exit statement kinds, GoTo label
// declarations/references, statement context (Case/Loop/Wend/Next placement),
// For Each control-variable and source types, and Else/ElseIf branch order.

import { resolveTypeName } from '../../completion/typeCompletion';
import type { ConditionalActivityTracker } from '../../conditional/conditionalCompilation';
import {
	collectProcedureLabelDeclarations,
	collectProcedureLabelReferences,
	collectProcedureLabels,
} from '../../flow/procedureLabels';
import { tokenize } from '../../lexer/tokenize';
import type { VbaToken } from '../../lexer/tokenKinds';
import {
	type BodyNode,
	type ForBlockNode,
	isLeafStatement,
	type ModuleNode,
	type ProcedureNode,
	type SelectBlockNode,
	type Span,
	type LeafStatementNode,
} from '../../parser/nodes';
import { buildModuleSymbols } from '../../symbols/buildModuleSymbols';
import type { BareIdentifierContext } from '../../symbols/nameResolution';
import {
	type AnalyzeModuleOptions,
	procedureSymbolFor,
	type PushFn,
} from '../analysisContext';
import { scanConditionalCompilationBranchOrder } from '../rules/shared';
import {
	declarationShapeEnvironmentFor,
	declaredShapeForSourceBinding,
	type DeclaredValueShape,
	isKnownScalarType,
	normalizeType,
	type SourceDeclaredShape,
} from '../typeInference';
import {
	absoluteSpan,
	activeModuleMembers,
	isInactiveNode,
	statementTokens,
	statementTokensAfterLeadingLabel,
	tokenName,
	tokenText,
	type ProcedureStatementVisitor,
} from '../walker';

/** Index of the `)` matching the `(` at `open`, or -1 if unbalanced. */
/**
 * Rule: an `Exit Sub` / `Exit Function` / `Exit Property` must match the kind of
 * the procedure that encloses it (the three Property accessors all map to
 * `Property`). `Exit Do` / `Exit For` are loop exits and are ignored here.
 *
 * Per-statement rule: rides the shared procedure-statement walk (audit #0).
 */
export function checkExitStatements(
	source: string,
	push: PushFn,
): ProcedureStatementVisitor {
	return (member) => {
		const expected = expectedExitWord(member.procKind);
		const label = enclosingProcLabel(member.procKind);
		return (stmt) => {
			const hit = exitTarget(source, stmt.span);
			if (hit && hit.word !== expected) {
				push(
					'exitWrongProcedure',
					`'Exit ${hit.word}' is not valid inside a ${label}; use 'Exit ${expected}'.`,
					hit.span,
				);
			}
		};
	};
}

/** Maps a procedure kind to the keyword its `Exit` statement must use. */
function expectedExitWord(kind: ProcedureNode['procKind']): string {
	if (kind === 'Sub') {
		return 'Sub';
	}
	if (kind === 'Function') {
		return 'Function';
	}
	return 'Property';
}

/** Human label for a procedure kind, for diagnostic messages. */
function enclosingProcLabel(kind: ProcedureNode['procKind']): string {
	if (kind === 'Sub') {
		return 'Sub';
	}
	if (kind === 'Function') {
		return 'Function';
	}
	return 'Property procedure';
}

/** If a statement is `Exit Sub|Function|Property`, returns the word and span. */
function exitTarget(
	source: string,
	span: Span,
): { word: string; span: Span } | undefined {
	const toks = statementTokensAfterLeadingLabel(source, span);
	if (toks.length < 2 || toks[0].rawText.toLowerCase() !== 'exit') {
		return undefined;
	}
	const w = toks[1].rawText.toLowerCase();
	let word: string;
	if (w === 'sub') {
		word = 'Sub';
	} else if (w === 'function') {
		word = 'Function';
	} else if (w === 'property') {
		word = 'Property';
	} else {
		return undefined; // Exit Do / Exit For etc.
	}
	return {
		word,
		span: { start: span.start + toks[0].start, end: span.start + toks[1].end },
	};
}

/**
 * Rule: procedure-local control-flow labels must exist in the same procedure.
 * This covers the deterministic VBE "Label not defined" cases without letting
 * labels leak across procedures or across inactive conditional-compilation code.
 */
export function checkUndefinedLabels(
	source: string,
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		const labels = collectProcedureLabels(source, member, activity);
		for (const ref of collectProcedureLabelReferences(source, member, activity)) {
			if (!labels.has(ref.key)) {
				push(
					'undefinedLabel',
					`Label '${ref.text}' is not defined in procedure '${member.name}'.`,
					ref.span,
				);
			}
		}
	}
}

/**
 * Rule: procedure-local labels must be unique within the same procedure. This
 * catches duplicate named labels and normalized decimal line labels.
 */
export function checkDuplicateLabels(
	source: string,
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		const seen = new Set<string>();
		for (const label of collectProcedureLabelDeclarations(source, member, activity)) {
			if (!seen.has(label.key)) {
				seen.add(label.key);
				continue;
			}
			push(
				'duplicateLabel',
				`Label '${label.text}' is already defined in procedure '${member.name}'.`,
				label.span,
			);
		}
	}
}

interface StatementContext {
	forDepth: number;
	doDepth: number;
	withDepth: number;
	selectDepth: number;
}

/**
 * Rules that depend on where a statement appears in the block tree:
 * `If` requires `Then`, `Case` belongs to `Select Case`, a leading `.member`
 * requires `With`, and loop exits require their matching loop.
 */
export function checkStatementContext(
	source: string,
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	const root: StatementContext = {
		forDepth: 0,
		doDepth: 0,
		withDepth: 0,
		selectDepth: 0,
	};

	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind === 'Statement') {
			continue;
		} else if (member.kind === 'Procedure') {
			checkContextBody(source, member.body, root, activity, push);
		}
	}
}

function checkContextBody(
	source: string,
	body: BodyNode[],
	ctx: StatementContext,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	for (const node of body) {
		if (isInactiveNode(activity, node)) {
			continue;
		}
		switch (node.kind) {
			case 'Statement':
			case 'Assignment':
			case 'Call':
				// Leaf statements: leading-dot (outside-With) and Exit checks
				// re-tokenize the span, so Assignment/Call route here too.
				checkContextStatement(source, node, ctx, push);
				break;
			case 'ForBlock':
				checkForNextControlVariable(source, node, activity, push);
				checkContextBody(
					source,
					node.body,
					{ ...ctx, forDepth: ctx.forDepth + 1 },
					activity,
					push,
				);
				break;
			case 'DoBlock':
				checkContextBody(
					source,
					node.body,
					{ ...ctx, doDepth: ctx.doDepth + 1 },
					activity,
					push,
				);
				break;
			case 'WithBlock':
				checkContextBody(
					source,
					node.body,
					{ ...ctx, withDepth: ctx.withDepth + 1 },
					activity,
					push,
				);
				break;
			case 'SelectBlock':
				checkContextBody(
					source,
					node.body,
					{ ...ctx, selectDepth: ctx.selectDepth + 1 },
					activity,
					push,
				);
				break;
			case 'IfBlock':
			case 'WhileBlock':
				checkContextBody(source, node.body, ctx, activity, push);
				break;
			case 'ConditionalDirective':
			case 'VariableGroup':
				break;
		}
	}
}

function checkForNextControlVariable(
	source: string,
	node: ForBlockNode,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	if (
		!node.controlVariable ||
		!node.controlVariableSpan ||
		!node.nextVariable ||
		!node.nextVariableSpan ||
		isInactiveNode(activity, { span: node.controlVariableSpan }) ||
		isInactiveNode(activity, { span: node.nextVariableSpan })
	) {
		return;
	}
	if (node.controlVariable.toLowerCase() === node.nextVariable.toLowerCase()) {
		return;
	}
	push(
		'nextVariableMismatch',
		`Next variable '${node.nextVariable}' does not match active For control variable '${node.controlVariable}'.`,
		node.nextVariableSpan,
	);
}

export function checkForEachLoopTypes(
	mod: ModuleNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
	opts: AnalyzeModuleOptions,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		const shapes = declarationShapeEnvironmentFor(symbols, member);
		const procSym = procedureSymbolFor(symbols, member);
		const resolveShape = (name: string, context: BareIdentifierContext): SourceDeclaredShape =>
			declaredShapeForSourceBinding(
				symbols,
				procSym,
				opts.projectVisibleSymbols,
				name,
				context,
			);
		checkForEachLoopTypesInBody(member.body, shapes, opts, activity, push, resolveShape);
	}
}

function checkForEachLoopTypesInBody(
	body: BodyNode[],
	shapes: ReadonlyMap<string, DeclaredValueShape>,
	opts: AnalyzeModuleOptions,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
	resolveShape?: (name: string, context: BareIdentifierContext) => SourceDeclaredShape,
): void {
	for (const node of body) {
		if (isInactiveNode(activity, node)) {
			continue;
		}
		if ('body' in node && Array.isArray(node.body)) {
			if (node.kind === 'ForBlock') {
				checkForEachControlVariableType(node, shapes, opts, activity, push, resolveShape);
				checkForEachSourceType(node, shapes, opts, activity, push, resolveShape);
			}
			checkForEachLoopTypesInBody(node.body, shapes, opts, activity, push, resolveShape);
		}
	}
}

function checkForEachControlVariableType(
	node: ForBlockNode,
	shapes: ReadonlyMap<string, DeclaredValueShape>,
	opts: AnalyzeModuleOptions,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
	resolveShape?: (name: string, context: BareIdentifierContext) => SourceDeclaredShape,
): void {
	if (
		!node.each ||
		!node.controlVariable ||
		!node.controlVariableSpan ||
		isInactiveNode(activity, { span: node.controlVariableSpan })
	) {
		return;
	}
	const resolvedShape = resolveShape?.(node.controlVariable, 'assignmentTarget');
	const shape = resolvedShape?.resolved
		? resolvedShape.shape
		: shapes.get(node.controlVariable.toLowerCase());
	if (!shape) {
		return;
	}
	const problem = forEachControlVariableTypeProblem(shape, opts);
	if (!problem) {
		return;
	}
	push(
		'forEachControlVariableType',
		`For Each control variable '${node.controlVariable}' must be Variant or Object, but ${problem}.`,
		node.controlVariableSpan,
	);
}

function checkForEachSourceType(
	node: ForBlockNode,
	shapes: ReadonlyMap<string, DeclaredValueShape>,
	opts: AnalyzeModuleOptions,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
	resolveShape?: (name: string, context: BareIdentifierContext) => SourceDeclaredShape,
): void {
	if (
		!node.each ||
		!node.sourceExpression ||
		!node.sourceExpressionSpan ||
		isInactiveNode(activity, { span: node.sourceExpressionSpan })
	) {
		return;
	}
	const sourceName = simpleForEachSourceName(node.sourceExpression);
	if (!sourceName) {
		return;
	}
	const resolvedShape = resolveShape?.(sourceName, 'expression');
	const shape = resolvedShape?.resolved
		? resolvedShape.shape
		: shapes.get(sourceName.toLowerCase());
	if (!shape) {
		return;
	}
	const problem = forEachSourceTypeProblem(shape, opts);
	if (!problem) {
		return;
	}
	push(
		'forEachSourceType',
		`For Each source '${sourceName}' must be a collection object or array, but ${problem}.`,
		node.sourceExpressionSpan,
	);
}

function forEachControlVariableTypeProblem(
	shape: DeclaredValueShape,
	opts: AnalyzeModuleOptions,
): string | undefined {
	if (shape.isArray) {
		return 'it is an array variable';
	}
	if (!shape.asType) {
		return undefined;
	}
	const resolved = resolveTypeName(shape.asType, {
		projectTypes: opts.projectTypes,
		model: opts.hostModel,
	});
	if (resolved?.kind === 'userType') {
		return `it is declared As user-defined Type '${shape.asType}'`;
	}
	if (resolved?.kind === 'enum') {
		return `it is declared As Enum '${shape.asType}'`;
	}
	if (resolved && resolved.kind !== 'primitive') {
		return undefined;
	}
	const normalized = normalizeType(shape.asType);
	if (!normalized || normalized === 'variant' || normalized === 'object') {
		return undefined;
	}
	if (isKnownScalarType(normalized)) {
		return `it is declared As ${shape.asType}`;
	}
	return undefined;
}

function forEachSourceTypeProblem(
	shape: DeclaredValueShape,
	opts: AnalyzeModuleOptions,
): string | undefined {
	if (shape.isArray || !shape.asType) {
		return undefined;
	}
	const resolved = resolveTypeName(shape.asType, {
		projectTypes: opts.projectTypes,
		model: opts.hostModel,
	});
	if (resolved && resolved.kind !== 'primitive') {
		return undefined;
	}
	const normalized = normalizeType(shape.asType);
	if (!normalized || normalized === 'variant' || normalized === 'object') {
		return undefined;
	}
	if (isKnownScalarType(normalized)) {
		return `it is declared As ${shape.asType}`;
	}
	return undefined;
}

function simpleForEachSourceName(sourceExpression: string): string | undefined {
	const toks = statementTokens(sourceExpression, { start: 0, end: sourceExpression.length });
	return toks.length === 1 ? tokenName(toks[0]) : undefined;
}

export function checkElseBranchOrder(
	source: string,
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	checkConditionalCompilationElseBranchOrder(source, mod, push);
	checkIfBlockElseBranchOrder(source, mod, activity, push);
}

function checkConditionalCompilationElseBranchOrder(
	source: string,
	mod: ModuleNode,
	push: PushFn,
): void {
	for (const issue of scanConditionalCompilationBranchOrder(mod).issues) {
		if (issue.kind === 'elseifAfterElse') {
			push(
				'elseBranchOrder',
				"'#ElseIf' cannot appear after '#Else' in the same conditional-compilation block.",
				conditionalDirectiveKeywordSpan(source, issue.directive),
			);
		} else {
			push(
				'elseBranchOrder',
				"Only one '#Else' branch is allowed in a conditional-compilation block.",
				conditionalDirectiveKeywordSpan(source, issue.directive),
			);
		}
	}
}

function checkIfBlockElseBranchOrder(
	source: string,
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind === 'Procedure') {
			checkIfBlockElseBranchOrderInBody(source, member.body, activity, push);
		}
	}
}

function checkIfBlockElseBranchOrderInBody(
	source: string,
	body: BodyNode[],
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	for (const node of body) {
		if (isInactiveNode(activity, node)) {
			continue;
		}
		if (node.kind === 'IfBlock') {
			checkSingleIfBlockElseBranchOrder(source, node, activity, push);
		}
		if ('body' in node && Array.isArray(node.body)) {
			checkIfBlockElseBranchOrderInBody(source, node.body, activity, push);
		}
	}
}

function checkSingleIfBlockElseBranchOrder(
	source: string,
	node: { body: BodyNode[] },
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	let seenElse = false;
	for (const child of node.body) {
		if (isInactiveNode(activity, child)) {
			continue;
		}
		if (child.kind !== 'Statement') {
			continue;
		}
		const toks = statementTokensAfterLeadingLabel(source, child.span);
		const first = toks[0];
		const word = first ? tokenText(first) : undefined;
		if (word === 'elseif' && seenElse) {
			push(
				'elseBranchOrder',
				"'ElseIf' cannot appear after 'Else' in the same If block.",
				absoluteSpan(child.span, first),
			);
		} else if (word === 'else') {
			if (seenElse) {
				push(
					'elseBranchOrder',
					"Only one 'Else' branch is allowed in an If block.",
					absoluteSpan(child.span, first),
				);
			}
			seenElse = true;
		}
	}
}

function conditionalDirectiveKeywordSpan(
	source: string,
	directive: { span: Span },
): Span {
	const tokens = tokenize(source.slice(directive.span.start, directive.span.end))
		.filter((token) => token.kind !== 'comment' && token.kind !== 'newline');
	const marker = tokens[0];
	const keyword = tokens[1];
	if (marker?.kind === 'directive' && keyword) {
		return {
			start: directive.span.start + marker.start,
			end: directive.span.start + keyword.end,
		};
	}
	return directive.span;
}

/**
 * Rule: a `Select Case` block may contain at most one `Case Else`. VBE rejects a
 * second `Case Else` (it reports the later one as "Case without Select Case").
 * MS-VBAL 5.4.2.10; oracle-verified `duplicate_case_else_compile`. Only the
 * block's own direct `Case Else` clauses are counted (nested Selects are handled
 * independently), and inactive-branch clauses are skipped (no false positive).
 */
export function checkDuplicateCaseElse(
	source: string,
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		forEachSelectBlock(member.body, activity, (select) => {
			let seenCaseElse = false;
			for (const node of select.body) {
				// Only provably-active Case Else clauses collide; clauses in an
				// inactive or unknown #If branch are not guaranteed to compile
				// together, so flagging them would be a false positive.
				if (!isLeafStatement(node) || (activity && activity.activityForSpan(node.span) !== 'active')) {
					continue;
				}
				const toks = statementTokensAfterLeadingLabel(source, node.span);
				const caseTok = toks[0];
				const elseTok = toks[1];
				if (!caseTok || !elseTok || tokenText(caseTok) !== 'case' || tokenText(elseTok) !== 'else') {
					continue;
				}
				if (seenCaseElse) {
					push(
						'duplicateCaseElse',
						"A 'Select Case' block can have only one 'Case Else'.",
						{
							start: absoluteSpan(node.span, caseTok).start,
							end: absoluteSpan(node.span, elseTok).end,
						},
					);
				} else {
					seenCaseElse = true;
				}
			}
		});
	}
}

/** Visits every active `SelectBlock` in a body, descending into nested blocks. */
function forEachSelectBlock(
	body: BodyNode[],
	activity: ConditionalActivityTracker | undefined,
	visit: (select: SelectBlockNode) => void,
): void {
	for (const node of body) {
		if (isInactiveNode(activity, node)) {
			continue;
		}
		if (node.kind === 'SelectBlock') {
			visit(node);
		}
		if ('body' in node && Array.isArray((node as { body?: unknown }).body)) {
			forEachSelectBlock((node as { body: BodyNode[] }).body, activity, visit);
		}
	}
}

/**
 * Rule: an `Else` or `ElseIf` clause may only appear inside an `If` block. A
 * stray one (an `Else`/`ElseIf` with no opening `If ... Then` block) is a compile
 * error - VBE reports "Else without If". Oracle-verified
 * (`corpus_ctrl_if_004_compile`). The block-balance pass already covers a stray
 * `End If`, but `Else`/`ElseIf` are not block closers, so they need this check.
 * Legit clauses - which the parser places directly inside an `IfBlock` body - are
 * never flagged (no false positive).
 */
export function checkElseWithoutIf(
	source: string,
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	const visit = (body: BodyNode[], insideIfBlock: boolean): void => {
		for (const node of body) {
			if (isInactiveNode(activity, node)) {
				continue;
			}
			if (isLeafStatement(node)) {
				if (insideIfBlock) {
					continue; // a legit Else/ElseIf header inside its If block
				}
				const first = statementTokensAfterLeadingLabel(source, node.span)[0];
				const word = first ? tokenText(first) : '';
				if (word === 'else' || word === 'elseif') {
					push(
						'elseWithoutIf',
						`'${first!.rawText}' can only appear inside an 'If' block.`,
						absoluteSpan(node.span, first!),
					);
				}
				continue;
			}
			if ('body' in node && Array.isArray((node as { body?: unknown }).body)) {
				visit((node as { body: BodyNode[] }).body, node.kind === 'IfBlock');
			}
		}
	};
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind === 'Procedure') {
			visit(member.body, false);
		}
	}
}

function checkContextStatement(
	source: string,
	stmt: LeafStatementNode,
	ctx: StatementContext,
	push: PushFn,
): void {
	const toks = statementTokensAfterLeadingLabel(source, stmt.span);
	const first = toks[0];
	if (!first) {
		return;
	}
	const w0 = tokenText(first);

	if (w0 === 'if' && !toks.some((t) => tokenText(t) === 'then')) {
		push(
			'ifMissingThen',
			"If statement is missing 'Then'.",
			absoluteSpan(stmt.span, first),
		);
	}

	if (w0 === 'case' && ctx.selectDepth === 0) {
		push(
			'caseOutsideSelect',
			"'Case' can only appear inside a 'Select Case' block.",
			absoluteSpan(stmt.span, first),
		);
	}

	if (first.rawText === '.' && ctx.withDepth === 0) {
		push(
			'memberAccessOutsideWith',
			"A statement that starts with '.' must be inside a With block.",
			absoluteSpan(stmt.span, first),
		);
	}
	const leadingMember = toks[1];
	if (first.rawText === '.' && ctx.withDepth > 0 && (!leadingMember || !tokenName(leadingMember))) {
		push(
			'invalidExpressionSyntax',
			"Incomplete member access: type a member name after '.'.",
			absoluteSpan(stmt.span, first),
		);
	}

	if (w0 === 'exit') {
		const target = toks[1];
		const targetWord = tokenText(target);
		if (target && targetWord === 'for' && ctx.forDepth === 0) {
			push(
				'exitOutsideBlock',
				"'Exit For' can only appear inside a For loop.",
				exitPhraseSpan(stmt.span, first, target),
			);
		} else if (target && targetWord === 'do' && ctx.doDepth === 0) {
			push(
				'exitOutsideBlock',
				"'Exit Do' can only appear inside a Do loop.",
				exitPhraseSpan(stmt.span, first, target),
			);
		}
	}
}

function exitPhraseSpan(base: Span, first: VbaToken, target: VbaToken): Span {
	return { start: base.start + first.start, end: base.start + target.end };
}
