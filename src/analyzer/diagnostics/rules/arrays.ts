// Rule family: array declaration and allocation rules (audit #0).
//
// Extracted verbatim from analyzeModule.ts: ReDim target/bounds validation,
// unallocated dynamic-array access, Erase targets, and LBound/UBound argument
// checks.

import type { ConditionalActivityTracker } from '../../conditional/conditionalCompilation';
import { parseVbaIntegerLiteral } from '../../constants/integerConstantExpression';
import type { VbaToken } from '../../lexer/tokenKinds';
import type {
	BodyNode,
	ModuleNode,
	Span,
	LeafStatementNode,
	VariableDeclNode,
	VariableGroupNode,
} from '../../parser/nodes';
import { buildModuleSymbols } from '../../symbols/buildModuleSymbols';
import type { VbaSymbol } from '../../symbols/symbolModel';
import {
	procedureSymbolFor,
	type PushFn,
} from '../analysisContext';
import { splitArgSlots } from '../callExtraction';
import { walkBranchMergedBody, walkStraightLineBody } from '../dataflow';
import { procedureHasUnstructuredFlow } from '../../flow/procedureUnstructured';
import { isBareOrVbaQualifiedIntrinsicCall } from '../rules/shared';
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
	bareAssignmentTarget,
	forEachStatement,
	forEachVariableGroup,
	isInactiveNode,
	localsPassedAsCallArguments,
	matchParenFrom,
	pluralizeCount,
	statementTokens,
	statementTokensAfterLeadingLabel,
	tokenName,
	tokenText,
	type ProcedureStatementVisitor,
} from '../walker';

/** Per-statement rule: rides the shared procedure-statement walk (audit #0). */
export function checkArrayBoundIntrinsicArguments(
	source: string,
	symbols: ReturnType<typeof buildModuleSymbols>,
	projectVisibleSymbols: readonly VbaSymbol[] | undefined,
	push: PushFn,
): ProcedureStatementVisitor {
	return (member) => {
		const shapes = declarationShapeEnvironmentFor(symbols, member);
		const procSym = procedureSymbolFor(symbols, member);
		return (stmt) => {
			for (const hit of arrayBoundScalarArguments(
				source,
				stmt.span,
				shapes,
				(name) => declaredShapeForSourceBinding(
					symbols,
					procSym,
					projectVisibleSymbols,
					name,
					'expression',
				),
			)) {
				push(
					'arrayBoundRequiresArray',
					`${hit.functionName} requires an array argument, but '${hit.name}' is declared As ${hit.asType}.`,
					hit.span,
				);
			}
		};
	};
}

interface ArrayBoundScalarArgument {
	functionName: string;
	name: string;
	span: Span;
	asType: string;
}

function arrayBoundScalarArguments(
	source: string,
	span: Span,
	shapes: ReadonlyMap<string, DeclaredValueShape>,
	resolveShape?: (name: string) => SourceDeclaredShape,
): ArrayBoundScalarArgument[] {
	const toks = statementTokens(source, span);
	const hits: ArrayBoundScalarArgument[] = [];
	for (let i = 0; i < toks.length - 2; i++) {
		const functionName = tokenName(toks[i]);
		const lower = functionName?.toLowerCase();
		if (lower !== 'lbound' && lower !== 'ubound') {
			continue;
		}
		if (toks[i + 1]?.rawText !== '(' || !isBareOrVbaQualifiedIntrinsicCall(toks, i)) {
			continue;
		}
		const close = matchParenFrom(toks, i + 1);
		if (close < 0) {
			continue;
		}
		const inner = toks.slice(i + 2, close);
		if (inner.length === 0) {
			continue;
		}
		const split = splitArgSlots(inner, span.start);
		const firstSlot = split.slots[0] ?? [];
		if (firstSlot.length !== 1) {
			continue;
		}
		const argName = tokenName(firstSlot[0]);
		if (!argName) {
			continue;
		}
		const resolvedShape = resolveShape?.(argName);
		const shape = resolvedShape?.resolved
			? resolvedShape.shape
			: shapes.get(argName.toLowerCase());
		if (!shape || shape.isArray || !shape.asType) {
			continue;
		}
		const normalized = normalizeType(shape.asType);
		if (!normalized || !isKnownScalarType(normalized)) {
			continue;
		}
		hits.push({
			functionName: functionName!,
			name: argName,
			span: split.spans[0] ?? { start: span.start + firstSlot[0].start, end: span.start + firstSlot[0].end },
			asType: shape.asType,
		});
	}
	return hits;
}

interface RedimBlockedDeclaration {
	name: string;
	span: Span;
	kind: 'fixedArray' | 'scalar';
}

interface RedimDimension {
	key?: string;
	lowerKey?: string;
	lowerValue?: number;
	upperValue?: number;
	span: Span;
}

interface RedimTarget {
	name: string;
	span: Span;
	preserve: boolean;
	dimensions: RedimDimension[];
}

/**
 * Rule: ReDim can allocate dynamic arrays, but it cannot resize a variable that
 * was already declared as a scalar or as a fixed-size array.
 */
export function checkInvalidRedimTargets(
	source: string,
	mod: ModuleNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
	projectVisibleSymbols: readonly VbaSymbol[] | undefined,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): ProcedureStatementVisitor {
	const moduleDeclarations = redimBlockedDeclarationsForModule(mod, activity);
	return (member) => {
		const localDeclarations = redimBlockedDeclarationsForBody(member.body, activity);
		const localNames = declarationNamesForBody(member.body, activity);
		const procSym = procedureSymbolFor(symbols, member);
		return (stmt) => {
			for (const target of redimTargets(source, stmt.span)) {
				const lower = target.name.toLowerCase();
				const resolvedShape = declaredShapeForSourceBinding(
					symbols,
					procSym,
					projectVisibleSymbols,
					target.name,
					'assignmentTarget',
				);
				const declaration = resolvedShape.resolved
					? redimBlockedDeclarationForShape(target.name, target.span, resolvedShape.shape)
					: localDeclarations.get(lower) ??
						(localNames.has(lower) ? undefined : moduleDeclarations.get(lower));
				if (!declaration) {
					continue;
				}
				if (declaration.kind === 'scalar') {
					push(
						'scalarRedim',
						`Scalar variable '${target.name}' cannot be resized with ReDim; declare it as a dynamic array first.`,
						target.span,
					);
					continue;
				}
				push(
					'fixedArrayRedim',
					`Fixed-size array '${target.name}' cannot be resized with ReDim.`,
					target.span,
				);
			}
		};
	};
}

function redimBlockedDeclarationForShape(
	name: string,
	span: Span,
	shape: DeclaredValueShape | undefined,
): RedimBlockedDeclaration | undefined {
	if (!shape) {
		return undefined;
	}
	if (!shape.isArray) {
		if (isVariantLikeRedimTargetType(shape.asType)) {
			return undefined;
		}
		return { name, span, kind: 'scalar' };
	}
	if (shape.isFixedArray) {
		return { name, span, kind: 'fixedArray' };
	}
	return undefined;
}

function isVariantLikeRedimTargetType(asType: string | undefined): boolean {
	return !asType || normalizeType(asType) === 'variant';
}

function redimBlockedDeclarationsForModule(
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
): Map<string, RedimBlockedDeclaration> {
	const declarations = new Map<string, RedimBlockedDeclaration>();
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind === 'VariableGroup') {
			addRedimBlockedDeclarations(member, declarations);
		}
	}
	return declarations;
}

function redimBlockedDeclarationsForBody(
	body: BodyNode[],
	activity: ConditionalActivityTracker | undefined,
): Map<string, RedimBlockedDeclaration> {
	const declarations = new Map<string, RedimBlockedDeclaration>();
	forEachVariableGroup(body, (group) => addRedimBlockedDeclarations(group, declarations), activity);
	return declarations;
}

function declarationNamesForBody(
	body: BodyNode[],
	activity: ConditionalActivityTracker | undefined,
): Set<string> {
	const names = new Set<string>();
	forEachVariableGroup(body, (group) => {
		for (const decl of group.declarations) {
			names.add(decl.name.toLowerCase());
		}
	}, activity);
	return names;
}

function addRedimBlockedDeclarations(
	group: VariableGroupNode,
	out: Map<string, RedimBlockedDeclaration>,
): void {
	for (const decl of group.declarations) {
		const kind = redimBlockedDeclarationKind(decl);
		if (!kind) {
			continue;
		}
		const lower = decl.name.toLowerCase();
		if (!out.has(lower)) {
			out.set(lower, { name: decl.name, span: decl.span, kind });
		}
	}
}

function redimBlockedDeclarationKind(decl: VariableDeclNode): RedimBlockedDeclaration['kind'] | undefined {
	if (!decl.isArray) {
		if (isVariantLikeRedimTargetType(decl.asType)) {
			return undefined;
		}
		return 'scalar';
	}
	return decl.arrayBounds ? 'fixedArray' : undefined;
}

function redimTargets(source: string, span: Span): Array<{ name: string; span: Span }> {
	return redimStatementTargets(source, span).map((target) => ({
		name: target.name,
		span: target.span,
	}));
}

function redimStatementTargets(source: string, span: Span): RedimTarget[] {
	const toks = statementTokensAfterLeadingLabel(source, span);
	if (tokenText(toks[0]) !== 'redim') {
		return [];
	}
	const preserve = tokenText(toks[1]) === 'preserve';
	const start = preserve ? 2 : 1;
	const out: RedimTarget[] = [];
	for (const group of splitTopLevelTokenGroups(toks.slice(start), ',')) {
		const target = redimTargetFromGroup(span, group, preserve);
		if (target) {
			out.push(target);
		}
	}
	return out;
}

function redimTargetFromGroup(
	base: Span,
	group: readonly VbaToken[],
	preserve: boolean,
): RedimTarget | undefined {
	const content = group.filter((tok) => tok.kind !== 'comment');
	const nameTok = content[0];
	const name = tokenName(nameTok);
	if (!name || !nameTok) {
		return undefined;
	}
	const dimensions: RedimDimension[] = [];
	if (content[1]?.rawText === '(') {
		const close = matchParenFrom(content, 1);
		if (close > 1) {
			for (const part of splitTopLevelTokenGroups(content.slice(2, close), ',')) {
				const dimTokens = part.filter((tok) => tok.kind !== 'comment');
				if (dimTokens.length === 0) {
					continue;
				}
				const bound = comparableArrayBoundKey(dimTokens);
				dimensions.push({
					key: bound.key,
					lowerKey: bound.lowerKey,
					lowerValue: bound.lowerValue,
					upperValue: bound.upperValue,
					span: tokenGroupSpan(base, dimTokens),
				});
			}
		}
	}
	return {
		name,
		span: absoluteSpan(base, nameTok),
		preserve,
		dimensions,
	};
}

function comparableArrayBoundKey(
	toks: readonly VbaToken[],
): { key?: string; lowerKey?: string; lowerValue?: number; upperValue?: number } {
	const toIndex = toks.findIndex((tok) => tokenText(tok) === 'to');
	if (toIndex > 0) {
		const lower = comparableArrayBoundExpression(toks.slice(0, toIndex));
		const upper = comparableArrayBoundExpression(toks.slice(toIndex + 1));
		return {
			key: lower.key && upper.key ? `${lower.key}to${upper.key}` : undefined,
			lowerKey: lower.key,
			lowerValue: lower.value,
			upperValue: upper.value,
		};
	}
	const upper = comparableArrayBoundExpression(toks);
	return { key: upper.key, upperValue: upper.value };
}

function comparableArrayBoundExpression(toks: readonly VbaToken[]): { key?: string; value?: number } {
	return {
		key: comparableArrayBoundExpressionKey(toks),
		value: comparableArrayBoundExpressionValue(toks),
	};
}

function comparableArrayBoundExpressionKey(toks: readonly VbaToken[]): string | undefined {
	const parts: string[] = [];
	for (const tok of toks) {
		const word = tokenText(tok);
		if (
			tok.kind === 'integerLiteral' ||
			tok.rawText === '+' ||
			tok.rawText === '-' ||
			word === 'to'
		) {
			parts.push(word || tok.rawText.toLowerCase());
			continue;
		}
		return undefined;
	}
	return parts.length > 0 ? parts.join('') : undefined;
}

function comparableArrayBoundExpressionValue(toks: readonly VbaToken[]): number | undefined {
	let value = 0;
	let sign = 1;
	let expectingValue = true;
	let sawValue = false;
	for (const tok of toks) {
		if (expectingValue) {
			if (tok.rawText === '+' || tok.rawText === '-') {
				sign *= tok.rawText === '-' ? -1 : 1;
				continue;
			}
			if (tok.kind !== 'integerLiteral') {
				return undefined;
			}
			const parsed = parseVbaIntegerLiteral(tok.rawText);
			if (parsed === undefined) {
				return undefined;
			}
			const next = value + sign * parsed;
			if (!Number.isSafeInteger(next)) {
				return undefined;
			}
			value = next;
			sign = 1;
			expectingValue = false;
			sawValue = true;
			continue;
		}
		if (tok.rawText === '+' || tok.rawText === '-') {
			sign = tok.rawText === '-' ? -1 : 1;
			expectingValue = true;
			continue;
		}
		return undefined;
	}
	return sawValue && !expectingValue ? value : undefined;
}

/**
 * Rule: ReDim lower bounds must not be greater than their upper bounds.
 * This only reports explicit, literal-style `lower To upper` dimensions.
 */
export function checkRedimImpossibleBounds(
	source: string,
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): ProcedureStatementVisitor {
	const moduleDeclarations = redimBlockedDeclarationsForModule(mod, activity);
	return (member) => {
		const localDeclarations = redimBlockedDeclarationsForBody(member.body, activity);
		const localNames = declarationNamesForBody(member.body, activity);
		return (stmt) => {
			for (const target of redimStatementTargets(source, stmt.span)) {
				const lowerName = target.name.toLowerCase();
				const blockedDeclaration = localDeclarations.get(lowerName) ??
					(localNames.has(lowerName) ? undefined : moduleDeclarations.get(lowerName));
				if (blockedDeclaration) {
					continue;
				}
				target.dimensions.forEach((dimension, index) => {
					if (
						dimension.lowerValue === undefined ||
						dimension.upperValue === undefined ||
						dimension.lowerValue <= dimension.upperValue
					) {
						return;
					}
					push(
						'redimImpossibleBounds',
						`ReDim lower bound ${dimension.lowerValue} is greater than upper bound ${dimension.upperValue} for dimension ${index + 1} of '${target.name}'; this will raise Run-time error '9': Subscript out of range.`,
						dimension.span,
					);
				});
			}
		};
	};
}

/** VBA allows at most 60 array dimensions. */
const MAX_ARRAY_DIMENSIONS = 60;

/**
 * Rule family on `Dim`/`Static`/`Private`/`Public` array *declarations*:
 *  - `array-declaration-impossible-bounds`: an explicit literal `lower To upper`
 *    dimension with `lower > upper` (e.g. `Dim a(10 To 1)`). Only literal bounds
 *    are reported; variable/constant-reference bounds stay quiet (no-FP).
 *  - `too-many-array-dimensions`: more than 60 dimensions (the VBA maximum;
 *    oracle-verified `corpus_array_limit_001b_compile`).
 * ReDim is covered separately by checkRedimImpossibleBounds.
 */
export function checkArrayDeclarationBounds(
	source: string,
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	const inspectGroup = (group: VariableGroupNode): void => {
		for (const decl of group.declarations) {
			if (!decl.isArray || decl.arrayBounds === undefined || isInactiveNode(activity, decl)) {
				continue;
			}
			inspectArrayDeclaration(source, decl, push);
		}
	};
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind === 'VariableGroup') {
			inspectGroup(member);
		} else if (member.kind === 'Procedure') {
			forEachVariableGroup(member.body, inspectGroup, activity);
		}
	}
}

function inspectArrayDeclaration(
	source: string,
	decl: VariableDeclNode,
	push: PushFn,
): void {
	const toks = statementTokens(source, decl.span);
	const open = toks.findIndex((tok) => tok.rawText === '(');
	if (open < 0) {
		return;
	}
	const close = matchParenFrom(toks, open);
	if (close < 0) {
		return;
	}
	const dims = splitTopLevelTokenGroups(toks.slice(open + 1, close), ',')
		.map((part) => part.filter((tok) => tok.kind !== 'comment'))
		.filter((dimTokens) => dimTokens.length > 0);

	if (dims.length > MAX_ARRAY_DIMENSIONS) {
		push(
			'tooManyArrayDimensions',
			`Array '${decl.name}' has ${dims.length} dimensions; VBA allows at most ${MAX_ARRAY_DIMENSIONS}.`,
			decl.nameSpan ?? decl.span,
		);
	}

	dims.forEach((dimTokens, index) => {
		const bound = comparableArrayBoundKey(dimTokens);
		if (
			bound.lowerValue === undefined ||
			bound.upperValue === undefined ||
			bound.lowerValue <= bound.upperValue
		) {
			return;
		}
		push(
			'arrayDeclarationImpossibleBounds',
			`Array '${decl.name}' lower bound ${bound.lowerValue} is greater than upper bound ${bound.upperValue} for dimension ${index + 1}; this is not a valid array bound.`,
			tokenGroupSpan(decl.span, dimTokens),
		);
	});
}

/**
 * Rule: ReDim Preserve may only resize the last dimension of an already
 * allocated dynamic array. This tracks simple, active ReDim shapes in a
 * conservative per-body flow so nested branch updates do not leak outward.
 */
export function checkRedimPreserveDimensions(
	source: string,
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		checkRedimPreserveDimensionsInBody(source, member.body, new Map(), activity, push);
	}
}

function checkRedimPreserveDimensionsInBody(
	source: string,
	body: BodyNode[],
	initialShapes: ReadonlyMap<string, RedimTarget>,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	const shapes = new Map(initialShapes);
	for (const node of body) {
		if (isInactiveNode(activity, node)) {
			continue;
		}
		if (node.kind === 'Statement') {
			for (const target of redimStatementTargets(source, node.span)) {
				if (target.preserve) {
					const previous = shapes.get(target.name.toLowerCase());
					const reason = previous
						? redimPreserveDimensionMismatch(previous, target)
						: undefined;
					if (reason) {
						push(
							'redimPreserveDimensionChange',
							`ReDim Preserve can only resize the last dimension of '${target.name}'. ${reason}`,
							target.span,
						);
					}
				}
				if (target.dimensions.length > 0) {
					shapes.set(target.name.toLowerCase(), target);
				}
			}
			continue;
		}
		if ('body' in node && Array.isArray((node as { body?: unknown }).body)) {
			checkRedimPreserveDimensionsInBody(
				source,
				(node as { body: BodyNode[] }).body,
				shapes,
				activity,
				push,
			);
		}
	}
}

function redimPreserveDimensionMismatch(
	previous: RedimTarget,
	current: RedimTarget,
): string | undefined {
	if (
		previous.dimensions.length > 0 &&
		current.dimensions.length > 0 &&
		previous.dimensions.length !== current.dimensions.length
	) {
		return `Previous ReDim has ${pluralizeCount(previous.dimensions.length, 'dimension')}, but this ReDim Preserve has ${current.dimensions.length}.`;
	}
	const comparableCount = Math.min(previous.dimensions.length, current.dimensions.length) - 1;
	for (let i = 0; i < comparableCount; i++) {
		const before = previous.dimensions[i]?.key;
		const after = current.dimensions[i]?.key;
		if (before && after && before !== after) {
			return `Dimension ${i + 1} changes before the final dimension.`;
		}
	}
	const finalIndex = Math.min(previous.dimensions.length, current.dimensions.length) - 1;
	if (finalIndex >= 0) {
		const beforeLower = previous.dimensions[finalIndex]?.lowerKey;
		const afterLower = current.dimensions[finalIndex]?.lowerKey;
		if (beforeLower && afterLower && beforeLower !== afterLower) {
			return `The lower bound of dimension ${finalIndex + 1} changes under Preserve.`;
		}
	}
	return undefined;
}

interface DynamicArrayDeclaration {
	name: string;
	span: Span;
}

type DynamicArrayAllocationState = 'unallocated' | 'allocated' | 'unknown';

/**
 * Rule: a local dynamic array declared as `Dim values() As T` has no storage
 * until ReDim allocates it. This tracks only straight-line local state; nested
 * runtime blocks and helper calls make the state unknown instead of guessed.
 */
export function checkUnallocatedDynamicArrayAccess(
	source: string,
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		const arrays = localDynamicArrayDeclarationsForBody(member.body, activity);
		if (arrays.size === 0) {
			continue;
		}
		const state = new Map<string, DynamicArrayAllocationState>();
		for (const lower of arrays.keys()) {
			state.set(lower, 'unallocated');
		}
		const walk = procedureHasUnstructuredFlow(source, member, activity)
			? walkStraightLineBody
			: walkBranchMergedBody;
		walk(member.body, (node) => isInactiveNode(activity, node), {
			onStatement: (stmt) =>
				checkUnallocatedDynamicArrayAccessStatement(source, stmt, arrays, state, push),
			touchesInStatement: (stmt) => dynamicArrayTouchesInStatement(source, stmt, arrays),
			demoteToUnknown: (lower) => {
				state.set(lower, 'unknown');
			},
			snapshotState: () => new Map(state),
			restoreState: (snapshot) => {
				state.clear();
				for (const [key, value] of snapshot) {
					state.set(key, value as DynamicArrayAllocationState);
				}
			},
			setState: (key, value) => state.set(key, value as DynamicArrayAllocationState),
			lattice: { init: 'unallocated', good: 'allocated', unknown: 'unknown' },
		});
	}
}

function checkUnallocatedDynamicArrayAccessStatement(
	source: string,
	stmt: LeafStatementNode,
	arrays: ReadonlyMap<string, DynamicArrayDeclaration>,
	state: Map<string, DynamicArrayAllocationState>,
	push: PushFn,
): void {
	const redimmed = redimStatementTargets(source, stmt.span);
	if (redimmed.length > 0) {
		for (const target of redimmed) {
			const lower = target.name.toLowerCase();
			if (arrays.has(lower) && target.dimensions.length > 0) {
				state.set(lower, 'allocated');
			}
		}
		return;
	}
	const erased = eraseStatementSimpleTargets(source, stmt.span);
	if (erased.size > 0) {
		for (const lower of erased) {
			if (arrays.has(lower)) {
				state.set(lower, 'unallocated');
			}
		}
		return;
	}
	for (const hit of unallocatedDynamicArrayIndexAccesses(source, stmt.span, arrays, state)) {
		push(
			'unallocatedDynamicArrayAccess',
			`Dynamic array '${hit.name}' is not allocated before indexed access. This will raise Run-time error '9': Subscript out of range.`,
			hit.span,
		);
	}
	for (const hit of unallocatedDynamicArrayBoundCalls(source, stmt.span, arrays, state)) {
		push(
			'unallocatedDynamicArrayAccess',
			`Dynamic array '${hit.name}' is not allocated before ${hit.functionName}. This will raise Run-time error '9': Subscript out of range.`,
			hit.span,
		);
	}
	const assignment = bareAssignmentTarget(source, stmt.span);
	const assignmentLower = assignment?.name.toLowerCase();
	if (assignmentLower && arrays.has(assignmentLower)) {
		state.set(assignmentLower, 'unknown');
	}
	for (const lower of localsPassedAsCallArguments(source, stmt.span, arrays)) {
		if (state.get(lower) === 'unallocated') {
			state.set(lower, 'unknown');
		}
	}
}

function localDynamicArrayDeclarationsForBody(
	body: readonly BodyNode[],
	activity: ConditionalActivityTracker | undefined,
): Map<string, DynamicArrayDeclaration> {
	const out = new Map<string, DynamicArrayDeclaration>();
	forEachVariableGroup(body as BodyNode[], (group) => {
		if (group.isConst || group.modifier === 'Static') {
			return;
		}
		for (const decl of group.declarations) {
			if (!decl.isArray || decl.arrayBounds) {
				continue;
			}
			const lower = decl.name.toLowerCase();
			if (!out.has(lower)) {
				out.set(lower, { name: decl.name, span: decl.span });
			}
		}
	}, activity);
	return out;
}

function unallocatedDynamicArrayIndexAccesses(
	source: string,
	span: Span,
	arrays: ReadonlyMap<string, DynamicArrayDeclaration>,
	state: ReadonlyMap<string, DynamicArrayAllocationState>,
): Array<{ name: string; span: Span }> {
	const toks = statementTokensAfterLeadingLabel(source, span);
	const out: Array<{ name: string; span: Span }> = [];
	for (let i = 0; i < toks.length - 1; i++) {
		if (
			toks[i + 1].rawText !== '(' ||
			toks[i - 1]?.rawText === '.' ||
			toks[i - 1]?.rawText === '!'
		) {
			continue;
		}
		const name = tokenName(toks[i]);
		const lower = name?.toLowerCase();
		if (!name || !lower || !arrays.has(lower) || state.get(lower) !== 'unallocated') {
			continue;
		}
		const close = matchParenFrom(toks, i + 1);
		if (close <= i + 1) {
			continue;
		}
		out.push({
			name,
			span: { start: span.start + toks[i].start, end: span.start + toks[i].end },
		});
	}
	return out;
}

function unallocatedDynamicArrayBoundCalls(
	source: string,
	span: Span,
	arrays: ReadonlyMap<string, DynamicArrayDeclaration>,
	state: ReadonlyMap<string, DynamicArrayAllocationState>,
): Array<{ functionName: string; name: string; span: Span }> {
	const toks = statementTokens(source, span);
	const out: Array<{ functionName: string; name: string; span: Span }> = [];
	for (let i = 0; i < toks.length - 2; i++) {
		const functionName = tokenName(toks[i]);
		const lowerFunction = functionName?.toLowerCase();
		if (lowerFunction !== 'lbound' && lowerFunction !== 'ubound') {
			continue;
		}
		if (toks[i + 1]?.rawText !== '(' || !isBareOrVbaQualifiedIntrinsicCall(toks, i)) {
			continue;
		}
		const close = matchParenFrom(toks, i + 1);
		if (close < 0) {
			continue;
		}
		const split = splitArgSlots(toks.slice(i + 2, close), span.start);
		const firstSlot = split.slots[0] ?? [];
		if (firstSlot.length !== 1) {
			continue;
		}
		const name = tokenName(firstSlot[0]);
		const lower = name?.toLowerCase();
		if (!name || !lower || !arrays.has(lower) || state.get(lower) !== 'unallocated') {
			continue;
		}
		out.push({
			functionName: functionName!,
			name,
			span: split.spans[0] ?? { start: span.start + firstSlot[0].start, end: span.start + firstSlot[0].end },
		});
	}
	return out;
}

function dynamicArrayTouchesInStatement(
	source: string,
	stmt: LeafStatementNode,
	arrays: ReadonlyMap<string, DynamicArrayDeclaration>,
): Set<string> {
	const out = new Set<string>();
	for (const target of redimStatementTargets(source, stmt.span)) {
		const lower = target.name.toLowerCase();
		if (arrays.has(lower)) {
			out.add(lower);
		}
	}
	for (const lower of eraseStatementSimpleTargets(source, stmt.span)) {
		if (arrays.has(lower)) {
			out.add(lower);
		}
	}
	const assignment = bareAssignmentTarget(source, stmt.span);
	const assignmentLower = assignment?.name.toLowerCase();
	if (assignmentLower && arrays.has(assignmentLower)) {
		out.add(assignmentLower);
	}
	for (const lower of localsPassedAsCallArguments(source, stmt.span, arrays)) {
		out.add(lower);
	}
	return out;
}

function eraseStatementSimpleTargets(source: string, span: Span): Set<string> {
	const toks = statementTokensAfterLeadingLabel(source, span);
	if (tokenText(toks[0]) !== 'erase') {
		return new Set();
	}
	const out = new Set<string>();
	for (const group of splitTopLevelTokenGroups(toks.slice(1), ',')) {
		const content = group.filter((tok) => tok.kind !== 'comment');
		if (content.length !== 1) {
			continue;
		}
		const name = tokenName(content[0]);
		if (name) {
			out.add(name.toLowerCase());
		}
	}
	return out;
}

/**
 * Rule: Erase targets must be variable/array target names, not arbitrary
 * expressions. This intentionally stays syntax-shaped: array-ness/type
 * resolution is a separate binder-backed slice.
 */
export function checkEraseTargets(
	source: string,
	symbols: ReturnType<typeof buildModuleSymbols>,
	projectVisibleSymbols: readonly VbaSymbol[] | undefined,
	push: PushFn,
): ProcedureStatementVisitor {
	return (member) => {
		const shapes = declarationShapeEnvironmentFor(symbols, member);
		const procSym = procedureSymbolFor(symbols, member);
		return (stmt) => {
			for (const hit of invalidEraseTargets(source, stmt.span)) {
				push(
					'invalidEraseTarget',
					`Erase target must be a variable or array name, not an arbitrary expression.`,
					hit.span,
				);
			}
			for (const hit of eraseScalarTargets(
				source,
				stmt.span,
				shapes,
				(name) => declaredShapeForSourceBinding(
					symbols,
					procSym,
					projectVisibleSymbols,
					name,
					'assignmentTarget',
				),
			)) {
				push(
					'eraseRequiresArray',
					`Erase target '${hit.name}' must be an array or Variant, but it is declared As ${hit.asType}.`,
					hit.span,
				);
			}
		};
	};
}

function invalidEraseTargets(source: string, span: Span): Array<{ span: Span }> {
	const toks = statementTokensAfterLeadingLabel(source, span);
	if (tokenText(toks[0]) !== 'erase') {
		return [];
	}
	const out: Array<{ span: Span }> = [];
	for (const group of splitTopLevelTokenGroups(toks.slice(1), ',')) {
		const content = group.filter((tok) => tok.kind !== 'comment');
		if (content.length === 0) {
			continue;
		}
		if (eraseTargetLooksVariableLike(content)) {
			continue;
		}
		out.push({ span: tokenGroupSpan(span, content) });
	}
	return out;
}

function eraseScalarTargets(
	source: string,
	span: Span,
	shapes: ReadonlyMap<string, DeclaredValueShape>,
	resolveShape?: (name: string) => SourceDeclaredShape,
): Array<{ name: string; span: Span; asType: string }> {
	const toks = statementTokensAfterLeadingLabel(source, span);
	if (tokenText(toks[0]) !== 'erase') {
		return [];
	}
	const out: Array<{ name: string; span: Span; asType: string }> = [];
	for (const group of splitTopLevelTokenGroups(toks.slice(1), ',')) {
		const content = group.filter((tok) => tok.kind !== 'comment');
		if (content.length !== 1) {
			continue;
		}
		const name = tokenName(content[0]);
		if (!name) {
			continue;
		}
		const resolvedShape = resolveShape?.(name);
		const shape = resolvedShape?.resolved
			? resolvedShape.shape
			: shapes.get(name.toLowerCase());
		if (!shape || shape.isArray || !shape.asType) {
			continue;
		}
		const normalized = normalizeType(shape.asType);
		if (!normalized || normalized === 'variant') {
			continue;
		}
		if (normalized === 'object' || isKnownScalarType(normalized)) {
			out.push({ name, span: tokenGroupSpan(span, content), asType: shape.asType });
		}
	}
	return out;
}

function eraseTargetLooksVariableLike(toks: readonly VbaToken[]): boolean {
	if (!tokenName(toks[0])) {
		return false;
	}
	if (toks.some((tok) => ERASE_EXPRESSION_OPERATORS.has(tok.rawText))) {
		return false;
	}
	if (toks[0]?.rawText === '(') {
		return false;
	}
	return true;
}

const ERASE_EXPRESSION_OPERATORS = new Set([
	'+',
	'-',
	'*',
	'/',
	'\\',
	'^',
	'&',
	'=',
	'<',
	'>',
	'<=',
	'>=',
	'<>',
]);

function tokenGroupSpan(base: Span, toks: readonly VbaToken[]): Span {
	const first = toks[0];
	const last = toks[toks.length - 1];
	return {
		start: base.start + (first?.start ?? 0),
		end: base.start + (last?.end ?? 0),
	};
}

function splitTopLevelTokenGroups(
	toks: readonly VbaToken[],
	separator: string,
): VbaToken[][] {
	const groups: VbaToken[][] = [];
	let current: VbaToken[] = [];
	let depth = 0;
	for (const tok of toks) {
		if (tok.rawText === '(') {
			depth++;
		} else if (tok.rawText === ')') {
			depth = Math.max(0, depth - 1);
		}
		if (depth === 0 && tok.rawText === separator) {
			groups.push(current);
			current = [];
			continue;
		}
		current.push(tok);
	}
	groups.push(current);
	return groups;
}

interface FixedArrayBound {
	name: string;
	lowerValue?: number;
	upperValue: number;
	hasExplicitLower: boolean;
}

/**
 * Parses the single-dimension literal bounds of a fixed-size array declaration.
 * Returns undefined unless the declaration has exactly one dimension whose upper
 * bound folds to a literal integer (statically known). The lower bound is
 * reported only for an explicit literal `lower To upper` form; a single-bound
 * `Dim a(n)` leaves the lower bound Option-Base-dependent (0 or 1).
 */
function parseFixedArrayBoundsForDecl(
	source: string,
	decl: VariableDeclNode,
): { lowerValue?: number; upperValue: number; hasExplicitLower: boolean } | undefined {
	const toks = statementTokens(source, decl.span);
	const open = toks.findIndex((tok) => tok.rawText === '(');
	if (open < 0) {
		return undefined;
	}
	const close = matchParenFrom(toks, open);
	if (close < 0) {
		return undefined;
	}
	const dims = splitTopLevelTokenGroups(toks.slice(open + 1, close), ',')
		.map((part) => part.filter((tok) => tok.kind !== 'comment'))
		.filter((dimTokens) => dimTokens.length > 0);
	if (dims.length !== 1) {
		return undefined; // multi-dimension subscript matching is out of scope (v1)
	}
	const bound = comparableArrayBoundKey(dims[0]);
	if (bound.upperValue === undefined) {
		return undefined; // non-literal upper bound (e.g. a Const) is not statically known
	}
	return {
		lowerValue: bound.lowerValue,
		upperValue: bound.upperValue,
		hasExplicitLower: bound.lowerValue !== undefined,
	};
}

/** Local, single-dimension, statically-bounded fixed arrays in a procedure body. */
function localFixedArrayDeclarationsForBody(
	source: string,
	body: readonly BodyNode[],
	activity: ConditionalActivityTracker | undefined,
): Map<string, FixedArrayBound> {
	const out = new Map<string, FixedArrayBound>();
	forEachVariableGroup(body as BodyNode[], (group) => {
		if (group.isConst) {
			return;
		}
		for (const decl of group.declarations) {
			if (!decl.isArray || !decl.arrayBounds) {
				continue; // dynamic arrays (no static bounds) are out of scope
			}
			const lower = decl.name.toLowerCase();
			if (out.has(lower)) {
				continue;
			}
			const bounds = parseFixedArrayBoundsForDecl(source, decl);
			if (bounds) {
				out.set(lower, { name: decl.name, ...bounds });
			}
		}
	}, activity);
	return out;
}

/** Names that are ReDim targets anywhere in the body (defensive exclusion). */
function redimTargetNamesInBody(
	source: string,
	body: readonly BodyNode[],
	activity: ConditionalActivityTracker | undefined,
): Set<string> {
	const out = new Set<string>();
	forEachStatement(body as BodyNode[], (stmt) => {
		for (const target of redimStatementTargets(source, stmt.span)) {
			out.add(target.name.toLowerCase());
		}
	}, activity);
	return out;
}

/** Literal-subscript accesses of a tracked fixed array that fall outside its bounds. */
function fixedArraySubscriptViolations(
	source: string,
	span: Span,
	fixed: ReadonlyMap<string, FixedArrayBound>,
	excluded: ReadonlySet<string>,
): Array<{ span: Span; message: string }> {
	const toks = statementTokensAfterLeadingLabel(source, span);
	const out: Array<{ span: Span; message: string }> = [];
	for (let i = 0; i < toks.length - 1; i++) {
		if (
			toks[i + 1].rawText !== '(' ||
			toks[i - 1]?.rawText === '.' ||
			toks[i - 1]?.rawText === '!'
		) {
			continue;
		}
		const name = tokenName(toks[i]);
		const lower = name?.toLowerCase();
		if (!name || !lower || !fixed.has(lower) || excluded.has(lower)) {
			continue;
		}
		const close = matchParenFrom(toks, i + 1);
		if (close <= i + 1) {
			continue;
		}
		const argToks = toks.slice(i + 2, close).filter((tok) => tok.kind !== 'comment');
		const slots = splitTopLevelTokenGroups(argToks, ',');
		if (slots.length !== 1 || slots[0].length === 0) {
			continue; // not a single-subscript index access (multi-dim/empty) -> skip
		}
		const value = comparableArrayBoundExpressionValue(slots[0]);
		if (value === undefined) {
			continue; // non-literal subscript (variable / Const / member chain) -> not provable
		}
		const decl = fixed.get(lower)!;
		const lowGate = decl.hasExplicitLower ? decl.lowerValue! : 0;
		if (value <= decl.upperValue && value >= lowGate) {
			continue; // in bounds (lower of single-bound dims is Option-Base-dependent -> only < 0 flagged)
		}
		const slot = slots[0];
		const detail =
			value > decl.upperValue
				? `is above the array's declared upper bound ${decl.upperValue}`
				: decl.hasExplicitLower
					? `is below the array's declared lower bound ${decl.lowerValue}`
					: 'is negative and out of range';
		out.push({
			span: { start: span.start + slot[0].start, end: span.start + slot[slot.length - 1].end },
			message:
				`Subscript ${value} for array '${decl.name}' ${detail}. ` +
				`This will raise Run-time error '9': Subscript out of range.`,
		});
	}
	return out;
}

/**
 * Rule: a constant subscript proven outside a LOCAL fixed-size array's declared
 * bounds raises Run-time error '9' (oracle-verified `runtime006_*`). No-FP scope:
 * only local, single-dimension fixed arrays with a literal upper bound, accessed
 * with a literal (or folded signed-integer) subscript, are checked. Dynamic /
 * ReDim'd arrays, variable/Const subscripts, multi-dimension arrays, parameters,
 * and the Option-Base-dependent lower region of single-bound `Dim a(n)` decls
 * stay quiet. Flags subscripts above the upper bound, below an explicit literal
 * lower bound, or negative.
 */
export function checkFixedArraySubscriptBounds(
	source: string,
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		const fixed = localFixedArrayDeclarationsForBody(source, member.body, activity);
		if (fixed.size === 0) {
			continue;
		}
		const excluded = redimTargetNamesInBody(source, member.body, activity);
		forEachStatement(member.body, (stmt) => {
			for (const hit of fixedArraySubscriptViolations(source, stmt.span, fixed, excluded)) {
				push('arraySubscriptOutOfBounds', hit.message, hit.span);
			}
		}, activity);
	}
}
