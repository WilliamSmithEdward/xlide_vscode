// Rule family: array declaration and allocation rules (audit #0).
//
// Extracted verbatim from analyzeModule.ts: ReDim target/bounds validation,
// unallocated dynamic-array access, Erase targets, and LBound/UBound argument
// checks.

import type { ConditionalActivityTracker } from '../../conditional/conditionalCompilation';
import {
	bankersRound,
	evaluateIntegerConstantExpression,
	parseVbaIntegerLiteral,
	resolveRawIntegerConstants,
	type IntegerConstantLookup,
} from '../../constants/integerConstantExpression';
import type { HostObjectModel } from '../../host/excelObjectModel';
import { tokenize } from '../../lexer/tokenize';
import type { VbaToken } from '../../lexer/tokenKinds';
import type {
	BodyNode,
	ModuleNode,
	ProcedureNode,
	Span,
	LeafStatementNode,
	TypeFieldNode,
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
import { collectModuleLiteralIntegerConstants } from '../constExpr';
import { walkBranchMergedBody, walkStraightLineBody } from '../dataflow';
import { straightLineAssignments, type ReachingAssignments } from '../straightLineValues';
import { counterText, loopCountersAt, numericCounterPasses, type CounterValue, type CountersAt } from '../loopCounters';
import { procedureHasUnstructuredFlow } from '../../flow/procedureUnstructured';
import { isBareOrVbaQualifiedIntrinsicCall } from '../rules/shared';
import {
	declarationShapeEnvironmentFor,
	declaredShapeForSourceBinding,
	type DeclaredValueShape,
	isKnownScalarType,
	normalizeType,
	procedureIntegerConstantLookup,
	scopedIntegerConstantLookup,
	type SourceDeclaredShape,
} from '../typeInference';
import {
	absoluteSpan,
	activeModuleMembers,
	bareAssignmentTarget,
	forEachStatement,
	forEachStatementWithHeaders,
	forEachVariableGroup,
	isInactiveNode,
	localsNamedWhole,
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
	/** The element type an `As` after the bounds names, with its token. */
	asType?: { name: string; span: Span };
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
	return redimTargetsFromTokens(span, statementTokensAfterLeadingLabel(source, span));
}

function redimTargetsFromTokens(span: Span, toks: readonly VbaToken[]): RedimTarget[] {
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

/**
 * ReDim allocations embedded in a single-line `If cond Then ReDim a(...)` (and
 * its `Else ReDim b(...)` arm). The parser keeps single-line If statements as
 * one leaf, so redimStatementTargets - which requires `redim` as the first
 * token - never sees them, and the generic index-access scan would otherwise
 * flag the ReDim's own target as an unallocated access.
 */
function singleLineIfRedimTargets(source: string, span: Span): RedimTarget[] {
	const toks = statementTokensAfterLeadingLabel(source, span);
	if (tokenText(toks[0]) !== 'if') {
		return [];
	}
	const out: RedimTarget[] = [];
	for (let i = 1; i < toks.length - 1; i++) {
		const word = tokenText(toks[i]);
		if ((word !== 'then' && word !== 'else') || tokenText(toks[i + 1]) !== 'redim') {
			continue;
		}
		let end = toks.length;
		for (let j = i + 2; j < toks.length; j++) {
			if (tokenText(toks[j]) === 'else') {
				end = j;
				break;
			}
		}
		out.push(...redimTargetsFromTokens(span, toks.slice(i + 1, end)));
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
	// A qualified ReDim target (`ReDim x.arr(...)` or `ReDim x!arr(...)`) resizes
	// a member array, not the base variable. The scalar/fixed-array shape checks
	// only apply to a simple local/module variable, and the member's declared
	// shape is not resolvable here, so skip qualified targets rather than mistake
	// the container for the array being resized.
	if (content[1]?.rawText === '.' || content[1]?.rawText === '!') {
		return undefined;
	}
	const dimensions: RedimDimension[] = [];
	let asType: RedimTarget['asType'];
	if (content[1]?.rawText === '(') {
		const close = matchParenFrom(content, 1);
		const typeTokens = close > 1 && tokenText(content[close + 1]) === 'as' ? content.slice(close + 2) : [];
		if (typeTokens.length > 0 && typeTokens.every((tok) => tokenName(tok) || tok.rawText === '.')) {
			asType = {
				name: typeTokens.map((tok) => tok.rawText).join(''),
				span: tokenGroupSpan(base, typeTokens),
			};
		}
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
		...(asType ? { asType } : {}),
	};
}

/** The type a suffix character gives a name: `x$` is a String. */
const SUFFIX_TYPES: Readonly<Record<string, string>> = {
	'%': 'integer', '&': 'long', '^': 'longlong', '@': 'currency', '!': 'single', '#': 'double', '$': 'string',
};

/**
 * The element type a dynamic array was declared with, lowercased: `Dim x() As
 * String` is string and `Dim v()` Variant. Undefined for anything that is not
 * a dynamic array, and for a fixed-length string, whose ReDim is not judged.
 */
function dynamicArrayElementType(decl: VariableDeclNode): string | undefined {
	if (!decl.isArray || (decl.arrayBounds ?? '').trim() !== '' || decl.fixedLength !== undefined) {
		return undefined;
	}
	if (decl.asType) {
		return decl.asType.replace(/\s+/g, '').toLowerCase();
	}
	return decl.typeSuffix ? SUFFIX_TYPES[decl.typeSuffix] : 'variant';
}

/**
 * Rule: a ReDim may not give a dynamic array another element type. `Dim x()
 * As String` then `ReDim x(1) As Long` is "Can't change data types of array
 * elements" in the VBE, and so is `Dim v()` then `ReDim v(1) As Long`, with
 * Preserve or without; a Variant that is not an array takes any ReDim
 * (issue #212, measured in Excel 16.0).
 */
export function checkRedimTypeChange(
	source: string,
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): ProcedureStatementVisitor {
	const moduleArrays = new Map<string, string | undefined>();
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind === 'VariableGroup' && !member.isConst) {
			for (const decl of member.declarations) {
				const key = decl.name.toLowerCase();
				moduleArrays.set(key, moduleArrays.has(key) ? undefined : dynamicArrayElementType(decl));
			}
		}
	}
	return (member) => {
		const localArrays = new Map<string, string | undefined>();
		forEachVariableGroup(member.body, (group) => {
			for (const decl of group.declarations) {
				const key = decl.name.toLowerCase();
				localArrays.set(key, localArrays.has(key) ? undefined : dynamicArrayElementType(decl));
			}
		}, activity);
		const params = new Set(member.params.map((param) => param.name.toLowerCase()));
		return (stmt) => {
			for (const target of redimStatementTargets(source, stmt.span)) {
				if (!target.asType || /\*/.test(target.asType.name)) {
					continue;
				}
				const key = target.name.toLowerCase();
				if (params.has(key)) {
					continue;
				}
				const declared = localArrays.has(key) ? localArrays.get(key) : moduleArrays.get(key);
				if (declared === undefined || declared === target.asType.name.toLowerCase()) {
					continue;
				}
				push(
					'redimTypeChange',
					`'${target.name}' was declared with elements of another type, and a ReDim cannot change it to ${target.asType.name}. This is a VBE compile error: Can't change data types of array elements.`,
					target.asType.span,
				);
			}
		};
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

/**
 * Folds a `lower To upper` array bound that is built only from signed integer
 * literals (`-3`, `1 + 2`). This is intentionally a literal-only subset of
 * {@link evaluateIntegerConstantExpression} — variable/Const bounds are
 * deliberately left unevaluated so the rule stays quiet on them (see rule docs)
 * rather than reusing the full constant evaluator.
 */
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
 * Rule: ReDim lower bounds must not be greater than their upper bounds, a
 * bound being a literal or a constant expression (issue #209): `ReDim a(LO To
 * HI)` with LO = 5 and HI = 1 raises error 9. Also: a ReDim of more than 60
 * dimensions, which the VBE refuses as a syntax error (issue #209).
 */
export function checkRedimImpossibleBounds(
	source: string,
	mod: ModuleNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
	projectIntegerConstants: ReadonlyMap<string, string | undefined> | undefined,
	projectVisibleSymbols: readonly VbaSymbol[] | undefined,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
	hostModel?: HostObjectModel,
): ProcedureStatementVisitor {
	const moduleDeclarations = redimBlockedDeclarationsForModule(mod, activity);
	const optionBase = moduleOptionBase(mod, activity);
	const moduleConstants = moduleIntegerConstants(mod, projectIntegerConstants, activity);
	return (member) => {
		const localDeclarations = redimBlockedDeclarationsForBody(member.body, activity);
		const localNames = declarationNamesForBody(member.body, activity);
		let constants: IntegerConstantLookup | undefined;
		const lookup = (): IntegerConstantLookup => constants ??= procedureIntegerConstantLookup(
			member, moduleConstants, symbols, projectVisibleSymbols, activity, hostModel,
		);
		return (stmt) => {
			for (const target of redimStatementTargets(source, stmt.span)) {
				if (target.dimensions.length > MAX_ARRAY_DIMENSIONS) {
					push(
						'tooManyArrayDimensions',
						`ReDim of '${target.name}' has ${target.dimensions.length} dimensions; VBA allows at most ${MAX_ARRAY_DIMENSIONS}. This is a VBE compile error: Syntax error.`,
						target.span,
					);
				}
				const lowerName = target.name.toLowerCase();
				const blockedDeclaration = localDeclarations.get(lowerName) ??
					(localNames.has(lowerName) ? undefined : moduleDeclarations.get(lowerName));
				if (blockedDeclaration) {
					continue;
				}
				target.dimensions.forEach((dimension, index) => {
					// `ReDim a(-1)`: the lower bound is Option Base, 0 by default,
					// and an upper bound below it is the same impossibility as
					// `ReDim a(5 To 1)` (issue #120, measured in Excel 16.0).
					const bounds = impossibleBounds(source, dimension.span, lookup(), optionBase);
					if (!bounds) {
						return;
					}
					push(
						'redimImpossibleBounds',
						`ReDim lower bound ${bounds.lowerText} is greater than upper bound ${bounds.upper} for dimension ${index + 1} of '${target.name}'; this will raise Run-time error '9': Subscript out of range.`,
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
 * Rule family on `Dim`/`Static`/`Private`/`Public` array *declarations* and on
 * Type members:
 *  - `array-declaration-impossible-bounds`: a dimension whose lower bound is
 *    above its upper one, "Range has no values" in the VBE. A bound is a
 *    literal or a constant expression the module, the procedure, another
 *    module's Public Const, an Enum or the host defines (issue #209,
 *    measured in Excel 16.0): `Dim a(LO To HI)`, `Dim a(-LO)` below Option
 *    Base, `Dim a(1.5 To 1)`, which rounds to 2. A bound that names a
 *    variable stays quiet.
 *  - `too-many-array-dimensions`: more than 60 dimensions (the VBA maximum;
 *    oracle-verified `corpus_array_limit_001b_compile`), in a Type member too.
 * ReDim is covered separately by checkRedimImpossibleBounds.
 */
export function checkArrayDeclarationBounds(
	source: string,
	mod: ModuleNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
	projectIntegerConstants: ReadonlyMap<string, string | undefined> | undefined,
	projectVisibleSymbols: readonly VbaSymbol[] | undefined,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
	hostModel?: HostObjectModel,
): void {
	const optionBase = moduleOptionBase(mod, activity);
	const moduleConstants = moduleIntegerConstants(mod, projectIntegerConstants, activity);
	let moduleLookup: IntegerConstantLookup | undefined;
	const moduleScope = (): IntegerConstantLookup => moduleLookup ??= scopedIntegerConstantLookup(
		moduleConstants, symbols, undefined, projectVisibleSymbols, hostModel,
	);
	const inspectGroup = (group: VariableGroupNode, lookup: () => IntegerConstantLookup): void => {
		for (const decl of group.declarations) {
			if (!decl.isArray || decl.arrayBounds === undefined || isInactiveNode(activity, decl)) {
				continue;
			}
			inspectArrayDeclaration(source, decl, lookup, optionBase, push);
		}
	};
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind === 'VariableGroup') {
			inspectGroup(member, moduleScope);
		} else if (member.kind === 'Type') {
			for (const field of member.fields) {
				if (field.isArray && !isInactiveNode(activity, field)) {
					inspectArrayDeclaration(source, field, moduleScope, optionBase, push);
				}
			}
		} else if (member.kind === 'Procedure') {
			let constants: IntegerConstantLookup | undefined;
			const procedureScope = (): IntegerConstantLookup => constants ??= procedureIntegerConstantLookup(
				member, moduleConstants, symbols, projectVisibleSymbols, activity, hostModel,
			);
			forEachVariableGroup(member.body, (group) => inspectGroup(group, procedureScope), activity);
		}
	}
}

/** The module's integer constants and Enum members, over the project's Public ones. */
function moduleIntegerConstants(
	mod: ModuleNode,
	projectIntegerConstants: ReadonlyMap<string, string | undefined> | undefined,
	activity: ConditionalActivityTracker | undefined,
): Map<string, number | undefined> {
	const projectConstants = resolveRawIntegerConstants(projectIntegerConstants ?? new Map(), new Map());
	return collectModuleLiteralIntegerConstants(mod, activity, projectConstants);
}

/**
 * A dimension's bounds when the lower is above the upper, folded through
 * constants. With no `To`, the lower bound is Option Base. A bound that is a
 * lone float literal is rounded half to even, as VBA rounds it: `1.5 To 1` is
 * refused and `2.5 To 2` is not (issue #209). Undefined when either bound is
 * unknown or the bounds are fine.
 */
function impossibleBounds(
	source: string,
	span: Span,
	lookup: IntegerConstantLookup,
	optionBase: number,
): { lowerText: string; upper: number } | undefined {
	const text = source.slice(span.start, span.end);
	const toks = tokenize(text).filter((tok) => tok.kind !== 'comment' && tok.kind !== 'newline');
	let depth = 0;
	let to = -1;
	for (let i = 0; i < toks.length && to < 0; i++) {
		const raw = toks[i].rawText;
		depth += raw === '(' ? 1 : raw === ')' ? -1 : 0;
		if (depth === 0 && tokenText(toks[i]) === 'to') {
			to = i;
		}
	}
	const side = (part: readonly VbaToken[]): number | undefined => {
		if (part.length === 0) {
			return undefined;
		}
		if (part.length === 1 && part[0].kind === 'floatLiteral') {
			const value = Number(part[0].rawText.replace(/[!#@]$/, '').replace(/[dD]/, 'e'));
			return Number.isFinite(value) ? bankersRound(value) : undefined;
		}
		return evaluateIntegerConstantExpression(text.slice(part[0].start, part[part.length - 1].end), lookup);
	};
	const upper = side(to < 0 ? toks : toks.slice(to + 1));
	const lower = to < 0 ? optionBase : side(toks.slice(0, to));
	if (upper === undefined || lower === undefined || lower <= upper) {
		return undefined;
	}
	return { lowerText: to < 0 ? `${lower} (Option Base ${lower})` : String(lower), upper };
}

function inspectArrayDeclaration(
	source: string,
	decl: VariableDeclNode | TypeFieldNode,
	lookup: () => IntegerConstantLookup,
	optionBase: number,
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
		const span = tokenGroupSpan(decl.span, dimTokens);
		const bounds = impossibleBounds(source, span, lookup(), optionBase);
		if (!bounds) {
			return;
		}
		push(
			'arrayDeclarationImpossibleBounds',
			`Array '${decl.name}' lower bound ${bounds.lowerText} is greater than upper bound ${bounds.upper} for dimension ${index + 1}; this is not a valid array bound.`,
			span,
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
		walk(source, member.body, (node) => isInactiveNode(activity, node), {
			onStatement: (stmt) =>
				checkUnallocatedDynamicArrayAccessStatement(source, stmt, arrays, state, push),
			onBlock: (node) => {
				// `For Each x In a` over an array with no storage raises 92, For
				// loop not initialized, not 9 (issue #181, measured in Excel 16.0).
				const over = node.kind === 'ForBlock' && node.each ? node.sourceExpression?.trim().toLowerCase() : undefined;
				if (over && node.kind === 'ForBlock' && node.sourceExpressionSpan && state.get(over) === 'unallocated') {
					push(
						'unallocatedDynamicArrayAccess',
						`Dynamic array '${arrays.get(over)!.name}' is not allocated when For Each asks it for its elements. This will raise Run-time error '92': For loop not initialized.`,
						node.sourceExpressionSpan,
					);
				}
			},
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
	const conditionalRedims = singleLineIfRedimTargets(source, stmt.span);
	const passedWhole = localsNamedWhole(source, stmt.span, arrays, ARRAY_READ_ONLY_INTRINSICS);
	// An access that follows a whole-array pass in the same statement, as in
	// `If Load(a) Then Debug.Print a(0)`, runs after the callee had its chance
	// to allocate. One that precedes it, as in `Load(a(0))`, does not.
	const followsPass = (hit: { name: string; span: Span }): boolean => {
		const passAt = passedWhole.get(hit.name.toLowerCase());
		return passAt !== undefined && hit.span.start > passAt;
	};
	for (const hit of unallocatedDynamicArrayIndexAccesses(source, stmt.span, arrays, state)) {
		// The "access" may be the target of a ReDim embedded in a single-line
		// If...Then - the allocation itself, not a read. Suppress exactly that
		// target's name-token span; bounds expressions still report.
		if (conditionalRedims.some((t) => t.span.start === hit.span.start && t.span.end === hit.span.end)) {
			continue;
		}
		if (followsPass(hit)) {
			continue;
		}
		push(
			'unallocatedDynamicArrayAccess',
			`Dynamic array '${hit.name}' is not allocated before indexed access. This will raise Run-time error '9': Subscript out of range.`,
			hit.span,
		);
	}
	for (const hit of unallocatedDynamicArrayBoundCalls(source, stmt.span, arrays, state)) {
		if (followsPass(hit)) {
			continue;
		}
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
	for (const lower of passedWhole.keys()) {
		if (state.get(lower) === 'unallocated') {
			state.set(lower, 'unknown');
		}
	}
	// A conditional (single-line If) ReDim allocates only on one path, so move
	// the array to 'unknown' - mirroring how block-If allocations degrade - not
	// 'allocated'. The 'unallocated' guard keeps an already-allocated array precise.
	for (const target of conditionalRedims) {
		const lower = target.name.toLowerCase();
		if (arrays.has(lower) && target.dimensions.length > 0 && state.get(lower) === 'unallocated') {
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
	for (const target of [
		...redimStatementTargets(source, stmt.span),
		...singleLineIfRedimTargets(source, stmt.span),
	]) {
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
	for (const lower of localsNamedWhole(source, stmt.span, arrays, ARRAY_READ_ONLY_INTRINSICS).keys()) {
		out.add(lower);
	}
	return out;
}

/** Intrinsics that read an array argument and allocate nothing. */
const ARRAY_READ_ONLY_INTRINSICS: ReadonlySet<string> = new Set(['lbound', 'ubound', 'isarray']);

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

/** One dimension of an array whose bounds the code fixes. */
export interface ArrayDimensionBound {
	lower: number;
	upper: number;
	/** False when the lower bound came from Option Base rather than the text. */
	explicitLower: boolean;
}

/** An array whose every dimension's bounds are known from the text. */
export interface FixedArrayBound {
	name: string;
	dims: ArrayDimensionBound[];
	/** Where the bounds came from, for the message: 'Dim', 'Array(...)', 'Split(...)', 'Range(...).Value'. */
	origin: string;
}

/**
 * Parses the single-dimension literal bounds of a fixed-size array declaration.
 * Returns undefined unless the declaration has exactly one dimension whose upper
 * bound folds to a literal integer (statically known). The lower bound is
 * reported only for an explicit literal `lower To upper` form; a single-bound
 * `Dim a(n)` leaves the lower bound Option-Base-dependent (0 or 1).
 */
/**
 * Parses the literal bounds of a fixed-size array declaration, one entry per
 * dimension. Undefined unless every dimension's upper bound (and any explicit
 * lower bound) folds to a literal integer. A dimension with no `To` takes
 * Option Base as its lower bound (issue #120: `Option Base 1` then `Dim a(3)`
 * refuses `a(0)`).
 */
function parseFixedArrayBoundsForDecl(
	source: string,
	decl: VariableDeclNode,
	optionBase: number,
): ArrayDimensionBound[] | undefined {
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
	if (dims.length === 0) {
		return undefined;
	}
	const out: ArrayDimensionBound[] = [];
	for (const dim of dims) {
		const bound = comparableArrayBoundKey(dim);
		if (bound.upperValue === undefined) {
			return undefined; // a Const or variable bound is not statically known
		}
		const hasTo = dim.some((tok) => tokenText(tok) === 'to');
		if (hasTo && bound.lowerValue === undefined) {
			return undefined;
		}
		out.push({
			lower: bound.lowerValue ?? optionBase,
			upper: bound.upperValue,
			explicitLower: bound.lowerValue !== undefined,
		});
	}
	return out;
}

/** Local, statically-bounded fixed arrays in a procedure body. */
function localFixedArrayDeclarationsForBody(
	source: string,
	body: readonly BodyNode[],
	activity: ConditionalActivityTracker | undefined,
	optionBase: number,
): Map<string, FixedArrayBound> {
	const out = new Map<string, FixedArrayBound>();
	forEachVariableGroup(body as BodyNode[], (group) => {
		if (group.isConst) {
			return;
		}
		for (const decl of group.declarations) {
			if (!decl.isArray || !decl.arrayBounds) {
				continue; // dynamic arrays take their bounds from ReDim or a value
			}
			const lower = decl.name.toLowerCase();
			if (out.has(lower)) {
				continue;
			}
			const dims = parseFixedArrayBoundsForDecl(source, decl, optionBase);
			if (dims) {
				out.set(lower, { name: decl.name, dims, origin: 'Dim' });
			}
		}
	}, activity);
	return out;
}

/** The module's `Option Base`, 0 when absent. */
export function moduleOptionBase(mod: ModuleNode, activity: ConditionalActivityTracker | undefined): number {
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind === 'Option') {
			const match = /^base\s+([01])\b/i.exec(member.optionText.trim());
			if (match) {
				return Number(match[1]);
			}
		}
	}
	return 0;
}

/**
 * Dynamic-array and Variant locals whose bounds a value fixes (issue #120):
 * the local's ONLY assignment is `Array(...)`, `VBA.Array(...)`,
 * `Split(literal, literal[, limit])` or `Range("A1:B2").Value`, and nothing
 * else touches it (no ReDim, Erase, whole pass to a call, or Set).
 *
 *  - `Array(a, b)` is based at Option Base; `VBA.Array` ignores Option Base and
 *    is based at 0 (both measured in Excel 16.0).
 *  - `Array()` has UBound -1: every index is out of range.
 *  - `Split` is always 0-based and yields one part per delimiter plus one;
 *    Split("abc", ",") is one element, Split("a,b", ",") two. Split("") is
 *    empty, UBound -1, like `Array()` (issue #181).
 *  - A Range's `.Value` over a multi-cell address literal is a 1-based
 *    two-dimensional array of the address's rows and columns.
 */
export function knownArrayShapes(
	source: string,
	body: readonly BodyNode[],
	symbols: ReturnType<typeof buildModuleSymbols>,
	proc: ProcedureNode,
	activity: ConditionalActivityTracker | undefined,
	optionBase: number,
): Map<string, FixedArrayBound> {
	const candidates = arrayValueLocals(symbols, proc);
	if (candidates.size === 0) {
		return new Map();
	}
	const assignments = new Map<string, FixedArrayBound[]>();
	const spoiled = new Set<string>();
	const spoil = (lower: string | undefined): void => {
		if (lower && candidates.has(lower)) {
			spoiled.add(lower);
		}
	};
	forEachStatement(body as BodyNode[], (stmt) => {
		for (const span of statementAndBranchSpansOf(stmt)) {
			const toks = statementTokensAfterLeadingLabel(source, span);
			const head = tokenText(toks[0]);
			const bare = bareAssignmentTarget(source, span);
			if (bare) {
				const lower = bare.name.toLowerCase();
				if (!candidates.has(lower)) {
					continue;
				}
				const shape = arrayValueShape(bare.valueTokens, bare.name, optionBase);
				if (!shape) {
					spoil(lower);
				} else {
					const list = assignments.get(lower) ?? [];
					list.push(shape);
					assignments.set(lower, list);
				}
				continue;
			}
			if (head === 'redim' || head === 'erase' || head === 'set' || head === 'input' || head === 'get' || head === 'line') {
				for (const tok of toks) {
					spoil(tokenName(tok)?.toLowerCase());
				}
				continue;
			}
			// Passed whole to a call: `Fill v`, `Fill(v)`, `x = Fill(v)`.
			for (let i = 0; i < toks.length; i++) {
				const name = tokenName(toks[i])?.toLowerCase();
				if (!name || !candidates.has(name)) {
					continue;
				}
				const prev = toks[i - 1];
				const next = toks[i + 1];
				if (next?.rawText === '(') {
					continue; // an index, not a whole pass
				}
				const opensSlot = prev === undefined || prev.rawText === '(' || prev.rawText === ',' || prev.kind === 'identifier' || prev.kind === 'keyword';
				const closesSlot = next === undefined || next.rawText === ')' || next.rawText === ',' || next.rawText === ':' || next.kind === 'comment';
				if (opensSlot && closesSlot && prev?.kind !== 'operator' && next?.kind !== 'operator' && tokenText(prev) !== 'in') {
					spoil(name);
				}
			}
		}
	}, activity);
	const out = new Map<string, FixedArrayBound>();
	for (const [lower, shapes] of assignments) {
		if (!spoiled.has(lower) && shapes.length === 1) {
			out.set(lower, shapes[0]);
		}
	}
	return out;
}

/** The dynamic-array and Variant locals a value can give bounds to, lowercased to declared name. */
function arrayValueLocals(symbols: ReturnType<typeof buildModuleSymbols>, proc: ProcedureNode): Map<string, string> {
	const out = new Map<string, string>();
	for (const child of procedureSymbolFor(symbols, proc)?.children ?? []) {
		if (child.kind !== 'localVariable' || child.visibility === 'Static') {
			continue;
		}
		const type = normalizeType(child.asType);
		if (child.isArray ? child.arrayBounds === undefined : (type === undefined || type === 'variant')) {
			out.set(child.name.toLowerCase(), child.name);
		}
	}
	return out;
}

/**
 * {@link knownArrayShapes} at each statement (issue #180). Where the last
 * assignment to reach a statement in a straight line builds an array, the
 * statement sees its bounds, though the local is assigned again elsewhere:
 * `v = Array(1, 2): Debug.Print v(2): v = Array(1, 2, 3)` reads past the end.
 * Where it reaches with any other value, the local is not known to be an
 * array there.
 */
export function knownArrayShapesAt(
	source: string,
	symbols: ReturnType<typeof buildModuleSymbols>,
	proc: ProcedureNode,
	activity: ConditionalActivityTracker | undefined,
	optionBase: number,
): (stmt: LeafStatementNode) => ReadonlyMap<string, FixedArrayBound> {
	const whole = knownArrayShapes(source, proc.body, symbols, proc, activity, optionBase);
	const locals = arrayValueLocals(symbols, proc);
	const reaching = locals.size === 0 ? new Map() : straightLineAssignments(source, proc.body, activity);
	const results = new Map<ReachingAssignments, ReadonlyMap<string, FixedArrayBound>>();
	return (stmt) => {
		const assignments = reaching.get(stmt);
		if (!assignments) {
			return whole;
		}
		let result = results.get(assignments);
		if (!result) {
			const next = new Map(whole);
			for (const [lower, value] of assignments) {
				if (!locals.has(lower)) {
					continue;
				}
				const shape = arrayValueShape(value, locals.get(lower)!, optionBase);
				if (shape) {
					next.set(lower, shape);
				} else {
					next.delete(lower);
				}
			}
			result = next;
			results.set(assignments, result);
		}
		return result;
	};
}

function statementAndBranchSpansOf(stmt: LeafStatementNode): Span[] {
	const branches = stmt.kind === 'Statement' ? stmt.singleLineIfBranches : undefined;
	return branches ? [stmt.span, ...branches] : [stmt.span];
}

/** The bounds of the array `Array(...)`, `Split(...)` or `Range(...).Value` builds, or undefined. */
function arrayValueShape(valueTokens: readonly VbaToken[], name: string, optionBase: number): FixedArrayBound | undefined {
	const toks = valueTokens.filter((tok) => tok.kind !== 'comment');
	if (toks.length === 0) {
		return undefined;
	}
	let index = 0;
	let vbaQualified = false;
	if (tokenText(toks[0]) === 'vba' && toks[1]?.rawText === '.') {
		vbaQualified = true;
		index = 2;
	}
	const callee = tokenText(toks[index]);
	if ((callee === 'array' || callee === 'split') && toks[index + 1]?.rawText === '(') {
		const close = matchParenFrom(toks, index + 1);
		if (close !== toks.length - 1) {
			return undefined;
		}
		const inner = toks.slice(index + 2, close);
		if (callee === 'array') {
			const count = inner.length === 0 ? 0 : splitTopLevelTokenGroups(inner, ',').length;
			const lower = vbaQualified ? 0 : optionBase;
			return { name, dims: [{ lower, upper: lower + count - 1, explicitLower: true }], origin: vbaQualified ? 'VBA.Array(...)' : 'Array(...)' };
		}
		const args = splitTopLevelTokenGroups(inner, ',');
		if (args.length < 1 || args.length > 2 || args[0].length !== 1 || args[0][0].kind !== 'stringLiteral') {
			return undefined;
		}
		if (args.length === 2 && (args[1].length !== 1 || args[1][0].kind !== 'stringLiteral')) {
			return undefined;
		}
		const text = args[0][0].rawText.slice(1, -1).replace(/""/g, '"');
		const delimiter = args.length === 2 ? args[1][0].rawText.slice(1, -1).replace(/""/g, '"') : ' ';
		if (delimiter.length === 0) {
			return undefined;
		}
		const parts = text.length === 0 ? 0 : text.split(delimiter).length;
		return { name, dims: [{ lower: 0, upper: parts - 1, explicitLower: true }], origin: 'Split(...)' };
	}
	// `Range("A1:B2").Value` and `Worksheets(1).Range("A1:B2").Value`.
	if (tokenText(toks[toks.length - 1]) === 'value' && toks[toks.length - 2]?.rawText === '.' && toks[toks.length - 3]?.rawText === ')') {
		const close = toks.length - 3;
		const open = toks.findIndex((tok, i) => tok.rawText === '(' && matchParenFrom(toks, i) === close);
		if (open > 0 && tokenText(toks[open - 1]) === 'range' && close === open + 2 && toks[open + 1].kind === 'stringLiteral') {
			const address = /^([A-Za-z]{1,3})(\d+):([A-Za-z]{1,3})(\d+)$/.exec(toks[open + 1].rawText.slice(1, -1));
			if (address) {
				const rows = Math.abs(Number(address[4]) - Number(address[2])) + 1;
				const cols = Math.abs(columnNumber(address[3]) - columnNumber(address[1])) + 1;
				if (rows > 1 || cols > 1) {
					return {
						name,
						dims: [{ lower: 1, upper: rows, explicitLower: true }, { lower: 1, upper: cols, explicitLower: true }],
						origin: 'Range(...).Value',
					};
				}
			}
		}
	}
	return undefined;
}

function columnNumber(letters: string): number {
	let n = 0;
	for (const ch of letters.toUpperCase()) {
		n = n * 26 + (ch.charCodeAt(0) - 64);
	}
	return n;
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

/** Whether `value` is outside `dim`, with the words for the message when it is. */
function subscriptDetail(value: number, dim: ArrayDimensionBound, index: number, dims: number): string | undefined {
	if (value >= dim.lower && value <= dim.upper) {
		return undefined;
	}
	const which = dims > 1 ? ` in dimension ${index + 1}` : '';
	if (dim.upper < dim.lower) {
		return `has no element to reach${which}: the array is empty (UBound ${dim.upper})`;
	}
	if (value > dim.upper) {
		return `is above the upper bound ${dim.upper}${which}`;
	}
	return dim.explicitLower
		? `is below the lower bound ${dim.lower}${which}`
		: `is below the lower bound ${dim.lower}${which} (Option Base ${dim.lower})`;
}

/**
 * Literal-subscript accesses of a tracked array that fall outside its bounds,
 * and a loop counter's subscript on its first or last pass (issue #200): a
 * number, or UBound or LBound of the array it indexes, which a bound like
 * `UBound(a) + 1` passes whatever the array holds.
 */
function fixedArraySubscriptViolations(
	source: string,
	span: Span,
	fixed: ReadonlyMap<string, FixedArrayBound>,
	excluded: ReadonlySet<string>,
	counters: CountersAt | undefined,
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
		if (!name || !lower) {
			continue;
		}
		const decl = fixed.has(lower) && !excluded.has(lower) ? fixed.get(lower) : undefined;
		const close = matchParenFrom(toks, i + 1);
		if (close <= i + 1) {
			continue;
		}
		const argToks = toks.slice(i + 2, close).filter((tok) => tok.kind !== 'comment');
		const slots = splitTopLevelTokenGroups(argToks, ',');
		if (slots.some((slot) => slot.length === 0)) {
			continue;
		}
		if (!decl) {
			// UBound(x) or LBound(x) in the loop's bounds says x is an array.
			const hit = symbolicCounterSubscript(span, name, lower, slots, counters);
			if (hit) {
				out.push(hit);
			}
			continue;
		}
		if (slots.length !== decl.dims.length) {
			continue; // the dimension count is the compiler's business, not this rule's
		}
		// One report per access: the first dimension that is out of range.
		for (let index = 0; index < slots.length; index++) {
			const hit = subscriptViolation(span, decl, fixed, slots[index], index, counters);
			if (hit) {
				out.push(hit);
				break;
			}
		}
	}
	return out;
}

function subscriptViolation(
	span: Span,
	decl: FixedArrayBound,
	fixed: ReadonlyMap<string, FixedArrayBound>,
	slot: readonly VbaToken[],
	index: number,
	counters: CountersAt | undefined,
): { span: Span; message: string } | undefined {
	const dim = decl.dims[index];
	const slotSpan = { start: span.start + slot[0].start, end: span.start + slot[slot.length - 1].end };
	const from = decl.origin === 'Dim' ? '' : ` (${decl.origin})`;
	const error = `This will raise Run-time error '9': Subscript out of range.`;
	const value = comparableArrayBoundExpressionValue(slot);
	if (value !== undefined) {
		const detail = subscriptDetail(value, dim, index, decl.dims.length);
		return detail
			? { span: slotSpan, message: `Subscript ${value} for array '${decl.name}'${from} ${detail}. ${error}` }
			: undefined;
	}
	const counter = slot.length === 1 ? counters?.get(tokenName(slot[0])?.toLowerCase() ?? '') : undefined;
	if (!counter) {
		return undefined; // a variable, Const or member chain: not provable
	}
	// `a(i)` inside `For i = 0 To 3`: the counter's first and last passes.
	const atomValue = (atom: { kind: string; name: string; dimension: number }): number | undefined => {
		const shape = fixed.get(atom.name)?.dims[atom.dimension - 1];
		return atom.kind === 'ubound' ? shape?.upper : atom.kind === 'lbound' ? shape?.lower : undefined;
	};
	for (const pass of numericCounterPasses(counter, atomValue)) {
		const detail = subscriptDetail(pass.value, dim, index, decl.dims.length);
		if (detail) {
			const reached = pass.pass === 'first'
				? `Counter '${slot[0].rawText}' is ${pass.value} on its first pass`
				: `Counter '${slot[0].rawText}' reaches ${pass.value} on its last pass`;
			return { span: slotSpan, message: `${reached}, which for array '${decl.name}'${from} ${detail}. ${error}` };
		}
	}
	return symbolicCounterSubscript(span, decl.name, decl.name.toLowerCase(), [slot], counters, index);
}

/**
 * `a(i)` where the counter runs past UBound(a) or below LBound(a) of the
 * same dimension: `For i = 1 To UBound(a) + 1`.
 */
function symbolicCounterSubscript(
	span: Span,
	name: string,
	lower: string,
	slots: readonly (readonly VbaToken[])[],
	counters: CountersAt | undefined,
	onlyIndex?: number,
): { span: Span; message: string } | undefined {
	for (let index = 0; index < slots.length; index++) {
		if (onlyIndex !== undefined && index !== 0) {
			break;
		}
		const slot = slots[index];
		const dimension = (onlyIndex ?? index) + 1;
		const counter = slot.length === 1 ? counters?.get(tokenName(slot[0])?.toLowerCase() ?? '') : undefined;
		if (!counter) {
			continue;
		}
		const passes: Array<['first' | 'last', CounterValue | undefined]> = [['first', counter.first], ['last', counter.last]];
		for (const [pass, value] of passes) {
			const atom = value?.atom;
			if (!value || !atom || atom.name !== lower || atom.dimension !== dimension) {
				continue;
			}
			const past = atom.kind === 'ubound' && value.offset > 0
				? 'above its upper bound'
				: atom.kind === 'lbound' && value.offset < 0 ? 'below its lower bound' : undefined;
			if (past) {
				const reached = pass === 'first' ? `is ${counterText(value)} on its first pass` : `reaches ${counterText(value)} on its last pass`;
				return {
					span: { start: span.start + slot[0].start, end: span.start + slot[slot.length - 1].end },
					message: `Counter '${slot[0].rawText}' ${reached}, which for array '${name}' is ${past}. This will raise Run-time error '9': Subscript out of range.`,
				};
			}
		}
	}
	return undefined;
}

/**
 * `UBound(a, 2)` / `LBound(a, 2)` on an array with fewer dimensions raises 9
 * (issue #120, measured in Excel 16.0).
 */
function boundIntrinsicDimensionViolations(
	source: string,
	span: Span,
	fixed: ReadonlyMap<string, FixedArrayBound>,
	excluded: ReadonlySet<string>,
): Array<{ span: Span; message: string }> {
	const toks = statementTokensAfterLeadingLabel(source, span);
	const out: Array<{ span: Span; message: string }> = [];
	for (let i = 0; i + 1 < toks.length; i++) {
		const callee = tokenText(toks[i]);
		if ((callee !== 'ubound' && callee !== 'lbound') || toks[i + 1].rawText !== '(' || !isBareOrVbaQualifiedIntrinsicCall(toks, i)) {
			continue;
		}
		const close = matchParenFrom(toks, i + 1);
		if (close < 0) {
			continue;
		}
		const args = splitTopLevelTokenGroups(toks.slice(i + 2, close).filter((tok) => tok.kind !== 'comment'), ',');
		if (args.length !== 2 || args[0].length !== 1) {
			continue;
		}
		const lower = tokenName(args[0][0])?.toLowerCase();
		const decl = lower ? fixed.get(lower) : undefined;
		if (!decl || !lower || excluded.has(lower)) {
			continue;
		}
		const dimension = comparableArrayBoundExpressionValue(args[1]);
		if (dimension === undefined || (dimension >= 1 && dimension <= decl.dims.length)) {
			continue;
		}
		out.push({
			span: { start: span.start + args[1][0].start, end: span.start + args[1][args[1].length - 1].end },
			message: `${toks[i].rawText} asks for dimension ${dimension} of '${decl.name}', which has ${pluralizeCount(decl.dims.length, 'dimension')}. This will raise Run-time error '9': Subscript out of range.`,
		});
	}
	return out;
}

/**
 * `Split("abc", ",")(1)`: indexing the result of Split on literals, whose one
 * element sits at 0 (issue #120).
 */
function inlineSplitIndexViolations(source: string, span: Span): Array<{ span: Span; message: string }> {
	const toks = statementTokensAfterLeadingLabel(source, span);
	const out: Array<{ span: Span; message: string }> = [];
	for (let i = 0; i + 1 < toks.length; i++) {
		if (tokenText(toks[i]) !== 'split' || toks[i + 1].rawText !== '(' || !isBareOrVbaQualifiedIntrinsicCall(toks, i)) {
			continue;
		}
		const close = matchParenFrom(toks, i + 1);
		if (close < 0 || toks[close + 1]?.rawText !== '(') {
			continue;
		}
		const indexClose = matchParenFrom(toks, close + 1);
		if (indexClose < 0) {
			continue;
		}
		const shape = arrayValueShape(toks.slice(i, close + 1), 'Split(...)', 0);
		const indexToks = toks.slice(close + 2, indexClose).filter((tok) => tok.kind !== 'comment');
		const value = comparableArrayBoundExpressionValue(indexToks);
		if (!shape || value === undefined) {
			continue;
		}
		const detail = subscriptDetail(value, shape.dims[0], 0, 1);
		if (detail) {
			out.push({
				span: { start: span.start + indexToks[0].start, end: span.start + indexToks[indexToks.length - 1].end },
				message: `Subscript ${value} for the array Split returns here ${detail}. This will raise Run-time error '9': Subscript out of range.`,
			});
		}
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
	symbols: ReturnType<typeof buildModuleSymbols>,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	const optionBase = moduleOptionBase(mod, activity);
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		const declared = localFixedArrayDeclarationsForBody(source, member.body, activity, optionBase);
		const shapesAt = knownArrayShapesAt(source, symbols, member, activity, optionBase);
		const merged = new Map<ReadonlyMap<string, FixedArrayBound>, ReadonlyMap<string, FixedArrayBound>>();
		const fixedAt = (stmt: LeafStatementNode): ReadonlyMap<string, FixedArrayBound> => {
			const shapes = shapesAt(stmt);
			let fixed = merged.get(shapes);
			if (!fixed) {
				const next = new Map(declared);
				for (const [lower, shape] of shapes) {
					if (!declared.has(lower)) {
						next.set(lower, shape);
					}
				}
				fixed = next;
				merged.set(shapes, fixed);
			}
			return fixed;
		};
		const excluded = redimTargetNamesInBody(source, member.body, activity);
		const counters = loopCountersAt(source, member.body, activity);
		// Headers too: `For i = 1 To a(5)`, `Select Case a(5)` (issue #233).
		forEachStatementWithHeaders(source, member.body, (stmt) => {
			for (const hit of inlineSplitIndexViolations(source, stmt.span)) {
				push('arraySubscriptOutOfBounds', hit.message, hit.span);
			}
			const fixed = fixedAt(stmt);
			const stmtCounters = counters.get(stmt);
			if (fixed.size === 0 && !stmtCounters) {
				return;
			}
			for (const hit of fixedArraySubscriptViolations(source, stmt.span, fixed, excluded, stmtCounters)) {
				push('arraySubscriptOutOfBounds', hit.message, hit.span);
			}
			for (const hit of boundIntrinsicDimensionViolations(source, stmt.span, fixed, excluded)) {
				push('arraySubscriptOutOfBounds', hit.message, hit.span);
			}
		}, activity);
	}
}
