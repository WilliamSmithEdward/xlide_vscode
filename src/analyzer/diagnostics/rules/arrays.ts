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
import { walkBranchMergedBody, walkEnteringBlocks, walkStraightLineBody } from '../dataflow';
import { isLeafStatement } from '../../parser/nodes';
import { jumpTargetLabelDeclaration } from '../../flow/procedureLabels';
import { straightLineAssignments, type ReachingAssignments } from '../straightLineValues';
import { counterText, loopCountersAt, numericCounterPasses, type CounterValue, type CountersAt } from '../loopCounters';
import { procedureHasUnstructuredFlow } from '../../flow/procedureUnstructured';
import { untouchedModuleVariablesIn } from '../moduleState';
import { isBareOrVbaQualifiedIntrinsicCall, sourceExpressionSyntaxProblem } from '../rules/shared';
import { foldStringExpression, moduleCompare, type ModuleCompare } from '../knownStringCalls';
import {
	declarationShapeEnvironmentFor,
	declaredShapeForSourceBinding,
	type DeclaredValueShape,
	isKnownScalarType,
	normalizeType,
	knownLocalLiteralValuesAt,
	procedureIntegerConstantLookup,
	scopedIntegerConstantLookup,
	sourceIdentifierBinding,
	stringConstantsInScope,
	stringLiteralValue,
	unreachableStatementsIn,
	withKnownLocals,
	type SourceDeclaredShape,
	runtimeCallableSourceShadowed,
	sourceNameScopeFor,
} from '../typeInference';
import {
	absoluteSpan,
	activeModuleMembers,
	bareAssignmentTarget,
	blockHeaderStatements,
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
			for (const hit of scalarLocalsIndexed(source, stmt.span, member, shapes)) {
				push('scalarIndexed', `'${hit.name}' is declared As ${hit.asType}, which is no array to index. This is a VBE compile error: Expected array.`, hit.span);
			}
			for (const hit of arrayBoundSyntaxProblems(source, stmt.span)) {
				push('malformedStatement', hit.message, hit.span);
			}
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

/**
 * `Dim f As Long: Main = f(1)`: a number, string or date variable given
 * a subscript is the VBE compile error Expected array (issue #417,
 * Excel 16.0). The procedure's own name is a call, and declarations are
 * left to the declaration rules.
 */
function scalarLocalsIndexed(
	source: string,
	span: Span,
	proc: ProcedureNode,
	shapes: ReadonlyMap<string, DeclaredValueShape>,
): Array<{ name: string; asType: string; span: Span }> {
	const toks = statementTokensAfterLeadingLabel(source, span).filter((tok) => tok.kind !== 'comment');
	const lead = tokenText(toks[0]);
	if (['dim', 'redim', 'static', 'const', 'private', 'public', 'global', 'erase'].includes(lead)) {
		return [];
	}
	const out: Array<{ name: string; asType: string; span: Span }> = [];
	for (let i = 0; i < toks.length - 1; i++) {
		const name = tokenName(toks[i]);
		if (!name || toks[i + 1].rawText !== '(' || toks[i - 1]?.rawText === '.' || toks[i - 1]?.rawText === '!'
			|| name.toLowerCase() === proc.name.toLowerCase()) {
			continue;
		}
		const shape = shapes.get(name.toLowerCase());
		const normalized = shape && !shape.isArray && shape.asType ? normalizeType(shape.asType) : undefined;
		if (!normalized || !isKnownScalarType(normalized)) {
			continue;
		}
		out.push({ name, asType: shape!.asType!, span: { start: span.start + toks[i].start, end: span.start + toks[i].end } });
	}
	return out;
}

/**
 * `UBound(5)`, `LBound("abc")`, `UBound(-5)` and `UBound(5 + 1)`: what
 * UBound and LBound read is an array variable, a member or a call, and
 * anything else is a Syntax error (measured in Excel 16.0). `VBA.UBound`
 * is left alone: UBound is no member of VBA, which the VBE reports first.
 */
function arrayBoundSyntaxProblems(source: string, span: Span): Array<{ span: Span; message: string }> {
	const toks = statementTokens(source, span);
	const out: Array<{ span: Span; message: string }> = [];
	for (let i = 0; i < toks.length - 2; i++) {
		const word = tokenText(toks[i]);
		if ((word !== 'ubound' && word !== 'lbound') || toks[i + 1].rawText !== '(' || toks[i - 1]?.rawText === '.') {
			continue;
		}
		const close = matchParenFrom(toks, i + 1);
		const split = close < 0 ? undefined : splitArgSlots(toks.slice(i + 2, close), span.start);
		const first = split?.slots[0]?.filter((tok) => tok.kind !== 'comment' && tok.kind !== 'newline') ?? [];
		// `UBound((a))` is reported as an array in parentheses already.
		if (first.length === 0 || first[0].rawText === '(') {
			continue;
		}
		const at = { start: span.start + first[0].start, end: span.start + first[first.length - 1].end };
		const what = sourceExpressionSyntaxProblem(source.slice(at.start, at.end));
		if (what) {
			out.push({
				span: at,
				message: `${toks[i].rawText} takes an array variable, a member or a call, and ${what} is none of them. This is a VBE compile error: Syntax error.`,
			});
		}
	}
	return out;
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
		// Only a Variant can hold an array: a Collection, Object or Type
		// is refused as well (issue #417, Excel 16.0).
		const normalized = normalizeType(shape.asType);
		if (!normalized || normalized === 'variant' || normalized === 'any') {
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
	/** Whether the bound writes a lower, `1 To n`; without one it takes the Option Base. */
	lowerWritten?: boolean;
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
				// A Const or Enum member is no array (issue #255): "Expected array".
				const binding = sourceIdentifierBinding(symbols, procSym, projectVisibleSymbols, target.name, 'expression');
				if (binding.scope !== 'ambiguous' && binding.definitions.length > 0 && binding.definitions.every((definition) => definition.kind === 'constant' || definition.kind === 'enumMember')) {
					push('scalarRedim', `'${target.name}' is a constant, which ReDim cannot resize. This is a VBE compile error: Expected array.`, target.span);
					continue;
				}
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
					lowerWritten: dimTokens.some((tok) => tokenText(tok) === 'to'),
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
		const constantsLookup = (): IntegerConstantLookup => constants ??= procedureIntegerConstantLookup(
			member, moduleConstants, symbols, projectVisibleSymbols, activity, hostModel,
		);
		const valuesAt = knownLocalLiteralValuesAt(source, member, symbols, activity);
		return (stmt) => {
			// `zz = -1` then `ReDim a(zz)` (issue #238).
			const lookup = (): IntegerConstantLookup => withKnownLocals(constantsLookup(), valuesAt(stmt));
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
		// A decimal, signed too: `ReDim d(-0.6)` rounds to -1 (issue #286).
		const negative = part.length === 2 && part[0].rawText === '-';
		const decimal = part[negative ? 1 : 0];
		if (part.length === (negative ? 2 : 1) && decimal.kind === 'floatLiteral') {
			const value = Number(decimal.rawText.replace(/[!#@]$/, '').replace(/[dD]/, 'e'));
			return Number.isFinite(value) ? bankersRound(negative ? -value : value) + 0 : undefined;
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
		checkRedimPreserveDimensionsInBody(source, member.body, new Map(), activity, push, moduleOptionBase(mod, activity));
	}
}

function checkRedimPreserveDimensionsInBody(
	source: string,
	body: BodyNode[],
	initialShapes: ReadonlyMap<string, RedimTarget>,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
	base: number,
): void {
	const shapes = new Map(initialShapes);
	for (const node of body) {
		if (isInactiveNode(activity, node)) {
			continue;
		}
		if (node.kind === 'Statement') {
			// After Erase a dynamic array has no bounds, and ReDim Preserve sets
			// them afresh, the first dimension included (issue #420).
			const toks = statementTokensAfterLeadingLabel(source, node.span);
			if (tokenText(toks[0]) === 'erase') {
				for (const tok of toks.slice(1)) {
					const erased = tokenName(tok);
					if (erased) {
						shapes.delete(erased.toLowerCase());
					}
				}
				continue;
			}
			for (const target of redimStatementTargets(source, node.span)) {
				if (target.preserve) {
					const previous = shapes.get(target.name.toLowerCase());
					const reason = previous
						? redimPreserveDimensionMismatch(previous, target, base)
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
				base,
			);
		}
	}
}

function redimPreserveDimensionMismatch(
	previous: RedimTarget,
	current: RedimTarget,
	base: number,
): string | undefined {
	// A bound written with no lower takes the Option Base: after
	// `ReDim a(1 To 3)`, `ReDim Preserve a(UBound(a) + 1)` moves the lower
	// bound to 0, which raises 9 (issue #342, measured in Excel 16.0).
	const movesLower = (before: RedimDimension | undefined, after: RedimDimension | undefined): boolean =>
		before !== undefined && after !== undefined
		&& ((before.lowerWritten === true && after.lowerWritten === false && before.lowerValue !== undefined && before.lowerValue !== base)
			|| (before.lowerWritten === false && after.lowerWritten === true && after.lowerValue !== undefined && after.lowerValue !== base));
	if (previous.dimensions.length === current.dimensions.length) {
		for (let i = 0; i < current.dimensions.length; i++) {
			if (movesLower(previous.dimensions[i], current.dimensions[i])) {
				const lower = previous.dimensions[i].lowerValue ?? base;
				const now = current.dimensions[i].lowerValue ?? base;
				return `The lower bound of dimension ${i + 1} changes under Preserve, from ${lower} to ${now}: a bound written without one takes the Option Base.`;
			}
		}
	}
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
	/** A Variant local that takes a copy of a dynamic array (issue #342). */
	variant?: boolean;
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
	symbols: ReturnType<typeof buildModuleSymbols>,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	const unsetFunctions = arrayFunctionsNeverSet(source, mod, activity);
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		if (unsetFunctions.size > 0) {
			checkUnsetArrayResults(source, member, symbols, unsetFunctions, activity, push);
		}
		const arrays = localDynamicArrayDeclarationsForBody(member.body, activity);
		const state = new Map<string, DynamicArrayAllocationState>();
		for (const lower of arrays.keys()) {
			state.set(lower, 'unallocated');
		}
		// A Variant that takes a copy of one: `v = a` with a unallocated
		// leaves v an array with no storage (issue #342, measured in Excel 16.0).
		// So does one given Array or Split and then erased (issue #420).
		for (const [lower, decl] of variantArrayCopies(source, member.body, arrays, activity)) {
			arrays.set(lower, decl);
			state.set(lower, 'unknown');
		}
		if (arrays.size === 0) {
			continue;
		}
		// The GoTo-following walk runs the body until its labels settle,
		// and reports on its last run (issue #271).
		let silent = false;
		const report: PushFn = (...finding) => {
			if (!silent) {
				push(...finding);
			}
		};
		const walk = procedureHasUnstructuredFlow(source, member, activity)
			? walkStraightLineBody
			: walkBranchMergedBody;
		// A statement a known guard keeps from running (issue #273).
		const unreachable = unreachableStatementsIn(source, member, symbols, activity);
		walk(source, member.body, (node) => isInactiveNode(activity, node) || unreachable.has(node), {
			onStatement: (stmt) =>
				checkUnallocatedDynamicArrayAccessStatement(source, stmt, arrays, state, report),
			onBlock: (node) => {
				// The header runs as the block is entered: `For i = 0 To
				// UBound(a)`, `Do While i <= UBound(a)`, `Select Case UBound(a)`
				// (issue #342, measured in Excel 16.0).
				if (node.kind === 'SelectBlock' || node.kind === 'DoBlock' || node.kind === 'WhileBlock' || (node.kind === 'ForBlock' && !node.each)) {
					const { before } = blockHeaderStatements(source, node);
					if (before) {
						checkUnallocatedDynamicArrayAccessStatement(source, before, arrays, state, report);
					}
				}
				// `For Each x In a` over an array with no storage raises 92, For
				// loop not initialized, not 9 (issue #181, measured in Excel 16.0).
				const over = node.kind === 'ForBlock' && node.each ? node.sourceExpression?.trim().toLowerCase() : undefined;
				if (over && node.kind === 'ForBlock' && node.sourceExpressionSpan && state.get(over) === 'unallocated') {
					report(
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
			setSilent: (quiet) => {
				silent = quiet;
			},
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
				// An erased Variant may hold a fixed array, which Erase clears and
				// keeps; one known to hold a dynamic array, from Array, Split, ReDim
				// or a copy, is left with none (issue #420, measured in Excel 16.0).
				const variant = arrays.get(lower)!.variant;
				state.set(lower, variant && state.get(lower) !== 'allocated' ? 'unknown' : 'unallocated');
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
		// `a = b` copies b's storage, or its lack of it (issue #342).
		const value = assignment!.valueTokens.filter((tok) => tok.kind !== 'comment');
		const from = value.length === 1 ? tokenName(value[0])?.toLowerCase() : undefined;
		state.set(assignmentLower, from && arrays.has(from) ? state.get(from) ?? 'unknown' : dynamicArrayCall(value) ? 'allocated' : 'unknown');
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
	// `If L > 0 Then tb = txt` gives the array storage on one path only.
	for (const branch of statementAndBranchSpansOf(stmt).slice(1)) {
		const lower = bareAssignmentTarget(source, branch)?.name.toLowerCase();
		if (lower && arrays.has(lower) && state.get(lower) === 'unallocated') {
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

/** The Variant locals some statement assigns one of the dynamic arrays whole: `v = a`. */
function variantArrayCopies(
	source: string,
	body: readonly BodyNode[],
	arrays: ReadonlyMap<string, DynamicArrayDeclaration>,
	activity: ConditionalActivityTracker | undefined,
): Map<string, DynamicArrayDeclaration> {
	const variants = new Map<string, DynamicArrayDeclaration>();
	forEachVariableGroup(body as BodyNode[], (group) => {
		if (group.isConst || group.modifier === 'Static') {
			return;
		}
		for (const decl of group.declarations) {
			const type = normalizeType(decl.asType);
			if (!decl.isArray && !decl.typeSuffix && (type === undefined || type === 'variant')) {
				variants.set(decl.name.toLowerCase(), { name: decl.name, span: decl.span, variant: true });
			}
		}
	}, activity);
	const out = new Map<string, DynamicArrayDeclaration>();
	if (variants.size === 0) {
		return out;
	}
	forEachStatement(body as BodyNode[], (stmt) => {
		const bare = bareAssignmentTarget(source, stmt.span);
		const value = bare?.valueTokens.filter((tok) => tok.kind !== 'comment') ?? [];
		const lower = bare?.name.toLowerCase() ?? '';
		if (variants.has(lower) && ((value.length === 1 && arrays.has(tokenName(value[0])?.toLowerCase() ?? '')) || dynamicArrayCall(value))) {
			out.set(lower, variants.get(lower)!);
		}
	}, activity);
	return out;
}

/** `Array(1, 2)` or `Split(s, ",")`, whole: a dynamic array (issue #420). */
function dynamicArrayCall(value: readonly VbaToken[]): boolean {
	const at = tokenText(value[0]) === 'vba' && value[1]?.rawText === '.' ? 2 : 0;
	const name = tokenText(value[at]);
	return (name === 'array' || name === 'split') && value[at + 1]?.rawText === '(' && matchParenFrom(value, at + 1) === value.length - 1;
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

/**
 * The module's Functions declared to return an array, `As Long()`, whose
 * body never names the result: each returns an array with no storage, so
 * `UBound(F())` raises 9 (issue #448, measured in Excel 16.0).
 */
function arrayFunctionsNeverSet(source: string, mod: ModuleNode, activity: ConditionalActivityTracker | undefined): Map<string, ProcedureNode> {
	const out = new Map<string, ProcedureNode>();
	const seen = new Set<string>();
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		const lower = member.name.toLowerCase();
		if (seen.has(lower)) {
			out.delete(lower);
			continue;
		}
		seen.add(lower);
		if (member.procKind !== 'Function' || !/\(\s*\)\s*$/.test(member.returnType ?? '')) {
			continue;
		}
		// The header names it once; any other mention may set it.
		const named = statementTokens(source, member.span).filter((tok, i, toks) => tokenName(tok)?.toLowerCase() === lower && toks[i - 1]?.rawText !== '.').length;
		if (named === 1) {
			out.set(lower, member);
		}
	}
	return out;
}

/** `UBound(F())` or `LBound(F)` on a Function that never sets its array result. */
function checkUnsetArrayResults(
	source: string,
	member: ProcedureNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
	functions: ReadonlyMap<string, ProcedureNode>,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	const own = new Set([member.name.toLowerCase(), ...member.params.map((param) => param.name.toLowerCase()), ...(procedureSymbolFor(symbols, member)?.children ?? []).map((child) => child.name.toLowerCase())]);
	forEachStatement(member.body, (stmt) => {
		for (const span of statementAndBranchSpansOf(stmt)) {
			const toks = statementTokens(source, span);
			for (let i = 0; i < toks.length - 2; i++) {
				const bound = tokenText(toks[i]);
				if ((bound !== 'lbound' && bound !== 'ubound') || toks[i + 1].rawText !== '(' || !isBareOrVbaQualifiedIntrinsicCall(toks, i)) {
					continue;
				}
				const lower = tokenName(toks[i + 2])?.toLowerCase() ?? '';
				const fn = own.has(lower) ? undefined : functions.get(lower);
				const end = toks[i + 3]?.rawText === '(' && toks[i + 4]?.rawText === ')' ? i + 4 : i + 2;
				if (!fn || (end === i + 2 && fn.params.length > 0) || (toks[end + 1]?.rawText !== ')' && toks[end + 1]?.rawText !== ',')) {
					continue;
				}
				push(
					'unallocatedDynamicArrayAccess',
					`Function '${fn.name}' never sets its result, so it returns an array with no storage, and ${toks[i].rawText} has no bounds to read. This will raise Run-time error '9': Subscript out of range.`,
					{ start: span.start + toks[i + 2].start, end: span.start + toks[end].end },
				);
			}
		}
	}, activity);
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
	for (const branch of statementAndBranchSpansOf(stmt).slice(1)) {
		const lower = bareAssignmentTarget(source, branch)?.name.toLowerCase();
		if (lower && arrays.has(lower)) {
			out.add(lower);
		}
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
	/** The arrays its elements hold, by position, where Array(...) of Array(...) built it. */
	elements?: ReadonlyArray<FixedArrayBound | undefined>;
	/** The literal each element holds, by position, where it is known (issue #260). */
	values?: ReadonlyArray<string | number | undefined>;
}

/** The String a name holds at the statement being read, or undefined (issue #260). */
export type StringValueOf = (tok: VbaToken) => string | undefined;

/** A subscript the array cannot take; a count the compiler refuses has its own rule. */
export interface SubscriptHit {
	span: Span;
	message: string;
	rule?: 'wrongNumberOfDimensions';
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
export function parseFixedArrayBoundsForDecl(
	source: string,
	decl: { span: Span },
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
	return literalDimensions(toks.slice(open + 1, close), optionBase);
}

/** The bounds a parenthesised bounds list states, `2, 1 To 3`, when every one is a literal. */
export function literalDimensions(inner: readonly VbaToken[], optionBase: number): ArrayDimensionBound[] | undefined {
	const dims = splitTopLevelTokenGroups(inner, ',')
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

/** A procedure's local fixed arrays whose bounds are literals, by lowercased name. */
export function localFixedArrays(
	source: string,
	proc: ProcedureNode,
	activity: ConditionalActivityTracker | undefined,
	optionBase: number,
): Map<string, FixedArrayBound> {
	return localFixedArrayDeclarationsForBody(source, proc.body, activity, optionBase);
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
	const stringsAt = stringValuesAt(source, symbols, proc, activity);
	const compare = moduleCompare(source);
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
				const shape = arrayValueShape(bare.valueTokens, bare.name, optionBase, stringsAt(stmt), compare);
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

/**
 * What a String local or a String Const holds as each statement starts, for
 * the text and delimiter of `Split(s, d)` and the match of `Filter` (issue
 * #260). A local is read where the array is built, not where it is indexed.
 */
function stringValuesAt(
	source: string,
	symbols: ReturnType<typeof buildModuleSymbols>,
	proc: ProcedureNode,
	activity: ConditionalActivityTracker | undefined,
): (stmt: BodyNode) => StringValueOf {
	let valuesAt: ReturnType<typeof knownLocalLiteralValuesAt> | undefined;
	let consts: ReadonlyMap<string, string> | undefined;
	return (stmt) => (tok) => {
		const lower = tokenName(tok)?.toLowerCase();
		if (!lower) {
			return undefined;
		}
		valuesAt ??= knownLocalLiteralValuesAt(source, proc, symbols, activity);
		const local = valuesAt(stmt).get(lower);
		if (local) {
			return local.kind === 'string' && !local.contentMutated ? (local.value as string) : undefined;
		}
		return (consts ??= stringConstantsInScope(symbols, proc)).get(lower);
	};
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
	const compare = moduleCompare(source);
	const locals = arrayValueLocals(symbols, proc);
	const reaching = locals.size === 0 ? new Map() : straightLineAssignments(source, proc.body, activity);
	// Each assignment's shape, built with the Strings its own statement sees
	// and keyed by its first value token, which the reaching value shares.
	const built = new Map<VbaToken, FixedArrayBound | undefined>();
	if (locals.size > 0) {
		const stringsAt = stringValuesAt(source, symbols, proc, activity);
		forEachStatement(proc.body as BodyNode[], (stmt) => {
			const bare = bareAssignmentTarget(source, stmt.span);
			const first = bare?.valueTokens.find((tok) => tok.kind !== 'comment');
			if (bare && first && locals.has(bare.name.toLowerCase())) {
				built.set(first, arrayValueShape(bare.valueTokens, locals.get(bare.name.toLowerCase())!, optionBase, stringsAt(stmt), compare));
			}
		}, activity);
	}
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
				const shape = built.has(value[0]) ? built.get(value[0]) : arrayValueShape(value, locals.get(lower)!, optionBase, undefined, compare);
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

/**
 * The bounds the last `ReDim` gave a dynamic array or Variant local, at each
 * statement it reaches (issue #238): `ReDim a(3)` then `a(7)` is error 9.
 * A ReDim whose bounds are not literals, any other whole mention of the name
 * (`Erase a`, `a = b`, `Fill a`), a ReDim inside a single-line If, and a
 * label or GoSub end what is known. Blocks are entered as issue #237 enters
 * them.
 */
export function redimShapesAt(
	source: string,
	symbols: ReturnType<typeof buildModuleSymbols>,
	proc: ProcedureNode,
	activity: ConditionalActivityTracker | undefined,
	optionBase: number,
): Map<LeafStatementNode, ReadonlyMap<string, FixedArrayBound>> {
	const out = new Map<LeafStatementNode, ReadonlyMap<string, FixedArrayBound>>();
	const locals = arrayValueLocals(symbols, proc);
	if (locals.size === 0) {
		return out;
	}
	let shapes = new Map<string, FixedArrayBound>();
	let seen: ReadonlyMap<string, FixedArrayBound> = shapes;
	const changed = (): void => {
		shapes = new Map(shapes);
		seen = shapes;
	};
	const forget = (names: Iterable<string>): void => {
		for (const lower of names) {
			if (shapes.has(lower)) {
				changed();
				shapes.delete(lower);
			}
		}
	};
	const visit = (node: BodyNode): void => {
		if (!isLeafStatement(node)) {
			return;
		}
		if (jumpTargetLabelDeclaration(source, node.span)) {
			forget([...shapes.keys()]);
		}
		const toks = statementTokensAfterLeadingLabel(source, node.span);
		// A ReDim's bounds are not subscripts: `ReDim Preserve a(5)` reads no a(5).
		if (shapes.size > 0 && !toks.some((tok) => tokenText(tok) === 'redim')) {
			out.set(node, seen);
		}
		if (tokenText(toks[0]) === 'gosub') {
			forget([...shapes.keys()]);
			return;
		}
		// A ReDim a single-line If runs may not run: it only ends what is known.
		const conditional = node.singleLineIfTail === true || (node.kind === 'Statement' && node.singleLineIfBranches !== undefined);
		const redims = conditional ? [] : redimStatementTargets(source, node.span);
		const reshaped = new Set(redims.map((target) => target.name.toLowerCase()));
		forget([...shapeTouches(source, node)].filter((lower) => !reshaped.has(lower)));
		for (const target of redims) {
			const lower = target.name.toLowerCase();
			const name = locals.get(lower);
			const dims = target.dimensions.map((dim): ArrayDimensionBound | undefined => (
				dim.upperValue === undefined || (dim.lowerKey !== undefined && dim.lowerValue === undefined)
					? undefined
					: { lower: dim.lowerValue ?? optionBase, upper: dim.upperValue, explicitLower: dim.lowerValue !== undefined }
			));
			changed();
			if (name && dims.length > 0 && dims.every((dim) => dim !== undefined)) {
				shapes.set(lower, { name, dims: dims as ArrayDimensionBound[], origin: 'ReDim' });
			} else {
				shapes.delete(lower);
			}
		}
	};
	walkEnteringBlocks(source, proc.body, (node) => isInactiveNode(activity, node), visit, {
		snapshot: () => shapes,
		restore: (saved) => {
			shapes = saved;
			seen = saved;
		},
		forget,
		touches: (stmt) => shapeTouches(source, stmt),
	});
	return out;
}

/**
 * The names a statement may reshape: every name in a ReDim, and a name used
 * whole rather than indexed, as `Erase a` and `Fill a` use it. `a(1) = 2`
 * leaves a's bounds alone.
 */
function shapeTouches(source: string, stmt: LeafStatementNode): Set<string> {
	const toks = statementTokensAfterLeadingLabel(source, stmt.span);
	const redim = toks.some((tok) => tokenText(tok) === 'redim');
	const out = new Set<string>();
	for (let i = 0; i < toks.length; i++) {
		const lower = tokenName(toks[i])?.toLowerCase();
		if (!lower || toks[i - 1]?.rawText === '.' || toks[i - 1]?.rawText === '!') {
			continue;
		}
		if (redim || toks[i + 1]?.rawText !== '(') {
			out.add(lower);
		}
	}
	return out;
}

function statementAndBranchSpansOf(stmt: LeafStatementNode): Span[] {
	const branches = stmt.kind === 'Statement' ? stmt.singleLineIfBranches : undefined;
	return branches ? [stmt.span, ...branches] : [stmt.span];
}

/**
 * The bounds of the array `Array(...)`, `Split(...)`, `Filter(...)` or
 * `Range(...).Value` builds, or undefined. `strings` gives the String a name
 * holds where the value is read, so `Split(s, ",")` is known when s is
 * (issue #260).
 */
export function arrayValueShape(valueTokens: readonly VbaToken[], name: string, optionBase: number, strings?: StringValueOf, compare?: ModuleCompare): FixedArrayBound | undefined {
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
	if ((callee === 'array' || callee === 'split' || callee === 'filter') && toks[index + 1]?.rawText === '(') {
		const close = matchParenFrom(toks, index + 1);
		if (close !== toks.length - 1) {
			return undefined;
		}
		const inner = toks.slice(index + 2, close);
		if (callee === 'array') {
			const groups = inner.length === 0 ? [] : splitTopLevelTokenGroups(inner, ',');
			const lower = vbaQualified ? 0 : optionBase;
			const elements = groups.map((group) => arrayValueShape(group, name, optionBase, strings, compare));
			const values = groups.map((group) => literalElementValue(group));
			return {
				name,
				dims: [{ lower, upper: lower + groups.length - 1, explicitLower: true }],
				origin: vbaQualified ? 'VBA.Array(...)' : 'Array(...)',
				...(elements.some((element) => element !== undefined) ? { elements } : {}),
				...(values.some((value) => value !== undefined) ? { values } : {}),
			};
		}
		const args = splitTopLevelTokenGroups(inner, ',');
		if (callee === 'filter') {
			return filterShape(args, name, optionBase, strings, compare);
		}
		if (args.length < 1 || args.length > 4) {
			return undefined;
		}
		const text = stringArgument(args[0], strings);
		const delimiter = args.length >= 2 && args[1].length > 0 ? stringArgument(args[1], strings) : ' ';
		// Limit -1 keeps every part, 0 none, n at most n; any other
		// negative is error 5, which this leaves alone.
		const limit = args.length >= 3 && args[2].length > 0 ? signedIntegerArgument(args[2]) : -1;
		const textCompare = args.length === 4 ? compareArgument(args[3]) : defaultCompare(compare, delimiter);
		if (text === undefined || delimiter === undefined || delimiter.length === 0 || limit === undefined || limit < -1 || textCompare === undefined) {
			return undefined;
		}
		const values = text.length === 0 || limit === 0 ? [] : splitParts(text, delimiter, limit, textCompare);
		if (!values) {
			return undefined;
		}
		return {
			name,
			dims: [{ lower: 0, upper: values.length - 1, explicitLower: true }],
			origin: 'Split(...)',
			...(values.length > 0 ? { values } : {}),
		};
	}
	// `Range("A1:B2").Value` and `Worksheets(1).Range("A1:B2").Value`; Value2 too (issue #278).
	const block = rangeValueBlock(toks);
	if (block && (block.rows > 1 || block.cols > 1)) {
		return {
			name,
			dims: [{ lower: 1, upper: block.rows, explicitLower: true }, { lower: 1, upper: block.cols, explicitLower: true }],
			origin: 'Range(...).Value',
		};
	}
	// `Application.Transpose(Range("A1:A3").Value)`: one column becomes a
	// 1-D array from 1, a row or a block a 2-D one turned over (issue #278,
	// measured in Excel 16.0).
	const transposed = transposeArgument(toks);
	const inner = transposed ? rangeValueBlock(transposed) : undefined;
	if (inner && (inner.rows > 1 || inner.cols > 1)) {
		return {
			name,
			dims: inner.cols === 1
				? [{ lower: 1, upper: inner.rows, explicitLower: true }]
				: [{ lower: 1, upper: inner.cols, explicitLower: true }, { lower: 1, upper: inner.rows, explicitLower: true }],
			origin: 'Transpose(...)',
		};
	}
	return undefined;
}

/** The rows and columns of a literal `[...]Range("A1:B2").Value` or `.Value2`, the whole of `toks`. */
export function rangeValueBlock(toks: readonly VbaToken[]): { rows: number; cols: number } | undefined {
	const member = tokenText(toks[toks.length - 1]);
	if ((member !== 'value' && member !== 'value2') || toks[toks.length - 2]?.rawText !== '.' || toks[toks.length - 3]?.rawText !== ')') {
		return undefined;
	}
	const close = toks.length - 3;
	const open = toks.findIndex((tok, i) => tok.rawText === '(' && matchParenFrom(toks, i) === close);
	if (open <= 0 || tokenText(toks[open - 1]) !== 'range' || close !== open + 2 || toks[open + 1].kind !== 'stringLiteral') {
		return undefined;
	}
	const text = toks[open + 1].rawText.slice(1, -1);
	const address = /^([A-Za-z]{1,3})(\d+)(?::([A-Za-z]{1,3})(\d+))?$/.exec(text);
	if (!address) {
		return undefined;
	}
	const rows = Math.abs(Number(address[4] ?? address[2]) - Number(address[2])) + 1;
	const cols = Math.abs(columnNumber(address[3] ?? address[1]) - columnNumber(address[1])) + 1;
	return { rows, cols };
}

/**
 * Whether `toks` read one cell's value: `Range("A1")`, `Cells(1, 2)`, either
 * with `.Value` or `.Value2`, after any receiver (issue #278).
 */
export function singleCellValue(toks: readonly VbaToken[]): boolean {
	const member = tokenText(toks[toks.length - 1]);
	const end = (member === 'value' || member === 'value2') && toks[toks.length - 2]?.rawText === '.' ? toks.length - 3 : toks.length - 1;
	if (toks[end]?.rawText !== ')') {
		return false;
	}
	const open = toks.findIndex((tok, i) => tok.rawText === '(' && matchParenFrom(toks, i) === end);
	const callee = open > 0 ? tokenText(toks[open - 1]) : '';
	if (open <= 0 || (open > 1 && toks[open - 2]?.rawText !== '.')) {
		return false;
	}
	const args = splitTopLevelTokenGroups(toks.slice(open + 1, end), ',');
	if (callee === 'range') {
		const block = args.length === 1 && args[0].length === 1 && args[0][0].kind === 'stringLiteral'
			? /^([A-Za-z]{1,3})(\d+)(?::([A-Za-z]{1,3})(\d+))?$/.exec(args[0][0].rawText.slice(1, -1))
			: undefined;
		return block !== undefined && block !== null && (block[3] === undefined || (block[3].toLowerCase() === block[1].toLowerCase() && block[4] === block[2]));
	}
	return callee === 'cells' && args.length === 2 && args.every((arg) => arg.length === 1 && arg[0].kind === 'integerLiteral');
}

/** The argument of `Application.Transpose(...)` or `WorksheetFunction.Transpose(...)` that is the whole of `toks`. */
function transposeArgument(toks: readonly VbaToken[]): VbaToken[] | undefined {
	const at = toks.findIndex((tok, i) => tokenText(tok) === 'transpose' && toks[i + 1]?.rawText === '(' && toks[i - 1]?.rawText === '.');
	const receiver = at > 0 ? toks.slice(0, at - 1).map((tok) => tokenText(tok)).join('') : '';
	if (at < 0 || !['application', 'worksheetfunction', 'application.worksheetfunction'].includes(receiver) || matchParenFrom(toks, at + 1) !== toks.length - 1) {
		return undefined;
	}
	return toks.slice(at + 2, toks.length - 1) as VbaToken[];
}

/**
 * `Filter(source, match[, include[, compare]])` over an array whose every
 * element is a known string or whole number: the elements whose text holds
 * `match` (or lacks it, with include False), 0-based whatever Option Base
 * says (measured in Excel 16.0, issue #260). An empty match keeps every
 * element.
 */
function filterShape(args: readonly (readonly VbaToken[])[], name: string, optionBase: number, strings: StringValueOf | undefined, compare: ModuleCompare | undefined): FixedArrayBound | undefined {
	if (args.length < 2 || args.length > 4) {
		return undefined;
	}
	const source = arrayValueShape(args[0], name, optionBase, strings, compare);
	const match = stringArgument(args[1], strings);
	const include = args.length >= 3 && args[2].length > 0 ? booleanArgument(args[2]) : true;
	const textCompare = args.length === 4 ? compareArgument(args[3]) : defaultCompare(compare, match);
	if (!source || source.dims.length !== 1 || match === undefined || include === undefined || textCompare === undefined) {
		return undefined;
	}
	const texts: string[] = [];
	for (let k = 0; k <= source.dims[0].upper - source.dims[0].lower; k++) {
		const value = source.values?.[k];
		if (value === undefined || source.elements?.[k] !== undefined) {
			return undefined;
		}
		texts.push(String(value));
	}
	const kept: string[] = [];
	for (const text of texts) {
		const holds = textCompare ? caselessIndexOf(text, match, 0) : text.indexOf(match);
		if (holds === undefined) {
			return undefined;
		}
		if ((holds >= 0) === include) {
			kept.push(text);
		}
	}
	return {
		name,
		dims: [{ lower: 0, upper: kept.length - 1, explicitLower: true }],
		origin: 'Filter(...)',
		...(kept.length > 0 ? { values: kept } : {}),
	};
}

/**
 * The names whose elements a statement of the procedure may write, so the
 * values an Array or Split gave them are not trusted (issue #260): `v(1) = 2`,
 * a writing statement that mentions them (Set, LSet, Mid, Input #, Get #,
 * ReDim, Erase), and an element passed alone to a call, ByRef by default.
 */
export function elementsWrittenIn(source: string, proc: ProcedureNode, activity: ConditionalActivityTracker | undefined): Set<string> {
	const out = new Set<string>();
	forEachStatement(proc.body as BodyNode[], (stmt) => {
		for (const span of statementAndBranchSpansOf(stmt)) {
			const toks = statementTokensAfterLeadingLabel(source, span).filter((tok) => tok.kind !== 'comment');
			const head = tokenText(toks[0]);
			if (ELEMENT_WRITING_HEADS.has(head)) {
				for (const tok of toks) {
					const lower = tokenName(tok)?.toLowerCase();
					if (lower) {
						out.add(lower);
					}
				}
				continue;
			}
			const call = !toks.some((tok, k) => tok.rawText === '=' && tok.kind === 'operator' && topLevelAt(toks, k));
			for (let i = 0; i + 1 < toks.length; i++) {
				const lower = tokenName(toks[i])?.toLowerCase();
				if (!lower || toks[i + 1].rawText !== '(') {
					continue;
				}
				const close = matchParenFrom(toks, i + 1);
				const next = toks[close + 1];
				const prev = toks[i - 1];
				const target = i === (head === 'let' ? 1 : 0) && next?.rawText === '=';
				const opensSlot = prev?.rawText === '(' || prev?.rawText === ',' || (call && i > 0 && (prev.kind === 'identifier' || prev.kind === 'keyword'));
				const closesSlot = next === undefined || next.rawText === ')' || next.rawText === ',' || next.rawText === ':';
				if (target || (opensSlot && closesSlot)) {
					out.add(lower);
				}
			}
		}
	}, activity);
	return out;
}

const ELEMENT_WRITING_HEADS: ReadonlySet<string> = new Set(['set', 'lset', 'rset', 'mid', 'mid$', 'input', 'get', 'line', 'redim', 'erase']);

function topLevelAt(toks: readonly VbaToken[], index: number): boolean {
	let depth = 0;
	for (let k = 0; k < index; k++) {
		if (toks[k].rawText === '(') {
			depth++;
		} else if (toks[k].rawText === ')') {
			depth--;
		}
	}
	return depth === 0;
}

/** An operand that reads one element of an array whose values are known. */
export interface ElementOperand {
	/** The operand's first and last token. */
	first: number;
	last: number;
	value: string | number;
}

/**
 * The element `v(1)` or `Split("1 b")(1)` reads, where the operand ends at
 * token `end` (issue #260). `shapes` holds the arrays the locals are known
 * to be.
 */
export function elementOperandEndingAt(toks: readonly VbaToken[], end: number, shapes: ReadonlyMap<string, FixedArrayBound>, optionBase: number): ElementOperand | undefined {
	if (toks[end]?.rawText !== ')') {
		return undefined;
	}
	let depth = 0;
	for (let open = end; open >= 0; open--) {
		if (toks[open].rawText === ')') {
			depth++;
		} else if (toks[open].rawText === '(' && --depth === 0) {
			return elementAt(toks, open, end, shapes, optionBase);
		}
	}
	return undefined;
}

/** The element an operand starting at token `start` reads, as {@link elementOperandEndingAt}. */
export function elementOperandStartingAt(toks: readonly VbaToken[], start: number, shapes: ReadonlyMap<string, FixedArrayBound>, optionBase: number): ElementOperand | undefined {
	if (!tokenName(toks[start]) || toks[start + 1]?.rawText !== '(') {
		return undefined;
	}
	let close = matchParenFrom(toks, start + 1);
	if (close > 0 && toks[close + 1]?.rawText === '(') {
		close = matchParenFrom(toks, close + 1);
	}
	const element = close > 0 ? elementOperandEndingAt(toks, close, shapes, optionBase) : undefined;
	const next = toks[close + 1]?.rawText;
	return element?.first === start && next !== '(' && next !== '.' && next !== '!' ? element : undefined;
}

function elementAt(toks: readonly VbaToken[], open: number, close: number, shapes: ReadonlyMap<string, FixedArrayBound>, optionBase: number): ElementOperand | undefined {
	let shape: FixedArrayBound | undefined;
	let first = open - 1;
	if (toks[first]?.rawText === ')') {
		// `Split(...)(k)`, `Array(...)(k)`, `Filter(...)(k)`.
		let depth = 0;
		let callOpen = first;
		for (; callOpen >= 0; callOpen--) {
			if (toks[callOpen].rawText === ')') {
				depth++;
			} else if (toks[callOpen].rawText === '(' && --depth === 0) {
				break;
			}
		}
		first = callOpen - 1;
		if (first >= 2 && toks[first - 1].rawText === '.' && tokenText(toks[first - 2]) === 'vba') {
			first -= 2;
		}
		const callee = tokenText(toks[callOpen - 1]);
		if (callOpen < 1 || (callee !== 'split' && callee !== 'array' && callee !== 'filter')) {
			return undefined;
		}
		shape = arrayValueShape(toks.slice(first, open), '', optionBase);
	} else {
		const lower = tokenName(toks[first])?.toLowerCase();
		shape = lower ? shapes.get(lower) : undefined;
	}
	const before = toks[first - 1]?.rawText;
	if (!shape || shape.dims.length !== 1 || before === '.' || before === '!') {
		return undefined;
	}
	const index = signedIntegerArgument(toks.slice(open + 1, close).filter((tok) => tok.kind !== 'comment'));
	const position = index === undefined ? -1 : index - shape.dims[0].lower;
	const value = shape.values?.[position];
	return value === undefined || shape.elements?.[position] !== undefined ? undefined : { first, last: close, value };
}

/** The parts `Split` returns for a non-empty text, at most `limit` of them unless it is -1. */
function splitParts(text: string, delimiter: string, limit: number, textCompare: boolean): string[] | undefined {
	const parts: string[] = [];
	let from = 0;
	while (limit === -1 || parts.length < limit - 1) {
		const at = textCompare ? caselessIndexOf(text, delimiter, from) : text.indexOf(delimiter, from);
		if (at === undefined) {
			return undefined;
		}
		if (at < 0) {
			break;
		}
		parts.push(text.slice(from, at));
		from = at + delimiter.length;
	}
	parts.push(text.slice(from));
	return parts;
}

/** Where vbTextCompare finds `needle`, for ASCII text only, whose case folding is certain. */
function caselessIndexOf(text: string, needle: string, from: number): number | undefined {
	if (/[^\x00-\x7f]/.test(text + needle)) {
		return undefined;
	}
	return text.toLowerCase().indexOf(needle.toLowerCase(), from);
}

/** A string literal, or a name `strings` knows, as an argument. */
function stringArgument(arg: readonly VbaToken[], strings: StringValueOf | undefined): string | undefined {
	if (arg.length !== 1) {
		// `Split(LCase("a,b"), ",")`, `Split("ab" & "c", ",")` (issue #509).
		return foldStringExpression(arg, {
			nameValue: (tok) => (tok.kind === 'identifier' ? strings?.(tok) : undefined),
			integerValue: (toks) => signedIntegerArgument(toks.filter((tok) => tok.kind !== 'comment')),
		});
	}
	if (arg[0].kind === 'stringLiteral') {
		return stringLiteralValue(arg[0].rawText);
	}
	return arg[0].kind === 'identifier' ? strings?.(arg[0]) : undefined;
}

/** An integer literal, with a leading minus or not. */
function signedIntegerArgument(arg: readonly VbaToken[]): number | undefined {
	const negative = arg.length === 2 && arg[0].rawText === '-';
	if (arg.length !== (negative ? 2 : 1) || arg[arg.length - 1].kind !== 'integerLiteral') {
		return undefined;
	}
	const value = parseVbaIntegerLiteral(arg[arg.length - 1].rawText);
	return value === undefined ? undefined : negative ? -value : value;
}

/** True, False or an integer literal, as Filter's include. */
function booleanArgument(arg: readonly VbaToken[]): boolean | undefined {
	const word = arg.length === 1 ? tokenText(arg[0]) : '';
	if (word === 'true' || word === 'false') {
		return word === 'true';
	}
	const value = signedIntegerArgument(arg);
	return value === undefined ? undefined : value !== 0;
}

/** Whether a compare argument asks for vbTextCompare; undefined for anything else than the two. */
/**
 * Whether Split or Filter with no compare argument compares as text: it
 * takes the module's Option Compare, and Text and Database both ignore
 * case (issues #353 and #405, measured in Excel 16.0 and Access 16.0).
 * Where the module's setting is not known, only a delimiter or match with
 * no cased letter is settled, since both comparisons then agree.
 */
function defaultCompare(compare: ModuleCompare | undefined, text: string | undefined): boolean | undefined {
	if (compare !== undefined) {
		return compare !== 'binary';
	}
	return text !== undefined && text.toLowerCase() === text.toUpperCase() ? false : undefined;
}

function compareArgument(arg: readonly VbaToken[]): boolean | undefined {
	const word = arg.length === 1 ? tokenText(arg[0]) : '';
	if (word === 'vbbinarycompare' || word === 'vbtextcompare') {
		return word === 'vbtextcompare';
	}
	const value = signedIntegerArgument(arg);
	return value === 0 || value === 1 ? value === 1 : undefined;
}

/** The string or whole number one `Array(...)` element is written as. */
function literalElementValue(group: readonly VbaToken[]): string | number | undefined {
	const toks = group.filter((tok) => tok.kind !== 'comment');
	if (toks.length === 1 && toks[0].kind === 'stringLiteral') {
		return stringLiteralValue(toks[0].rawText);
	}
	return signedIntegerArgument(toks);
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
	lookup?: IntegerConstantLookup,
): SubscriptHit[] {
	const toks = statementTokensAfterLeadingLabel(source, span);
	const out: SubscriptHit[] = [];
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
			out.push(dimensionCountViolation(span, toks[i], toks[close], decl, slots.length));
			continue;
		}
		// One report per access: the first dimension that is out of range.
		let hit: SubscriptHit | undefined;
		for (let index = 0; index < slots.length && !hit; index++) {
			hit = subscriptViolation(span, decl, fixed, slots[index], index, counters, lookup);
		}
		hit ??= elementSubscriptViolation(span, toks, decl, slots, close, lookup);
		if (hit) {
			out.push(hit);
		}
	}
	return out;
}

/**
 * `g(1)` on `Dim g(2, 2)`: a subscript count other than the array's
 * dimensions (issue #248, measured in Excel 16.0). The compiler knows a
 * Dim's dimensions and refuses the line; bounds a ReDim or a value set are
 * found when it runs, error 9.
 */
export function dimensionCountViolation(span: Span, first: VbaToken, last: VbaToken, shape: FixedArrayBound, given: number): SubscriptHit {
	const at = { start: span.start + first.start, end: span.start + last.end };
	const counts = `has ${pluralizeCount(shape.dims.length, 'dimension')}, and ${pluralizeCount(given, 'subscript')} ${given === 1 ? 'is' : 'are'} given here`;
	return shape.origin === 'Dim'
		? { span: at, rule: 'wrongNumberOfDimensions', message: `Array '${shape.name}' ${counts}. This is a VBE compile error: Wrong number of dimensions.` }
		: { span: at, message: `Array '${shape.name}' (${shape.origin}) ${counts}. This will raise Run-time error '9': Subscript out of range.` };
}

const NO_SHAPES: ReadonlyMap<string, FixedArrayBound> = new Map();

/**
 * The subscripts in the parentheses at `open` against an array whose
 * bounds are known, and on through the arrays its elements hold where
 * Array(...) of Array(...) built it (issue #248, measured in Excel 16.0):
 * `v(0)(5)`, `c(1)(5)`.
 */
export function shapeSubscriptViolation(
	span: Span,
	toks: readonly VbaToken[],
	shape: FixedArrayBound,
	open: number,
	lookup?: IntegerConstantLookup,
): SubscriptHit | undefined {
	const close = matchParenFrom(toks, open);
	if (close <= open + 1) {
		return undefined;
	}
	const slots = splitTopLevelTokenGroups(toks.slice(open + 1, close).filter((tok) => tok.kind !== 'comment'), ',');
	if (slots.some((slot) => slot.length === 0)) {
		return undefined;
	}
	if (slots.length !== shape.dims.length) {
		return dimensionCountViolation(span, toks[open], toks[close], shape, slots.length);
	}
	for (let index = 0; index < slots.length; index++) {
		const hit = subscriptViolation(span, shape, NO_SHAPES, slots[index], index, undefined, lookup);
		if (hit) {
			return hit;
		}
	}
	return elementSubscriptViolation(span, toks, shape, slots, close, lookup);
}

/** `v(0)(5)`: the next parentheses, against the array element `v(0)` holds. */
function elementSubscriptViolation(
	span: Span,
	toks: readonly VbaToken[],
	shape: FixedArrayBound,
	slots: readonly (readonly VbaToken[])[],
	close: number,
	lookup?: IntegerConstantLookup,
): SubscriptHit | undefined {
	if (!shape.elements || slots.length !== 1 || toks[close + 1]?.rawText !== '(') {
		return undefined;
	}
	const value = comparableArrayBoundExpressionValue(slots[0])
		?? (lookup ? evaluateIntegerConstantExpression(slots[0].map((tok) => tok.rawText).join(' '), lookup) : undefined);
	const element = value === undefined ? undefined : shape.elements[value - shape.dims[0].lower];
	return element ? shapeSubscriptViolation(span, toks, { ...element, name: `${shape.name}(${value})` }, close + 1, lookup) : undefined;
}

const RETURN_SHAPES = new WeakMap<ModuleNode, ReadonlyMap<string, FixedArrayBound>>();

/** Words that may leave a Function before its one return assignment runs. */
const RETURN_SKIPPING_WORDS: ReadonlySet<string> = new Set(['exit', 'goto', 'gosub', 'return', 'resume', 'on', 'raise', 'error', 'stop']);

/**
 * The bounds of the array each Function of the module returns, where they
 * are known from its body (issue #240, measured in Excel 16.0): one top-level
 * `F = r` with r a fixed local array, or `F = Array(1, 2)`, the only
 * statement that names F, in a body nothing can leave early. So `F()(5)`
 * reads past the end.
 */
function functionReturnShapes(
	source: string,
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
	optionBase: number,
): ReadonlyMap<string, FixedArrayBound> {
	// Every procedure asks; a parse makes a new module, so the module is the key.
	const cached = RETURN_SHAPES.get(mod);
	if (cached) {
		return cached;
	}
	const out = new Map<string, FixedArrayBound>();
	RETURN_SHAPES.set(mod, out);
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure' || member.procKind !== 'Function') {
			continue;
		}
		const lower = member.name.toLowerCase();
		const body = statementTokens(source, member.span);
		const named = body.filter((tok) => tokenName(tok)?.toLowerCase() === lower).length;
		if (named !== 2 || body.some((tok, i) => RETURN_SKIPPING_WORDS.has(tokenText(tok)) || (tokenText(tok) === 'end' && !['function', 'if', 'select', 'with'].includes(tokenText(body[i + 1]))))) {
			continue; // the header, and one assignment
		}
		const assignment = member.body.find((node): node is LeafStatementNode => isLeafStatement(node)
			&& !(node.kind === 'Statement' && node.singleLineIfBranches)
			&& bareAssignmentTarget(source, node.span)?.name.toLowerCase() === lower);
		const target = assignment ? bareAssignmentTarget(source, assignment.span) : undefined;
		if (!target) {
			continue;
		}
		const value = target.valueTokens.filter((tok) => tok.kind !== 'comment');
		const local = value.length === 1 ? tokenName(value[0])?.toLowerCase() : undefined;
		const shape = local
			? localFixedArrayDeclarationsForBody(source, member.body, activity, optionBase).get(local)
			: arrayValueShape(value, member.name, optionBase);
		if (shape) {
			out.set(lower, { ...shape, name: `${member.name}()`, origin: `returned by ${member.name}` });
		}
	}
	return out;
}

/** `F()(5)` and `F(1)(5)`: a subscript on the array a Function returns. */
function returnedArraySubscriptViolations(
	source: string,
	span: Span,
	returned: ReadonlyMap<string, FixedArrayBound>,
	lookup: IntegerConstantLookup,
): Array<{ span: Span; message: string }> {
	if (returned.size === 0) {
		return [];
	}
	const toks = statementTokensAfterLeadingLabel(source, span);
	const out: Array<{ span: Span; message: string }> = [];
	for (let i = 0; i < toks.length - 1; i++) {
		const shape = returned.get(tokenName(toks[i])?.toLowerCase() ?? '');
		if (!shape || toks[i + 1].rawText !== '(' || toks[i - 1]?.rawText === '.' || toks[i - 1]?.rawText === '!') {
			continue;
		}
		const call = matchParenFrom(toks, i + 1);
		if (call < 0 || toks[call + 1]?.rawText !== '(') {
			continue;
		}
		const close = matchParenFrom(toks, call + 1);
		const slots = close < 0 ? [] : splitTopLevelTokenGroups(toks.slice(call + 2, close).filter((tok) => tok.kind !== 'comment'), ',');
		if (slots.length !== shape.dims.length || slots.some((slot) => slot.length === 0)) {
			continue;
		}
		for (let index = 0; index < slots.length; index++) {
			const hit = subscriptViolation(span, shape, returned, slots[index], index, undefined, lookup);
			if (hit) {
				out.push(hit);
				break;
			}
		}
	}
	return out;
}

/** A module's fixed arrays, by lowercased name. */
function moduleFixedArrayDeclarations(
	source: string,
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
	optionBase: number,
): Map<string, FixedArrayBound> {
	const groups = activeModuleMembers(mod, activity).filter((member) => member.kind === 'VariableGroup') as BodyNode[];
	return localFixedArrayDeclarationsForBody(source, groups, activity, optionBase);
}

/** The names a procedure's own locals and parameters take, which hide module variables. */
function hiddenIn(symbols: ReturnType<typeof buildModuleSymbols>, proc: ProcedureNode): Set<string> {
	const out = new Set((procedureSymbolFor(symbols, proc)?.children ?? []).map((child) => child.name.toLowerCase()));
	for (const param of proc.params) {
		out.add(param.name.toLowerCase());
	}
	return out;
}

/**
 * `a(0)`, `UBound(a)` and `LBound(a)` on a module's dynamic array that
 * nothing ReDims: it has no elements, and each raises 9 (issue #241,
 * measured in Excel 16.0).
 */
function unallocatedModuleArrayUses(
	source: string,
	span: Span,
	arrays: ReadonlyMap<string, VbaSymbol>,
): Array<{ span: Span; message: string }> {
	const toks = statementTokensAfterLeadingLabel(source, span);
	const out: Array<{ span: Span; message: string }> = [];
	for (let i = 0; i < toks.length; i++) {
		const variable = arrays.get(tokenName(toks[i])?.toLowerCase() ?? '');
		if (!variable || toks[i - 1]?.rawText === '.' || toks[i - 1]?.rawText === '!') {
			continue;
		}
		const bound = toks[i - 1]?.rawText === '(' && ['ubound', 'lbound'].includes(tokenText(toks[i - 2])) && toks[i - 3]?.rawText !== '.'
			&& (toks[i + 1]?.rawText === ')' || toks[i + 1]?.rawText === ',');
		if (!bound && toks[i + 1]?.rawText !== '(') {
			continue;
		}
		const scope = variable.visibility === 'Public' || variable.visibility === 'Global' ? 'the project' : 'this module';
		const what = bound ? `${toks[i - 2].rawText} reads the bounds of '${toks[i].rawText}'` : `'${toks[i].rawText}' is indexed`;
		out.push({
			span: { start: span.start + toks[i].start, end: span.start + toks[i].end },
			message: `${what}, a dynamic array nothing in ${scope} ever ReDims, so it has no elements. This will raise Run-time error '9': Subscript out of range.`,
		});
	}
	return out;
}

/** A signed decimal literal, `2.6` or `-0.6`, and the whole number VBA rounds it to. */
function roundedDecimalLiteral(slot: readonly VbaToken[]): { text: string; whole: number } | undefined {
	const toks = slot.filter((tok) => tok.kind !== 'comment');
	const negative = toks.length === 2 && toks[0].rawText === '-';
	const literal = toks[negative ? 1 : 0];
	if (toks.length !== (negative ? 2 : 1) || literal?.kind !== 'floatLiteral') {
		return undefined;
	}
	const read = Number(literal.rawText.replace(/[!#@]$/, ''));
	if (!Number.isFinite(read)) {
		return undefined;
	}
	const value = negative ? -read : read;
	return { text: toks.map((tok) => tok.rawText).join(''), whole: bankersRound(value) + 0 };
}

export function subscriptViolation(
	span: Span,
	decl: FixedArrayBound,
	fixed: ReadonlyMap<string, FixedArrayBound>,
	slot: readonly VbaToken[],
	index: number,
	counters: CountersAt | undefined,
	lookup?: IntegerConstantLookup,
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
	// `a(2.6)` rounds half to even, to 3 (issue #286, measured in Excel 16.0).
	const decimal = roundedDecimalLiteral(slot);
	if (decimal !== undefined) {
		const detail = subscriptDetail(decimal.whole, dim, index, decl.dims.length);
		return detail
			? { span: slotSpan, message: `Subscript ${decimal.text} rounds to ${decimal.whole}, which for array '${decl.name}'${from} ${detail}. ${error}` }
			: undefined;
	}
	// `a(i)`, and `a(i + 1)` or `a(i - 1)` a whole number off it (issue #263).
	const offsetLiteral = slot.length === 3 && (slot[1].rawText === '+' || slot[1].rawText === '-') && slot[2].kind === 'integerLiteral'
		? parseVbaIntegerLiteral(slot[2].rawText)
		: undefined;
	const offset = offsetLiteral === undefined ? 0 : slot[1].rawText === '-' ? -offsetLiteral : offsetLiteral;
	const counter = slot.length === 1 || offsetLiteral !== undefined ? counters?.get(tokenName(slot[0])?.toLowerCase() ?? '') : undefined;
	if (!counter) {
		// A Const, or a local with one known value here (issue #238).
		const text = slot.map((tok) => tok.rawText).join(' ');
		const known = lookup ? evaluateIntegerConstantExpression(text, lookup) : undefined;
		const detail = known === undefined ? undefined : subscriptDetail(known, dim, index, decl.dims.length);
		return detail
			? { span: slotSpan, message: `Subscript ${text} is ${known} here, which for array '${decl.name}'${from} ${detail}. ${error}` }
			: undefined;
	}
	// `a(i)` inside `For i = 0 To 3`: the counter's first and last passes.
	const atomValue = (atom: { kind: string; name: string; dimension: number }): number | undefined => {
		// `For i = s To 2` with s a local known to hold 1 (issue #346): the
		// loop never writes s, so it holds the same here as where it starts.
		if (atom.kind === 'local') {
			return lookup ? evaluateIntegerConstantExpression(atom.name, lookup) : undefined;
		}
		const shape = fixed.get(atom.name)?.dims[atom.dimension - 1];
		return atom.kind === 'ubound' ? shape?.upper : atom.kind === 'lbound' ? shape?.lower : undefined;
	};
	for (const pass of numericCounterPasses(counter, atomValue)) {
		const detail = subscriptDetail(pass.value + offset, dim, index, decl.dims.length);
		if (detail) {
			const reached = pass.pass === 'first'
				? `Counter '${slot[0].rawText}' is ${pass.value} on its first pass`
				: `Counter '${slot[0].rawText}' reaches ${pass.value} on its last pass`;
			const subscript = offset === 0 ? '' : `, so ${slot.map((tok) => tok.rawText).join(' ')} is ${pass.value + offset}`;
			return { span: slotSpan, message: `${reached}${subscript}, which for array '${decl.name}'${from} ${detail}. ${error}` };
		}
	}
	return offset === 0 ? symbolicCounterSubscript(span, decl.name, decl.name.toLowerCase(), [slot], counters, index) : undefined;
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
 * element sits at 0 (issue #120), and of Filter (issue #260), and on a
 * String local known to hold its text (issue #559).
 */
function inlineSplitIndexViolations(source: string, span: Span, shadowed: (name: string) => boolean, strings?: StringValueOf): Array<{ span: Span; message: string }> {
	const toks = statementTokensAfterLeadingLabel(source, span);
	const out: Array<{ span: Span; message: string }> = [];
	for (let i = 0; i + 1 < toks.length; i++) {
		const callee = tokenText(toks[i]);
		// A project procedure named Split takes the call (issue #280).
		if ((callee !== 'split' && callee !== 'filter') || toks[i + 1].rawText !== '(' || !isBareOrVbaQualifiedIntrinsicCall(toks, i) || (toks[i - 1]?.rawText !== '.' && shadowed(toks[i].rawText))) {
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
		const shape = arrayValueShape(toks.slice(i, close + 1), 'Split(...)', 0, strings, moduleCompare(source));
		const indexToks = toks.slice(close + 2, indexClose).filter((tok) => tok.kind !== 'comment');
		const value = comparableArrayBoundExpressionValue(indexToks);
		if (!shape || value === undefined) {
			continue;
		}
		const detail = subscriptDetail(value, shape.dims[0], 0, 1);
		if (detail) {
			out.push({
				span: { start: span.start + indexToks[0].start, end: span.start + indexToks[indexToks.length - 1].end },
				message: `Subscript ${value} for the array ${callee === 'split' ? 'Split' : 'Filter'} returns here ${detail}. This will raise Run-time error '9': Subscript out of range.`,
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
	projectIntegerConstants?: ReadonlyMap<string, string | undefined>,
	projectVisibleSymbols?: readonly VbaSymbol[],
	hostModel?: HostObjectModel,
): void {
	const optionBase = moduleOptionBase(mod, activity);
	const moduleConstants = moduleIntegerConstants(mod, projectIntegerConstants, activity);
	let moduleFixed: Map<string, FixedArrayBound> | undefined;
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		// A module's fixed arrays keep their bounds, and its dynamic arrays
		// that nothing ReDims have none (issue #241).
		const moduleVariables = untouchedModuleVariablesIn(source, symbols, member);
		const declared = new Map([
			...[...(moduleFixed ??= moduleFixedArrayDeclarations(source, mod, activity, optionBase))].filter(([lower]) => !hiddenIn(symbols, member).has(lower)),
			...localFixedArrayDeclarationsForBody(source, member.body, activity, optionBase),
		]);
		const unallocated = new Map([...moduleVariables].filter(([, variable]) => variable.isArray && variable.arrayBounds === undefined));
		const shapesAt = knownArrayShapesAt(source, symbols, member, activity, optionBase);
		const returned = functionReturnShapes(source, mod, activity, optionBase);
		const redimmed = redimShapesAt(source, symbols, member, activity, optionBase);
		const merged = new Map<ReadonlyMap<string, FixedArrayBound>, Map<ReadonlyMap<string, FixedArrayBound> | undefined, ReadonlyMap<string, FixedArrayBound>>>();
		const fixedAt = (stmt: LeafStatementNode): ReadonlyMap<string, FixedArrayBound> => {
			const shapes = shapesAt(stmt);
			const reshaped = redimmed.get(stmt);
			const byReshape = merged.get(shapes) ?? new Map();
			merged.set(shapes, byReshape);
			let fixed = byReshape.get(reshaped);
			if (!fixed) {
				const next = new Map(declared);
				for (const [lower, shape] of [...shapes, ...(reshaped ?? [])]) {
					if (!declared.has(lower)) {
						next.set(lower, shape);
					}
				}
				fixed = next;
				byReshape.set(reshaped, fixed);
			}
			return fixed;
		};
		const redimTargets = redimTargetNamesInBody(source, member.body, activity);
		const excludedAt = (stmt: LeafStatementNode): ReadonlySet<string> => {
			const reshaped = redimmed.get(stmt);
			return reshaped ? new Set([...redimTargets].filter((lower) => !reshaped.has(lower))) : redimTargets;
		};
		const counters = loopCountersAt(source, member.body, activity);
		// A subscript through a Const or a local with one known value (issue #238).
		const constants = procedureIntegerConstantLookup(member, moduleConstants, symbols, projectVisibleSymbols, activity, hostModel);
		const valuesAt = knownLocalLiteralValuesAt(source, member, symbols, activity);
		// `Split(s, ",")(3)` with s a String local known to hold "q,r" (issue #559).
		const stringsAt = stringValuesAt(source, symbols, member, activity);
		// Code that never runs, after `GoTo Done` or in a loop of no pass,
		// raises nothing, whatever state it builds (issue #406).
		const unreachable = unreachableStatementsIn(source, member, symbols, activity);
		let sourceNames: ReturnType<typeof sourceNameScopeFor> | undefined;
		// Headers too: `For i = 1 To a(5)`, `Select Case a(5)` (issue #233).
		forEachStatementWithHeaders(source, member.body, (stmt) => {
			if (unreachable.has(stmt)) {
				return;
			}
			for (const hit of inlineSplitIndexViolations(source, stmt.span, (name) => runtimeCallableSourceShadowed(name, sourceNames ??= sourceNameScopeFor(symbols, member, projectVisibleSymbols)), stringsAt(stmt))) {
				push('arraySubscriptOutOfBounds', hit.message, hit.span);
			}
			for (const hit of unallocated.size === 0 ? [] : unallocatedModuleArrayUses(source, stmt.span, unallocated)) {
				push('arraySubscriptOutOfBounds', hit.message, hit.span);
			}
			for (const hit of returned.size === 0 ? [] : returnedArraySubscriptViolations(source, stmt.span, returned, withKnownLocals(constants, valuesAt(stmt)))) {
				push('arraySubscriptOutOfBounds', hit.message, hit.span);
			}
			const fixed = fixedAt(stmt);
			const stmtCounters = counters.get(stmt);
			if (fixed.size === 0 && !stmtCounters) {
				return;
			}
			const excluded = excludedAt(stmt);
			for (const hit of fixedArraySubscriptViolations(source, stmt.span, fixed, excluded, stmtCounters, withKnownLocals(constants, valuesAt(stmt)))) {
				push(hit.rule ?? 'arraySubscriptOutOfBounds', hit.message, hit.span);
			}
			for (const hit of boundIntrinsicDimensionViolations(source, stmt.span, fixed, excluded)) {
				push('arraySubscriptOutOfBounds', hit.message, hit.span);
			}
		}, activity);
	}
}
