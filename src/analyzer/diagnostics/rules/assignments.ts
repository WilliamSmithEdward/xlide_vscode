// Rule family: assignment validation (audit #0).
//
// Extracted verbatim from analyzeModule.ts: Const reassignment, scalar and
// member assignment type compatibility, Set assignment validation, and
// missing Function/Property Get return assignments.

import type { MemberCompletionContext } from '../../completion/memberAccess';
import type { ConditionalActivityTracker } from '../../conditional/conditionalCompilation';
import {
	matchParenFrom,
	splitTopLevelTokenGroups,
	statementTokensCached,
} from '../../lexer/tokenHelpers';
import type { VbaToken } from '../../lexer/tokenKinds';
import type {
	ModuleNode,
	ProcedureNode,
	Span,
} from '../../parser/nodes';
import { buildModuleSymbols } from '../../symbols/buildModuleSymbols';
import type {
	VbaProcedureSignature,
	VbaSymbol,
} from '../../symbols/symbolModel';
import {
	procedureSymbolFor,
	type PushFn,
} from '../analysisContext';
import {
	type CallableParamType,
	type CallableTypeSignature,
	type CallArguments,
	extractCall,
	extractQualifiedCall,
} from '../callExtraction';
import {
	buildModuleTypeSignatures,
	callableSignatureForCall,
	callableTypeSignaturesFor,
	declarationShapeEnvironmentFor,
	declaredShapeForSourceBinding,
	declaredTypeForSourceBinding,
	type DeclaredValueShape,
	declaredValueTypeForQualifiedSourceBinding,
	declaredValueTypeForSourceBinding,
	incompatibilityReason,
	inferArgumentType,
	isKnownObjectAssignmentType,
	isKnownScalarType,
	namedArgumentSlot,
	nonnumericStringArithmeticOperand,
	normalizeType,
	objectAssignmentIncompatibilityReason,
	resolveExactMemberCompletion,
	type SourceDeclaredShape,
	type SourceDeclaredTypeResolver,
	sourceIdentifierBinding,
	type SourceNameScope,
	sourceNameScopeFor,
	type SourceQualifiedDeclaredTypeResolver,
	typeEnvironmentFor,
} from '../typeInference';
import {
	activeModuleMembers,
	bareAssignmentTarget,
	declaredNameSpan,
	firstExecutableTokenIndex,
	forEachStatement,
	setAssignmentTarget,
	statementTokens,
	statementTokensAfterLeadingLabel,
	stripHeaderBrackets,
	tokenName,
	tokenText,
	topLevelOperatorIndex,
	type ProcedureStatementVisitor,
} from '../walker';

/**
 * Rule: assigning to a constant is illegal. High-confidence form only - the
 * left-hand side must be a bare identifier (no member access, no index) that
 * resolves to a Const declared at module level or in the enclosing procedure.
 */
export function checkConstAssignment(
	source: string,
	symbols: ReturnType<typeof buildModuleSymbols>,
	projectVisibleSymbols: readonly VbaSymbol[] | undefined,
	push: PushFn,
): ProcedureStatementVisitor {
	return (member) => {
		const procSym = procedureSymbolFor(symbols, member);
		return (stmt) => {
			const hit = bareAssignmentTarget(source, stmt.span);
			if (!hit) {
				return;
			}
			const binding = sourceIdentifierBinding(
				symbols,
				procSym,
				projectVisibleSymbols,
				hit.name,
				'assignmentTarget',
			);
			if (
				binding.scope !== 'ambiguous' &&
				binding.definitions.some((definition) => definition.kind === 'constant')
			) {
				push(
					'constAssignment',
					`Cannot assign to constant '${hit.name}'.`,
					hit.span,
				);
			}
		};
	};
}

function memberAssignmentTarget(
	source: string,
	span: Span,
): {
	member: string;
	label: string;
	memberSpan: Span;
	valueTokens: VbaToken[];
	usesSet: boolean;
} | undefined {
	const toks = statementTokens(source, span);
	let i = firstExecutableTokenIndex(toks);
	const usesSet = tokenText(toks[i]) === 'set';
	if (usesSet || tokenText(toks[i]) === 'let') {
		i++;
	}
	const eq = topLevelOperatorIndex(toks.slice(i), '=');
	if (eq < 0) {
		return undefined;
	}
	const equalsIndex = i + eq;
	const lhs = toks.slice(i, equalsIndex);
	if (lhs.length < 2) {
		return undefined;
	}
	const memberTok = lhs[lhs.length - 1];
	if (!tokenName(memberTok) || lhs[lhs.length - 2]?.rawText !== '.') {
		return undefined;
	}
	if (lhs.some((tok) => tok.kind === 'operator' && tok.rawText === '=')) {
		return undefined;
	}
	return {
		member: tokenName(memberTok)!,
		label: source
			.slice(span.start + lhs[0].start, span.start + memberTok.end)
			.trim(),
		memberSpan: {
			start: span.start + memberTok.start,
			end: span.start + memberTok.end,
		},
		valueTokens: toks.slice(equalsIndex + 1),
		usesSet,
	};
}

export function checkAssignmentTypes(
	source: string,
	mod: ModuleNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
	projectVisibleSymbols: readonly VbaSymbol[] | undefined,
	memberCtx: MemberCompletionContext,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	const moduleSignatures = buildModuleTypeSignatures(symbols);
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		const env = typeEnvironmentFor(symbols, member);
		const shapes = declarationShapeEnvironmentFor(symbols, member);
		const sourceNames = sourceNameScopeFor(symbols, member, projectVisibleSymbols);
		const procSym = procedureSymbolFor(symbols, member);
		const resolveExpressionType = (name: string) => declaredValueTypeForSourceBinding(
			symbols,
			procSym,
			projectVisibleSymbols,
			name,
		);
		const resolveQualifiedExpressionType = (qualifier: string, name: string) =>
			declaredValueTypeForQualifiedSourceBinding(
				symbols,
				projectVisibleSymbols,
				qualifier,
				name,
			);
		forEachStatement(member.body, (stmt) => {
			const assignment = bareAssignmentTarget(source, stmt.span);
			if (!assignment) {
				return;
			}
			const targetType = declaredTypeForSourceBinding(
				symbols,
				procSym,
				projectVisibleSymbols,
				assignment.name,
				'assignmentTarget',
			);
			const expected = targetType.resolved
				? targetType.asType
				: env.get(assignment.name.toLowerCase());
			if (!expected) {
				return;
			}
			if (isKnownObjectAssignmentType(expected, memberCtx)) {
				push(
					'setRequired',
					`Object assignment to '${assignment.name}' requires Set because it is declared as ${expected}.`,
					assignment.span,
				);
				return;
			}
			const arraySource = arrayAssignmentToScalarSource(
				assignment,
				stmt.span.start,
				expected,
				shapes,
				(name) => declaredShapeForSourceBinding(
					symbols,
					procSym,
					projectVisibleSymbols,
					name,
					'assignmentTarget',
				),
				(name) => declaredShapeForSourceBinding(
					symbols,
					procSym,
					projectVisibleSymbols,
					name,
					'expression',
				),
			);
			if (arraySource) {
				push(
					'arrayAssignmentToScalar',
					`Array variable '${arraySource.name}' cannot be assigned to scalar '${assignment.name}'. Assign an array element or use a Variant/array target.`,
					arraySource.span,
				);
				return;
			}
			const stringArithmetic = nonnumericStringArithmeticOperand(
				expected,
				assignment.valueTokens,
				stmt.span.start,
			);
			if (stringArithmetic) {
				push(
					'stringArithmeticCoercion',
					`Assignment to '${assignment.name}' expects ${expected}, but this numeric expression contains ${stringArithmetic.label}. This will raise Run-time error '13': Type mismatch.`,
					stringArithmetic.span,
				);
				return;
			}
			const actual = inferArgumentType(
				assignment.valueTokens,
				stmt.span.start,
				env,
				moduleSignatures,
				sourceNames,
				source,
				memberCtx,
				resolveExpressionType,
				resolveQualifiedExpressionType,
			);
			if (!actual) {
				return;
			}
			const reason = incompatibilityReason(expected, actual);
			if (!reason) {
				return;
			}
			push(
				'assignmentTypeMismatch',
				`Assignment to '${assignment.name}' expects ${expected}, but got ${actual.label}. ${reason}`,
				actual.span,
			);
		}, activity);
		checkMemberAssignmentTypes(
			source,
			member,
			env,
			moduleSignatures,
			sourceNames,
			memberCtx,
			activity,
			push,
			resolveExpressionType,
			resolveQualifiedExpressionType,
		);
	}
}

function arrayAssignmentToScalarSource(
	assignment: { name: string; valueTokens: VbaToken[] },
	baseOffset: number,
	expectedType: string,
	shapes: ReadonlyMap<string, DeclaredValueShape>,
	resolveTargetShape?: (name: string) => SourceDeclaredShape,
	resolveSourceShape?: (name: string) => SourceDeclaredShape,
): { name: string; span: Span } | undefined {
	const targetScalarType = knownScalarAssignmentTargetType(
		assignment.name,
		expectedType,
		shapes,
		resolveTargetShape,
	);
	if (!targetScalarType) {
		return undefined;
	}
	if (assignment.valueTokens.length !== 1) {
		return undefined;
	}
	const tok = assignment.valueTokens[0];
	const sourceName = tokenName(tok);
	if (!sourceName) {
		return undefined;
	}
	const resolvedSourceShape = resolveSourceShape?.(sourceName);
	const sourceShape = resolvedSourceShape?.resolved
		? resolvedSourceShape.shape
		: shapes.get(sourceName.toLowerCase());
	if (!sourceShape?.isArray) {
		return undefined;
	}
	// VBA special case (MS-VBAL Let-statement rules): a Byte array is directly
	// assignable to a String scalar - the idiomatic encoding-conversion pattern
	// (`s = bytes`). Only Byte element types are exempt; every other element
	// type remains a compile error.
	if (targetScalarType === 'string' && normalizeType(sourceShape.asType) === 'byte') {
		return undefined;
	}
	return {
		name: sourceName,
		span: { start: baseOffset + tok.start, end: baseOffset + tok.end },
	};
}

function knownScalarAssignmentTargetType(
	name: string,
	expectedType: string,
	shapes: ReadonlyMap<string, DeclaredValueShape>,
	resolveShape?: (name: string) => SourceDeclaredShape,
): string | undefined {
	const resolvedTargetShape = resolveShape?.(name);
	const targetShape = resolvedTargetShape?.resolved
		? resolvedTargetShape.shape
		: shapes.get(name.toLowerCase());
	if (targetShape?.isArray) {
		return undefined;
	}
	const normalized = normalizeType(targetShape?.asType ?? expectedType);
	return normalized && isKnownScalarType(normalized) ? normalized : undefined;
}

/**
 * Rule: a Function/Property Get returns through its hidden return variable.
 * Falling through without assigning that variable is legal VBA, but it silently
 * returns the default value. XLIDE only surfaces this for untyped returns, where
 * the implicit Variant fallthrough is more likely to be accidental than an
 * intentional typed default value.
 */
export function checkMissingReturnAssignments(
	source: string,
	mod: ModuleNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
	projectProcedures: ReadonlyMap<string, readonly VbaProcedureSignature[]> | undefined,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	const moduleSignatures = callableTypeSignaturesFor(symbols, projectProcedures);
	for (const member of activeModuleMembers(mod, activity)) {
		if (
			member.kind !== 'Procedure' ||
			(member.procKind !== 'Function' && member.procKind !== 'PropertyGet')
		) {
			continue;
		}
		if (!member.closed) {
			continue;
		}
		if (member.returnType) {
			continue;
		}
		if (procedureHasReturnAssignment(source, member, activity, moduleSignatures)) {
			continue;
		}
		const procLabel = member.procKind === 'PropertyGet' ? 'Property Get' : 'Function';
		push(
			'missingReturnAssignment',
			`${procLabel} '${member.name}' has no return assignment; VBA will return the default value. Assign to '${member.name}' before exit if a value is intended.`,
			declaredNameSpan(source, member.span, member.name),
		);
	}
}

function procedureHasReturnAssignment(
	source: string,
	proc: ProcedureNode,
	activity: ConditionalActivityTracker | undefined,
	moduleSignatures: ReadonlyMap<string, CallableTypeSignature>,
): boolean {
	const lower = proc.name.toLowerCase();
	let found = false;
	const assignsIn = (span: { start: number; end: number }): boolean => {
		const bare = bareAssignmentTarget(source, span);
		if (bare?.name.toLowerCase() === lower) {
			return true;
		}
		const set = setAssignmentTarget(source, span);
		if (set?.name.toLowerCase() === lower) {
			return true;
		}
		const call = extractCall(source, span);
		const qualifiedCall = call
			? undefined
			: extractQualifiedCall(source, span, moduleSignatures);
		const effectiveCall = call ?? qualifiedCall;
		return Boolean(
			effectiveCall && callPassesNameToByRefParam(effectiveCall, lower, moduleSignatures),
		);
	};
	forEachStatement(proc.body, (stmt) => {
		if (found) {
			return;
		}
		if (assignsIn(stmt.span)) {
			found = true;
			return;
		}
		// A single-line If is ONE leaf statement, so the statement walk never
		// reaches the assignment after `Then`, and `If ok Then F = 1` read as a
		// Function that never assigns its own name.
		for (const branch of singleLineIfBranchSpans(source, stmt.span)) {
			if (assignsIn(branch)) {
				found = true;
				return;
			}
		}
	}, activity);
	return found;
}

/**
 * The statement spans a single-line `If` carries after `Then` and after `Else`,
 * or nothing when the statement is not one. A block `If` header ends at `Then`
 * and yields no branch here; its body is walked as ordinary statements.
 */
function singleLineIfBranchSpans(
	source: string,
	span: { start: number; end: number },
): Array<{ start: number; end: number }> {
	const tokens = statementTokensCached(source, span);
	if (tokens.length === 0 || (tokens[0].canonicalText ?? tokens[0].rawText).toLowerCase() !== 'if') {
		return [];
	}
	const starts: number[] = [];
	for (let i = 1; i < tokens.length; i += 1) {
		const word = (tokens[i].canonicalText ?? tokens[i].rawText).toLowerCase();
		if (word !== 'then' && word !== 'else') {
			continue;
		}
		const next = tokens[i + 1];
		if (next) {
			// statementTokensCached numbers its tokens from the start of the span.
			starts.push(span.start + next.start);
		}
	}
	return starts.map((start, index) => ({
		start,
		end: index + 1 < starts.length ? starts[index + 1] : span.end,
	}));
}

function callPassesNameToByRefParam(
	call: CallArguments,
	lowerName: string,
	moduleSignatures: ReadonlyMap<string, CallableTypeSignature>,
): boolean {
	const sig = callableSignatureForCall(call, moduleSignatures);
	if (!sig) {
		return false;
	}
	let positionalIndex = 0;
	for (const slot of call.slots) {
		const named = namedArgumentSlot(slot);
		let param: CallableParamType | undefined;
		let valueSlot = slot;
		if (named) {
			param = sig.params.find((p) => stripHeaderBrackets(p.name).toLowerCase() === named.name.toLowerCase());
			valueSlot = named.value;
		} else {
			param = sig.params[Math.min(positionalIndex, sig.params.length - 1)];
			positionalIndex++;
		}
		if (!param?.byRef || !singleSlotNameEquals(valueSlot, lowerName)) {
			continue;
		}
		return true;
	}
	return false;
}

function singleSlotNameEquals(slot: readonly VbaToken[], lowerName: string): boolean {
	const toks = slot.filter((t) => t.kind !== 'comment' && t.kind !== 'newline');
	return toks.length === 1 && tokenName(toks[0])?.toLowerCase() === lowerName;
}

function checkMemberAssignmentTypes(
	source: string,
	member: ProcedureNode,
	env: ReadonlyMap<string, string>,
	moduleSignatures: ReadonlyMap<string, CallableTypeSignature>,
	sourceNames: SourceNameScope,
	memberCtx: MemberCompletionContext,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
	resolveExpressionType?: SourceDeclaredTypeResolver,
	resolveQualifiedExpressionType?: SourceQualifiedDeclaredTypeResolver,
): void {
	if (!memberCtx.projectClassMembers || memberCtx.projectClassMembers.length === 0) {
		return;
	}
	forEachStatement(member.body, (stmt) => {
		const assignment = memberAssignmentTarget(source, stmt.span);
		if (!assignment) {
			return;
		}
		const target = resolveExactMemberCompletion(
			source,
			assignment.member,
			assignment.memberSpan.end,
			memberCtx,
		);
		if (!target || target.writable === undefined) {
			return;
		}
		if (target.writable === false) {
			push(
				'readonlyMemberAssignment',
				`Cannot assign to read-only property '${assignment.label}'.`,
				assignment.memberSpan,
			);
			return;
		}
		const expected = target.writeType ?? target.returns;
		if (assignment.usesSet) {
			if (expected && isKnownScalarType(normalizeType(expected) ?? '')) {
				push(
					'setRequiresObject',
					`Set assignment requires an object-valued target, but '${assignment.label}' expects ${expected}.`,
					assignment.memberSpan,
				);
				return;
			}
			const actual = inferArgumentType(
				assignment.valueTokens,
				stmt.span.start,
				env,
				moduleSignatures,
				sourceNames,
				source,
				memberCtx,
				resolveExpressionType,
				resolveQualifiedExpressionType,
			);
			const reason = objectAssignmentIncompatibilityReason(
				expected,
				actual,
				memberCtx,
			);
			if (reason) {
				push(
					'assignmentObjectTypeMismatch',
					`Object assignment to '${assignment.label}' expects ${expected}, but got ${actual?.label}. ${reason}`,
					actual?.span ?? assignment.memberSpan,
				);
			}
			return;
		}
		if (isKnownObjectAssignmentType(expected, memberCtx)) {
			push(
				'setRequired',
				`Object assignment to '${assignment.label}' requires Set because it expects ${expected}.`,
				assignment.memberSpan,
			);
			return;
		}
		if (!expected || normalizeType(expected) === 'object') {
			return;
		}
		const stringArithmetic = nonnumericStringArithmeticOperand(
			expected,
			assignment.valueTokens,
			stmt.span.start,
		);
		if (stringArithmetic) {
			push(
				'stringArithmeticCoercion',
				`Assignment to '${assignment.label}' expects ${expected}, but this numeric expression contains ${stringArithmetic.label}. This will raise Run-time error '13': Type mismatch.`,
				stringArithmetic.span,
			);
			return;
		}
		const actual = inferArgumentType(
			assignment.valueTokens,
			stmt.span.start,
			env,
			moduleSignatures,
			sourceNames,
			source,
			memberCtx,
			resolveExpressionType,
			resolveQualifiedExpressionType,
		);
		if (!actual) {
			return;
		}
		const reason = incompatibilityReason(expected, actual);
		if (!reason) {
			return;
		}
		push(
			'assignmentTypeMismatch',
			`Assignment to '${assignment.label}' expects ${expected}, but got ${actual.label}. ${reason}`,
			actual.span,
		);
	}, activity);
}

export function checkSetAssignments(
	source: string,
	symbols: ReturnType<typeof buildModuleSymbols>,
	projectVisibleSymbols: readonly VbaSymbol[] | undefined,
	memberCtx: MemberCompletionContext,
	push: PushFn,
): ProcedureStatementVisitor {
	const moduleSignatures = buildModuleTypeSignatures(symbols);
	return (member) => {
		const env = typeEnvironmentFor(symbols, member);
		const sourceNames = sourceNameScopeFor(symbols, member, projectVisibleSymbols);
		const procSym = procedureSymbolFor(symbols, member);
		const resolveExpressionType = (name: string) => declaredValueTypeForSourceBinding(
			symbols,
			procSym,
			projectVisibleSymbols,
			name,
		);
		const resolveQualifiedExpressionType = (qualifier: string, name: string) =>
			declaredValueTypeForQualifiedSourceBinding(
				symbols,
				projectVisibleSymbols,
				qualifier,
				name,
			);
		return (stmt) => {
			const target = setAssignmentTarget(source, stmt.span);
			if (!target) {
				return;
			}
			const targetDeclaredType = declaredTypeForSourceBinding(
				symbols,
				procSym,
				projectVisibleSymbols,
				target.name,
				'assignmentTarget',
			);
			const expected = targetDeclaredType.resolved
				? targetDeclaredType.asType
				: env.get(target.name.toLowerCase());
			const targetType = normalizeType(expected);
			if (!targetType || !isKnownScalarType(targetType)) {
				if (!isKnownObjectAssignmentType(expected, memberCtx)) {
					return;
				}
				const actual = inferArgumentType(
					target.valueTokens,
					stmt.span.start,
					env,
					moduleSignatures,
					sourceNames,
					source,
					memberCtx,
					resolveExpressionType,
					resolveQualifiedExpressionType,
				);
				const reason = objectAssignmentIncompatibilityReason(
					expected,
					actual,
					memberCtx,
				);
				if (reason) {
					push(
						'assignmentObjectTypeMismatch',
						`Object assignment to '${target.name}' expects ${expected}, but got ${actual?.label}. ${reason}`,
						actual?.span ?? target.span,
					);
				}
				return;
			}
			push(
				'setRequiresObject',
				`Set assignment requires an object variable, but '${target.name}' is declared as ${expected}.`,
				target.span,
			);
		};
	};
}

/**
 * Rule: the target of a `Mid`/`Mid$`/`MidB`/`MidB$` replacement statement
 * (MS-VBAL §5.4.3.4) must be a writable String variable. A string-literal
 * target - `Mid$("abc", 2, 3) = "XY"` - is a compile error (oracle-verified
 * `mid_stmt_literal_target_probe`, `mid_stmt_no_suffix_literal_target_probe`,
 * `midb_stmt_literal_target_probe`; the variable-target control
 * `mid_stmt_variable_target_probe` is accepted).
 *
 * No-FP scope: fires only when a statement's first executable token is
 * `mid`/`midb` (optionally followed by the `$` type-character), immediately
 * followed by `(`, whose matching `)` is followed by a top-level `=`, and whose
 * first argument slot is exactly one string-literal token.
 *
 * Conservative shadowing guard: if the module names `mid`/`midb` in ANY
 * declaration form - a symbol-table entry (Dim/Static/Const/parameter/Function/
 * Property) or an implicit `ReDim` array declaration - the rule stays silent for
 * the whole module. The oracle confirms no shadow form makes a literal-target
 * Mid valid (`mid_shadow_dim_array_probe`, `mid_shadow_redim_implicit_probe`,
 * `mid_shadow_property_let_probe` all compile-error), so this guard is belt-and-
 * suspenders: it cannot prevent a real false positive, but it keeps the rule from
 * touching any module that even mentions a user `Mid` without a binder.
 */
export function checkMidStatementLiteralTarget(
	source: string,
	mod: ModuleNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	if (moduleShadowsMidIntrinsic(symbols) || moduleRedimDeclaresMidIntrinsic(source, mod, activity)) {
		return;
	}
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		forEachStatement(member.body, (stmt) => {
			const hit = midStatementLiteralTargetViolation(source, stmt.span);
			if (hit) {
				push('midStatementLiteralTarget', hit.message, hit.span);
			}
		}, activity);
	}
}

/** Suffix-stripped, lower-cased word for a token (keyword or identifier). */
function midBaseWord(tok: VbaToken | undefined): string {
	if (!tok) {
		return '';
	}
	return (tokenName(tok) ?? tok.rawText).toLowerCase().replace(/[$%&!#@]$/, '');
}

/** True when a module declares any symbol that shadows the Mid/MidB intrinsic. */
function moduleShadowsMidIntrinsic(symbols: ReturnType<typeof buildModuleSymbols>): boolean {
	return symbols.all.some((sym) => {
		const base = sym.name.toLowerCase().replace(/[$%&!#@]$/, '');
		return base === 'mid' || base === 'midb';
	});
}

/**
 * True when a module implicitly declares an array named `mid`/`midb` via a
 * `ReDim` with no prior `Dim` (the symbol builder models declared variable groups
 * only, so a ReDim-only name is absent from `symbols.all`).
 */
function moduleRedimDeclaresMidIntrinsic(
	source: string,
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
): boolean {
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		let found = false;
		forEachStatement(member.body, (stmt) => {
			if (found) {
				return;
			}
			const toks = statementTokensAfterLeadingLabel(source, stmt.span);
			if (midBaseWord(toks[0]) !== 'redim') {
				return;
			}
			const start = midBaseWord(toks[1]) === 'preserve' ? 2 : 1;
			for (const group of splitTopLevelTokenGroups(toks, start, ',')) {
				const base = midBaseWord(group[0]);
				if (base === 'mid' || base === 'midb') {
					found = true;
					return;
				}
			}
		}, activity);
		if (found) {
			return true;
		}
	}
	return false;
}

function midStatementLiteralTargetViolation(
	source: string,
	span: Span,
): { span: Span; message: string } | undefined {
	const toks = statementTokensAfterLeadingLabel(source, span);
	if (toks.length === 0) {
		return undefined;
	}
	// Strip a trailing type-character so both lexings of `Mid$` are handled: a
	// single `Mid$` token, or `Mid` followed by a separate `$` token (below).
	const head = midBaseWord(toks[0]);
	if (head !== 'mid' && head !== 'midb') {
		return undefined;
	}
	let parenIndex = 1;
	if (toks[parenIndex]?.rawText === '$') {
		parenIndex = 2;
	}
	if (toks[parenIndex]?.rawText !== '(') {
		return undefined;
	}
	const close = matchParenFrom(toks, parenIndex);
	if (close <= parenIndex + 1) {
		return undefined; // empty or unbalanced argument list
	}
	// The Mid replacement-statement form: the matching `)` is followed by `=`.
	if (toks[close + 1]?.rawText !== '=') {
		return undefined;
	}
	const argToks = toks.slice(parenIndex + 1, close).filter((tok) => tok.kind !== 'comment');
	const slots = splitTopLevelTokenGroups(argToks, 0, ',');
	const target = slots[0];
	if (!target || target.length !== 1 || target[0].kind !== 'stringLiteral') {
		return undefined; // target is not exactly one string literal
	}
	return {
		span: { start: span.start + target[0].start, end: span.start + target[0].end },
		message:
			"The target of a Mid statement must be a writable String variable, not a " +
			'string literal. Assigning into a literal is a compile error.',
	};
}
