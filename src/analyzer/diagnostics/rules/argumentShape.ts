// Rule: argument-shape-mismatch (XLIDE v2.5.0).
//
// A bare array variable, or a same-module user-defined `Type` (struct) value,
// passed where a parameter is a scalar - or, conversely, a scalar (including a
// Variant) passed where a parameter is declared an array - is a VBE compile
// error. This decides purely on declared SHAPE (array-ness / UDT-ness), never on
// element-type coercion.
//
// Oracle-verified rejected at COMPILE:
//   - argshape_array_to_scalar_byref_compile  ("ByRef argument type mismatch")
//   - argshape_array_to_scalar_byval_compile  ("Type mismatch")
//   - argshape_udt_to_scalar_byref_compile    ("ByRef argument type mismatch")
//   - argshape_scalar_to_array_param_compile  ("array or user-defined type expected")
//   - argshape_variant_scalar_to_array_param_compile (same; even a Variant scalar
//     is rejected by shape)
// with the accepted controls argshape_array_to_variant_param_control (an array
// boxes into a Variant parameter), argshape_udt_to_udt_match_control,
// argshape_array_to_array_match_control, and argshape_paramarray_mixed_types_control
// all staying quiet.
//
// No-false-positive discipline: fires only on a single bare identifier argument
// whose declared shape RESOLVES (resolved !== false) to a provable array or a
// same-module `Type`, against a parameter whose shape is the proven-incompatible
// one. Quiet on: Variant parameters (an array/UDT boxes), matching array / UDT
// parameters, ParamArray parameters (Variant, accept anything), object/class
// arguments, member access / indexed `a(i)` / call / parenthesised / multi-token
// arguments, and unresolved or ambiguous names. Disjoint from
// `byref-argument-type-mismatch`: when that rule would fire for the slot (a ByRef
// scalar parameter whose element type mismatches) this rule defers, so the two
// never double-report.

import type { MemberCompletionContext } from '../../completion/memberAccess';
import type { CallArguments, CallableParamType, CallableTypeSignature } from '../callExtraction';
import { extractCall, extractQualifiedCall } from '../callExtraction';
import type { buildModuleSymbols } from '../../symbols/buildModuleSymbols';
import type { VbaProcedureSignature, VbaSymbol } from '../../symbols/symbolModel';
import { procedureSymbolFor, type PushFn } from '../analysisContext';
import { tokenName } from '../../lexer/tokenHelpers';
import type { VbaToken } from '../../lexer/tokenKinds';
import type { Span } from '../../parser/nodes';
import {
	byRefVariableTypeMismatch,
	callableSignatureForCall,
	callableTypeSignaturesFor,
	declaredShapeForSourceBinding,
	declaredValueTypeForQualifiedSourceBinding,
	declaredValueTypeForSourceBinding,
	expressionCalls,
	isKnownScalarType,
	namedArgumentSlot,
	normalizeType,
	sameModuleTypeNames,
	sourceNameScopeFor,
	typeEnvironmentFor,
	type SourceDeclaredShape,
	type SourceDeclaredTypeResolver,
	type SourceQualifiedDeclaredTypeResolver,
} from '../typeInference';
import { matchParenFrom, stripHeaderBrackets, type ProcedureStatementVisitor } from '../walker';

/** Per-statement rule: rides the shared procedure-statement walk (audit #0). */
export function checkArgumentShape(
	source: string,
	symbols: ReturnType<typeof buildModuleSymbols>,
	projectProcedures: ReadonlyMap<string, readonly VbaProcedureSignature[]> | undefined,
	projectVisibleSymbols: readonly VbaSymbol[] | undefined,
	_memberCtx: MemberCompletionContext,
	push: PushFn,
): ProcedureStatementVisitor {
	const moduleSignatures = callableTypeSignaturesFor(symbols, projectProcedures);
	const udtNames = sameModuleTypeNames(symbols);
	return (member) => {
		const env = typeEnvironmentFor(symbols, member);
		const sourceNames = sourceNameScopeFor(symbols, member, projectVisibleSymbols);
		const procSym = procedureSymbolFor(symbols, member);
		const resolveType: SourceDeclaredTypeResolver = (name) =>
			declaredValueTypeForSourceBinding(symbols, procSym, projectVisibleSymbols, name);
		const resolveQualifiedType: SourceQualifiedDeclaredTypeResolver = (qualifier, name) =>
			declaredValueTypeForQualifiedSourceBinding(symbols, projectVisibleSymbols, qualifier, name);
		const resolveShape = (name: string): SourceDeclaredShape =>
			declaredShapeForSourceBinding(symbols, procSym, projectVisibleSymbols, name, 'expression');
		return (stmt) => {
			// `Call PutD((d))` is found both as an expression call and as the
			// statement's call; report each argument once.
			const reported = new Set<string>();
			const pushOnce: PushFn = (code, message, span, data) => {
				const key = `${span.start}:${span.end}:${message}`;
				if (!reported.has(key)) {
					reported.add(key);
					push(code, message, span, data);
				}
			};
			const checkCall = (call: CallArguments): void => {
				const sig = callableSignatureForCall(call, moduleSignatures, sourceNames);
				if (!sig || sig.params.length === 0) {
					return;
				}
				validateArgumentShapes(
					sig,
					call,
					udtNames,
					env,
					resolveType,
					resolveQualifiedType,
					resolveShape,
					pushOnce,
				);
			};
			for (const call of expressionCalls(source, stmt.span, moduleSignatures, sourceNames)) {
				checkCall(call);
			}
			const statementCall =
				extractCall(source, stmt.span) ??
				extractQualifiedCall(source, stmt.span, moduleSignatures);
			if (statementCall) {
				checkCall(statementCall);
			}
		};
	};
}

function validateArgumentShapes(
	sig: CallableTypeSignature,
	call: CallArguments,
	udtNames: ReadonlySet<string>,
	env: ReadonlyMap<string, string>,
	resolveType: SourceDeclaredTypeResolver,
	resolveQualifiedType: SourceQualifiedDeclaredTypeResolver,
	resolveShape: (name: string) => SourceDeclaredShape,
	push: PushFn,
): void {
	// Slot -> parameter pairing mirrors validateArgumentTypesForSignature (named
	// argument -> by name, otherwise positional by index). Kept local because the
	// shape rule treats ParamArray as always-accepting, so it never needs that
	// function's element-absorption logic.
	const paramsByName = new Map(
		sig.params.map((p) => [stripHeaderBrackets(p.name).toLowerCase(), p]),
	);
	let positionalIndex = 0;
	for (let i = 0; i < call.slots.length; i++) {
		const named = namedArgumentSlot(call.slots[i]);
		let param: CallableParamType | undefined;
		let valueSlot = call.slots[i];
		if (named) {
			param = paramsByName.get(named.name.toLowerCase());
			valueSlot = named.value;
		} else {
			param = sig.params[Math.min(positionalIndex, sig.params.length - 1)];
			if (!param || (positionalIndex >= sig.params.length && !param.paramArray)) {
				continue;
			}
			positionalIndex++;
		}
		if (!param || param.paramArray) {
			// ParamArray parameters are Variant and accept any shape (oracle: accepted).
			continue;
		}
		// Disjoint from byref-argument-type-mismatch: defer to it whenever it owns
		// this slot (a ByRef scalar parameter whose element type mismatches), so the
		// two rules never double-report on the same argument.
		if (
			byRefVariableTypeMismatch(
				param,
				valueSlot,
				call.sliceStart,
				env,
				resolveType,
				resolveQualifiedType,
			)
		) {
			continue;
		}
		if (param.isArray) {
			// `PutD (d)` as a statement passes `(d)`, a value (issue #218).
			const problem = call.argumentsParenthesized && call.slots.length === 1
				? parenthesizedArgument(valueSlot, call.sliceStart)
				: arrayArgumentProblem(valueSlot, call.sliceStart, param, resolveShape);
			if (problem) {
				push('argumentShapeMismatch', `${problem.what}, but parameter '${param.name}' of '${sig.name}' is an array of ${param.type ?? 'Variant'}. This is a VBE compile error: Type mismatch: array or user-defined type expected.`, problem.span);
				continue;
			}
		}
		const ident = soleIdentifier(valueSlot, call.sliceStart);
		if (!ident) {
			continue;
		}
		const shape = resolveShape(ident.name);
		if (!shape.resolved || !shape.shape) {
			continue; // unresolved / ambiguous -> quiet
		}
		if (shape.shape.isArray) {
			if (!param.isArray && paramIsKnownScalar(param)) {
				push('argumentShapeMismatch', arrayToScalarMessage(ident.name, param, sig.name), ident.span);
			}
			continue;
		}
		const asType = shape.shape.asType;
		if (asType && udtNames.has(asType.toLowerCase())) {
			if (!param.isArray && paramIsKnownScalar(param)) {
				push(
					'argumentShapeMismatch',
					udtToScalarMessage(ident.name, asType, param, sig.name),
					ident.span,
				);
			}
			continue;
		}
		if (param.isArray && asType && isScalarOrVariant(asType)) {
			push('argumentShapeMismatch', scalarToArrayMessage(ident.name, param, sig.name), ident.span);
		}
	}
}

/**
 * An array parameter takes an array variable of its own element type, and
 * nothing else (issue #216, measured in Excel 16.0): a function's result,
 * `Split(...)` or `Array(...)`, is refused, and so is an array of String
 * for an array of Variant.
 */
function arrayArgumentProblem(
	slot: readonly VbaToken[],
	sliceStart: number,
	param: CallableParamType,
	resolveShape: (name: string) => SourceDeclaredShape,
): { what: string; span: Span } | undefined {
	const toks = slot.filter((t) => t.kind !== 'comment' && t.kind !== 'newline');
	// `(d)` is a value, even around an array of the right type: IRR((d))
	// with d a Double array is refused (issue #218, measured).
	if (toks[0]?.rawText === '(' && matchParenFrom(toks, 0) === toks.length - 1) {
		return parenthesizedArgument(toks.slice(1, -1), sliceStart);
	}
	const name = toks[0] ? tokenName(toks[0]) : undefined;
	if (!name) {
		return undefined;
	}
	const span = { start: sliceStart + toks[0].start, end: sliceStart + toks[toks.length - 1].end };
	const shape = resolveShape(name);
	if (toks.length > 1 && toks[1].rawText === '(' && toks[toks.length - 1].rawText === ')') {
		// The result of VBA's Split or Array, which no project name shadows.
		if (!shape.resolved && /^(split|array)$/i.test(name)) {
			return { what: `'${name}(...)' is a function's result, not an array variable`, span };
		}
		return undefined;
	}
	if (toks.length !== 1 || !shape.resolved || !shape.shape?.isArray) {
		return undefined;
	}
	const element = normalizeType(shape.shape.asType) ?? 'variant';
	const expected = normalizeType(param.type) ?? 'variant';
	if (element !== expected && (isKnownScalarType(element) || element === 'variant') && (isKnownScalarType(expected) || expected === 'variant')) {
		return { what: `'${name}' is an array of ${shape.shape.asType ?? 'Variant'}`, span };
	}
	return undefined;
}

/** The problem with an argument in parentheses, given the tokens inside them. */
function parenthesizedArgument(inner: readonly VbaToken[], sliceStart: number): { what: string; span: Span } | undefined {
	const toks = inner.filter((t) => t.kind !== 'comment' && t.kind !== 'newline');
	if (toks.length === 0) {
		return undefined;
	}
	return {
		what: `'(${toks.map((t) => t.rawText).join('')})' is in parentheses, which pass a value rather than an array variable`,
		span: { start: sliceStart + toks[0].start, end: sliceStart + toks[toks.length - 1].end },
	};
}

/** A single bare identifier argument (not indexed / member / call / expression). */
function soleIdentifier(
	slot: readonly VbaToken[],
	sliceStart: number,
): { name: string; span: Span } | undefined {
	const toks = slot.filter((t) => t.kind !== 'comment' && t.kind !== 'newline');
	if (toks.length !== 1) {
		return undefined;
	}
	const name = tokenName(toks[0]);
	if (!name) {
		return undefined;
	}
	return { name, span: { start: sliceStart + toks[0].start, end: sliceStart + toks[0].end } };
}

/** True when a parameter is a known scalar (excludes Variant, array, object, UDT). */
function paramIsKnownScalar(param: CallableParamType): boolean {
	const norm = normalizeType(param.type);
	return !!norm && isKnownScalarType(norm);
}

/** True when a declared type is a known scalar or `Variant` (the array-param reject set). */
function isScalarOrVariant(asType: string): boolean {
	const norm = normalizeType(asType);
	return !!norm && (isKnownScalarType(norm) || norm === 'variant');
}

function vbeScalarError(param: CallableParamType): string {
	return param.byRef ? 'ByRef argument type mismatch' : 'Type mismatch';
}

function arrayToScalarMessage(name: string, param: CallableParamType, callee: string): string {
	return `Argument '${name}' is declared as an array, but parameter '${param.name}' of '${callee}' expects a scalar ${param.type}. This is a VBE compile error: ${vbeScalarError(param)}.`;
}

function udtToScalarMessage(
	name: string,
	asType: string,
	param: CallableParamType,
	callee: string,
): string {
	return `Argument '${name}' is declared As ${asType} (a user-defined Type), but parameter '${param.name}' of '${callee}' expects a scalar ${param.type}. This is a VBE compile error: ${vbeScalarError(param)}.`;
}

function scalarToArrayMessage(name: string, param: CallableParamType, callee: string): string {
	return `Argument '${name}' is a scalar, but parameter '${param.name}' of '${callee}' is declared as an array. This is a VBE compile error: Type mismatch: array or user-defined type expected.`;
}
