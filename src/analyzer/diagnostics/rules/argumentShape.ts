// Rule: argument-shape-mismatch (XLIDE v2.5.0).
//
// A bare array variable, or a same-module user-defined `Type` (struct) value,
// passed where a parameter is a scalar — or, conversely, a scalar (including a
// Variant) passed where a parameter is declared an array — is a VBE compile
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
import { stripHeaderBrackets, type ProcedureStatementVisitor } from '../walker';

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
					push,
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
