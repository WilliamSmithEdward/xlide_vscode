// Rule family: deterministic runtime argument/conversion values (audit #0).
//
// Extracted verbatim from analyzeModule.ts: constant arguments that are
// provably outside a runtime function's accepted range and conversions of
// provably invalid literals.

import type { ConditionalActivityTracker } from '../../conditional/conditionalCompilation';
import {
	evaluateIntegerConstantExpression,
	type IntegerConstantLookup,
	parseVbaIntegerLiteral,
	resolveRawIntegerConstants,
} from '../../constants/integerConstantExpression';
import type { VbaToken } from '../../lexer/tokenKinds';
import type {
	ModuleNode,
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
	type CallableTypeSignature,
	emptyArgSplit,
	splitArgSlots,
} from '../callExtraction';
import {
	collectBodyLiteralIntegerConstants,
	collectModuleLiteralIntegerConstants,
} from '../constExpr';
import { isBareOrVbaQualifiedIntrinsicCall } from '../rules/shared';
import {
	callableTypeSignaturesFor,
	namedArgumentSlot,
	runtimeCallableSourceShadowed,
	scopedIntegerConstantLookup,
	type SourceNameScope,
	sourceNameScopeFor,
	stringLiteralValue,
	typeEnvironmentFor,
	unwrapOuterParens,
} from '../typeInference';
import {
	matchParenFrom,
	statementTokens,
	tokenName,
	tokenText,
	type ProcedureStatementVisitor,
} from '../walker';

interface RuntimeArgumentValueSpec {
	canonicalName: 'Left' | 'Right' | 'String' | 'Space' | 'Mid' | 'Replace' | 'InStr' | 'Chr' | 'ChrW';
	parameterName: string;
	argumentIndex: number;
	minimum?: number;
	maximum?: number;
	minimumSlotCount?: number;
	allowNamed?: boolean;
}

interface RuntimeArgumentValueHit {
	displayName: string;
	parameterName: string;
	value: number;
	span: Span;
}

/**
 * Rule: some runtime-library arguments have deterministic value bounds even
 * when the argument type itself is valid. This slice is VBE-oracle-backed for
 * integer bounds on selected string runtime functions, which compile but raise
 * Run-time error 5 when the value is outside the proven range.
 */
export function checkRuntimeArgumentValues(
	source: string,
	mod: ModuleNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
	projectProcedures: ReadonlyMap<string, readonly VbaProcedureSignature[]> | undefined,
	projectIntegerConstants: ReadonlyMap<string, string | undefined> | undefined,
	projectVisibleSymbols: readonly VbaSymbol[] | undefined,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): ProcedureStatementVisitor {
	const moduleSignatures = callableTypeSignaturesFor(symbols, projectProcedures);
	const projectConstants = resolveRawIntegerConstants(projectIntegerConstants ?? new Map(), new Map());
	const moduleConstants = collectModuleLiteralIntegerConstants(mod, activity, projectConstants);
	return (member) => {
		const env = typeEnvironmentFor(symbols, member);
		const sourceNames = sourceNameScopeFor(symbols, member, projectVisibleSymbols);
		const procedureConstants = new Map(moduleConstants);
		collectBodyLiteralIntegerConstants(member.body, procedureConstants, activity);
		const procSym = procedureSymbolFor(symbols, member);
		const constants = scopedIntegerConstantLookup(
			procedureConstants,
			symbols,
			procSym,
			projectVisibleSymbols,
		);
		return (stmt) => {
			for (const hit of runtimeArgumentValueHits(source, stmt.span, moduleSignatures, env, constants, sourceNames)) {
				push(
					'runtimeArgumentValue',
					`Argument '${hit.parameterName}' of '${hit.displayName}' is ${hit.value}; this will raise Run-time error '5': Invalid procedure call or argument.`,
					hit.span,
				);
			}
		};
	};
}

function runtimeArgumentValueHits(
	source: string,
	span: Span,
	moduleSignatures: ReadonlyMap<string, CallableTypeSignature>,
	env: ReadonlyMap<string, string>,
	constants: IntegerConstantLookup,
	sourceNames: SourceNameScope,
): RuntimeArgumentValueHit[] {
	const toks = statementTokens(source, span);
	if (isDeclarationLikeStatement(toks)) {
		return [];
	}
	const hits: RuntimeArgumentValueHit[] = [];
	for (let i = 0; i < toks.length - 1; i++) {
		const call = runtimeArgumentValueCallAt(toks, i, span, moduleSignatures, env, sourceNames);
		if (!call) {
			continue;
		}
		for (const spec of call.specs) {
			const slot = runtimeArgumentValueSlot(call.slots, spec);
			const literal = slot
				? integerArgumentOutsideBounds(source, slot, span.start, spec, constants)
				: undefined;
			if (!literal) {
				continue;
			}
			hits.push({
				displayName: call.displayName,
				parameterName: spec.parameterName,
				value: literal.value,
				span: literal.span,
			});
		}
	}
	return hits;
}

function runtimeArgumentValueCallAt(
	toks: readonly VbaToken[],
	index: number,
	span: Span,
	moduleSignatures: ReadonlyMap<string, CallableTypeSignature>,
	env: ReadonlyMap<string, string>,
	sourceNames: SourceNameScope,
): {
	displayName: string;
	specs: readonly RuntimeArgumentValueSpec[];
	slots: VbaToken[][];
} | undefined {
	const name = tokenName(toks[index]);
	if (!name) {
		return undefined;
	}
	// Only a bare `Left(...)` or a genuine `VBA.Left(...)` is the intrinsic. The
	// shared helper also rejects `obj.vba.Left(...)` via the third-token ('.')
	// check that this ad-hoc logic previously omitted.
	if (!isBareOrVbaQualifiedIntrinsicCall(toks, index)) {
		return undefined;
	}
	// A bare call can be shadowed by a source symbol; a `VBA.`-qualified one
	// cannot, so only the bare form participates in the shadow gate below.
	const qualifier = index >= 2 && toks[index - 1].rawText === '.'
		? tokenName(toks[index - 2])
		: undefined;

	let parenIndex = index + 1;
	let suffix = '';
	if (isRuntimeStringFunctionSuffix(toks[parenIndex])) {
		suffix = toks[parenIndex].rawText;
		parenIndex++;
	}
	if (toks[parenIndex]?.rawText !== '(') {
		return undefined;
	}

	const specs = runtimeArgumentValueSpecs(name);
	if (specs.length === 0) {
		return undefined;
	}
	if (suffix && !runtimeArgumentValueAllowsStringSuffix(specs[0].canonicalName)) {
		return undefined;
	}
	const lower = specs[0].canonicalName.toLowerCase();
	if (!qualifier && (
		moduleSignatures.has(lower) ||
		env.has(lower) ||
		runtimeCallableSourceShadowed(name, sourceNames)
	)) {
		return undefined;
	}

	const close = matchParenFrom(toks, parenIndex);
	if (close < 0) {
		return undefined;
	}
	const inner = toks.slice(parenIndex + 1, close);
	const split = inner.length === 0 ? emptyArgSplit() : splitArgSlots(inner, span.start);
	return {
		displayName: `${specs[0].canonicalName}${suffix}`,
		specs,
		slots: split.slots,
	};
}

function runtimeArgumentValueSpecs(name: string): readonly RuntimeArgumentValueSpec[] {
	switch (name.toLowerCase()) {
		case 'left':
			return [{ canonicalName: 'Left', parameterName: 'Length', argumentIndex: 1, minimum: 0 }];
		case 'right':
			return [{ canonicalName: 'Right', parameterName: 'Length', argumentIndex: 1, minimum: 0 }];
		case 'string':
			return [{ canonicalName: 'String', parameterName: 'Number', argumentIndex: 0, minimum: 0 }];
		case 'space':
			return [{ canonicalName: 'Space', parameterName: 'Number', argumentIndex: 0, minimum: 0 }];
		case 'mid':
			return [
				{ canonicalName: 'Mid', parameterName: 'Start', argumentIndex: 1, minimum: 1 },
				{ canonicalName: 'Mid', parameterName: 'Length', argumentIndex: 2, minimum: 0 },
			];
		case 'replace':
			return [
				{ canonicalName: 'Replace', parameterName: 'Start', argumentIndex: 3, minimum: 1 },
				{ canonicalName: 'Replace', parameterName: 'Count', argumentIndex: 4, minimum: -1 },
			];
		case 'instr':
			return [
				{
					canonicalName: 'InStr',
					parameterName: 'Start',
					argumentIndex: 0,
					minimum: 1,
					minimumSlotCount: 3,
					allowNamed: false,
				},
			];
		case 'chr':
			return [{ canonicalName: 'Chr', parameterName: 'CharCode', argumentIndex: 0, minimum: 0, maximum: 255 }];
		case 'chrw':
			return [{ canonicalName: 'ChrW', parameterName: 'CharCode', argumentIndex: 0, maximum: 65535 }];
		default:
			return [];
	}
}

function runtimeArgumentValueAllowsStringSuffix(name: RuntimeArgumentValueSpec['canonicalName']): boolean {
	return name === 'Left' || name === 'Right' || name === 'String' || name === 'Space' || name === 'Mid';
}

function runtimeArgumentValueSlot(
	slots: readonly VbaToken[][],
	spec: RuntimeArgumentValueSpec,
): VbaToken[] | undefined {
	if (spec.minimumSlotCount !== undefined && slots.length < spec.minimumSlotCount) {
		return undefined;
	}
	let positionalIndex = 0;
	for (const slot of slots) {
		const named = namedArgumentSlot(slot);
		if (named) {
			if (spec.allowNamed === false) {
				continue;
			}
			if (named.name.toLowerCase() === spec.parameterName.toLowerCase()) {
				return named.value;
			}
			continue;
		}
		if (positionalIndex === spec.argumentIndex) {
			return slot;
		}
		positionalIndex++;
	}
	return undefined;
}

function integerArgumentOutsideBounds(
	source: string,
	slot: readonly VbaToken[],
	sliceStart: number,
	spec: RuntimeArgumentValueSpec,
	constants: IntegerConstantLookup,
): { value: number; span: Span } | undefined {
	const toks = unwrapOuterParens(
		slot.filter((t) => t.kind !== 'comment' && t.kind !== 'newline'),
	);
	if (toks.length === 0) {
		return undefined;
	}
	let sign = 1;
	let literal = toks[0];
	let start = literal?.start;
	let literalValue: number | undefined;
	const signedLiteral = toks.length === 2 && (toks[0].rawText === '-' || toks[0].rawText === '+');
	if (signedLiteral) {
		sign = toks[0].rawText === '-' ? -1 : 1;
		literal = toks[1];
		start = toks[0].start;
	}
	if (literal?.kind === 'integerLiteral' && start !== undefined && (toks.length === 1 || signedLiteral)) {
		const rawValue = parseVbaIntegerLiteral(literal.rawText);
		if (rawValue !== undefined) {
			literalValue = sign * rawValue;
		}
	}
	if (literalValue !== undefined) {
		if (integerArgumentValueInBounds(literalValue, spec)) {
			return undefined;
		}
		return {
			value: literalValue,
			span: { start: sliceStart + start!, end: sliceStart + literal.end },
		};
	}

	const expressionValue = evaluateIntegerConstantExpression(
		source.slice(sliceStart + toks[0].start, sliceStart + toks[toks.length - 1].end),
		constants,
	);
	if (expressionValue === undefined || integerArgumentValueInBounds(expressionValue, spec)) {
		return undefined;
	}
	return {
		value: expressionValue,
		span: { start: sliceStart + toks[0].start, end: sliceStart + toks[toks.length - 1].end },
	};
}

function integerArgumentValueInBounds(
	value: number,
	spec: RuntimeArgumentValueSpec,
): boolean {
	if (spec.minimum !== undefined && value < spec.minimum) {
		return false;
	}
	if (spec.maximum !== undefined && value > spec.maximum) {
		return false;
	}
	return true;
}

function isRuntimeStringFunctionSuffix(tok: VbaToken | undefined): boolean {
	return tok?.rawText === '$';
}

function isDeclarationLikeStatement(toks: readonly VbaToken[]): boolean {
	const first = tokenText(toks[0]);
	switch (first) {
		case 'dim':
		case 'static':
		case 'const':
		case 'private':
		case 'public':
		case 'friend':
		case 'declare':
		case 'sub':
		case 'function':
		case 'property':
		case 'type':
		case 'enum':
			return true;
		default:
			return false;
	}
}

interface RuntimeConversionValueHit {
	displayName: string;
	name: string;
	span: Span;
}

/**
 * Rule: selected conversion functions compile with Variant-like arguments but
 * can deterministically fail at runtime for literal values that cannot be
 * converted. This first slice is intentionally narrow for CDate string
 * literals that are plainly non-date text.
 */
export function checkRuntimeConversionValues(
	source: string,
	symbols: ReturnType<typeof buildModuleSymbols>,
	projectVisibleSymbols: readonly VbaSymbol[] | undefined,
	push: PushFn,
): ProcedureStatementVisitor {
	return (member) => {
		const sourceNames = sourceNameScopeFor(symbols, member, projectVisibleSymbols);
		return (stmt) => {
			for (const hit of runtimeConversionValueHits(source, stmt.span, sourceNames)) {
				push(
					'runtimeConversionValue',
					`${hit.displayName} cannot convert ${hit.name} to Date. This will raise Run-time error '13': Type mismatch.`,
					hit.span,
				);
			}
		};
	};
}

function runtimeConversionValueHits(
	source: string,
	span: Span,
	sourceNames: SourceNameScope,
): RuntimeConversionValueHit[] {
	const toks = statementTokens(source, span);
	if (isDeclarationLikeStatement(toks)) {
		return [];
	}
	const hits: RuntimeConversionValueHit[] = [];
	for (let i = 0; i < toks.length - 2; i++) {
		const name = tokenName(toks[i]);
		if (!name || name.toLowerCase() !== 'cdate') {
			continue;
		}
		if (toks[i + 1]?.rawText !== '(' || !isBareOrVbaQualifiedIntrinsicCall(toks, i)) {
			continue;
		}
		const qualified = toks[i - 1]?.rawText === '.';
		if (!qualified && runtimeCallableSourceShadowed(name, sourceNames)) {
			continue;
		}
		const close = matchParenFrom(toks, i + 1);
		if (close < 0) {
			continue;
		}
		const split = splitArgSlots(toks.slice(i + 2, close), span.start);
		const firstSlot = split.slots[0] ?? [];
		if (firstSlot.length !== 1 || firstSlot[0].kind !== 'stringLiteral') {
			continue;
		}
		const value = stringLiteralValue(firstSlot[0].rawText);
		if (!isDefinitelyInvalidDateString(value)) {
			continue;
		}
		hits.push({
			displayName: qualified ? `VBA.${name}` : name,
			name: firstSlot[0].rawText,
			span: split.spans[0] ?? { start: span.start + firstSlot[0].start, end: span.start + firstSlot[0].end },
		});
	}
	return hits;
}

function isDefinitelyInvalidDateString(value: string): boolean {
	const trimmed = value.trim();
	if (trimmed.length === 0) {
		return true;
	}
	if (/[0-9]/.test(trimmed) || /[^\x00-\x7F]/.test(trimmed)) {
		return false;
	}
	if (!/^[A-Za-z\s]+$/.test(trimmed)) {
		return false;
	}
	return !/\b(?:jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|jun(?:e)?|jul(?:y)?|aug(?:ust)?|sep(?:t(?:ember)?)?|oct(?:ober)?|nov(?:ember)?|dec(?:ember)?)\b/i.test(trimmed);
}
