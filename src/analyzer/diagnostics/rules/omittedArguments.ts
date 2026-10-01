// Rule family: an argument a call leaves out, read by the procedure it calls
// (issue #260).
//
// An omitted Optional Variant is Missing in the callee, a Variant holding
// error 448, and an omitted typed Optional holds its default: 0, "", or the
// value the declaration gives. A ParamArray holds what the call passed,
// 0-based whatever Option Base says. Each is known only with the call and
// the procedure together, so this reads them at the call: `Opt()` against
// `Opt = x + 1` raises 13, `Total()` against `Total = args(0)` raises 9.
// Every case was measured in Excel 16.0.
//
// The callee is read from its first line up to the first use of the
// parameter. Blocks that neither use it nor leave are passed over; anything
// that may jump (Exit, GoTo, Return, End, Resume, On Error, Stop, Error)
// ends the read, and so does a single-line If arm or a block that uses it,
// since that use runs only on some paths. Only a use measured to raise is
// reported: IsMissing, IsError, CStr, CLng and a Variant parameter read a
// Missing value cleanly. Calls into the same module only, where the body is
// known.

import type { ConditionalActivityTracker } from '../../conditional/conditionalCompilation';
import type { VbaToken } from '../../lexer/tokenKinds';
import type { BodyNode, LeafStatementNode, ModuleNode, ParameterNode, ProcedureNode, Span } from '../../parser/nodes';
import { isLeafStatement } from '../../parser/nodes';
import type { buildModuleSymbols } from '../../symbols/buildModuleSymbols';
import type { VbaSymbol } from '../../symbols/symbolModel';
import type { PushFn } from '../analysisContext';
import type { DiagnosticRuleName } from '../ruleMetadata';
import { type CallArguments, extractCall, isNamedSlot } from '../callExtraction';
import { isInvalidNumericString } from '../stringConversion';
import {
	bareCallableSourceShadowed,
	callableTypeSignaturesFor,
	expressionCalls,
	isKnownScalarType,
	isNumericType,
	normalizeType,
	sourceNameScopeFor,
	stringLiteralValue,
	typeEnvironmentFor,
} from '../typeInference';
import {
	activeModuleMembers,
	blockFooterLineSpan,
	blockHeaderLineSpan,
	matchParenFrom,
	rawExpressionTokens,
	statementAndBranchSpans,
	statementTokensAfterLeadingLabel,
	tokenName,
	tokenText,
	type ProcedureStatementVisitor,
} from '../walker';
import { isBareOrVbaQualifiedIntrinsicCall } from './shared';

/** What an omitted parameter holds in the callee. */
type OmittedValue = { kind: 'missing' } | { kind: 'number'; value: number } | { kind: 'string'; value: string };

/** The callee's first use of an omitted parameter, when that use raises. */
interface RaisingUse {
	rule: DiagnosticRuleName;
	/** What the use does, completing "and 'Proc' ...". */
	does: string;
	error: string;
	/** Absolute span of the use in the callee. */
	span: Span;
}

/** A ParamArray read past what the call passes. */
interface ParamArrayRead {
	index: number;
	span: Span;
	/** The tokens `args(k)` covers, for a read of an element the call skips. */
	element?: { toks: readonly VbaToken[]; first: number; last: number; spanStart: number };
}

/** Intrinsics that raise 13 on a Missing first argument (each measured). */
const MISSING_RAISING_INTRINSICS: ReadonlySet<string> = new Set([
	'len', 'lenb', 'left', 'left$', 'right', 'right$', 'mid', 'mid$', 'trim', 'trim$', 'ucase', 'ucase$',
	'lcase', 'lcase$', 'instr', 'int', 'abs', 'fix', 'sgn', 'val', 'format', 'format$', 'hex', 'hex$',
	'chr', 'chr$', 'asc', 'space', 'space$', 'str', 'str$', 'cdate', 'lbound',
]);

/** Intrinsics that raise 13 on a Missing second argument too. */
const MISSING_RAISING_SECOND: ReadonlySet<string> = new Set(['instr']);

/** Binary operators that raise 13 on a Missing operand (each measured). */
const MISSING_RAISING_OPERATORS: ReadonlySet<string> = new Set([
	'+', '-', '*', '/', '\\', '^', '&', '=', '<', '>', '<=', '>=', '<>', 'mod', 'and', 'or', 'like',
]);

/** Conversions that raise 13 on a string that is not a number (CLng("") and CDbl("") measured). */
const NUMBER_CONVERSIONS: ReadonlySet<string> = new Set(['clng', 'cint', 'cdbl', 'csng', 'ccur', 'cbyte', 'clnglng', 'clngptr', 'cdec']);

/** Statement heads after which the read cannot assume it runs on. */
const LEAVING_HEADS: ReadonlySet<string> = new Set(['exit', 'goto', 'gosub', 'return', 'end', 'resume', 'on', 'stop', 'error']);

const SUFFIX_TYPES: Readonly<Record<string, string>> = { '%': 'integer', '&': 'long', '^': 'longlong', '!': 'single', '#': 'double', '@': 'currency', '$': 'string' };

export function checkOmittedArgumentReads(
	source: string,
	mod: ModuleNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
	activity: ConditionalActivityTracker | undefined,
	projectVisibleSymbols: readonly VbaSymbol[] | undefined,
	push: PushFn,
): ProcedureStatementVisitor {
	const procedures = new Map<string, ProcedureNode | null>();
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure' || (member.procKind !== 'Sub' && member.procKind !== 'Function')) {
			continue;
		}
		const lower = member.name.toLowerCase();
		procedures.set(lower, procedures.has(lower) ? null : member);
	}
	if (![...procedures.values()].some((proc) => proc?.params.some((param) => param.optional || param.paramArray))) {
		return () => undefined;
	}
	const moduleSignatures = callableTypeSignaturesFor(symbols, undefined);
	const reads = new Map<string, RaisingUse | ParamArrayRead | undefined>();
	const callee = new CalleeReader(source, symbols, activity, procedures);
	return (member) => {
		const sourceNames = sourceNameScopeFor(symbols, member, projectVisibleSymbols);
		return (stmt) => {
			const calls: CallArguments[] = [];
			const statementCall = extractCall(source, stmt.span);
			if (statementCall) {
				calls.push(statementCall);
			}
			for (const call of expressionCalls(source, stmt.span, moduleSignatures, sourceNames)) {
				if (!calls.some((other) => other.nameSpan.start === call.nameSpan.start)) {
					calls.push(call);
				}
			}
			for (const branch of statementAndBranchSpans(stmt).slice(1)) {
				const branchCall = extractCall(source, branch);
				if (branchCall && !calls.some((other) => other.nameSpan.start === branchCall.nameSpan.start)) {
					calls.push(branchCall);
				}
			}
			for (const call of calls) {
				const proc = call.qualifier ? undefined : procedures.get(call.name.toLowerCase());
				if (!proc || bareCallableSourceShadowed(call.name, sourceNames)) {
					continue;
				}
				for (const omitted of omittedParameters(proc, call)) {
					const key = `${proc.name}|${omitted.param.name}|${omitted.key}`;
					if (!reads.has(key)) {
						reads.set(key, callee.firstRaisingUse(proc, omitted.param, omitted.value, omitted.passed, omitted.skipped));
					}
					const use = reads.get(key);
					if (use) {
						report(source, proc, omitted, use, push);
					}
				}
			}
		};
	};
}

interface Omitted {
	param: ParameterNode;
	/** What the parameter holds; undefined for a ParamArray. */
	value?: OmittedValue;
	/** For a ParamArray: how many values the call passes, and which it skips. */
	passed: number;
	skipped: ReadonlySet<number>;
	/** Where to report: the skipped slot, or the call's name. */
	span: Span;
	key: string;
}

/** The parameters a call leaves to their omitted value, and the ParamArray it fills. */
function omittedParameters(proc: ProcedureNode, call: CallArguments): Omitted[] {
	const params = proc.params;
	const named = call.slots.filter(isNamedSlot);
	const firstNamed = call.slots.findIndex(isNamedSlot);
	const positional = firstNamed < 0 ? call.slots : call.slots.slice(0, firstNamed);
	const arrayAt = params.findIndex((param) => param.paramArray);
	const fixed = arrayAt < 0 ? params.length : arrayAt;
	// A call argument-count refuses is that rule's.
	if ((arrayAt < 0 && positional.length > params.length) || (named.length > 0 && arrayAt >= 0 && positional.length > fixed)) {
		return [];
	}
	const namedLower = new Set(named.map((slot) => slot[0].rawText.replace(/^\[|\]$/g, '').toLowerCase()));
	const out: Omitted[] = [];
	for (let k = 0; k < fixed; k++) {
		const param = params[k];
		const supplied = (k < positional.length && positional[k].length > 0) || namedLower.has(param.name.toLowerCase());
		if (supplied) {
			continue;
		}
		if (!param.optional) {
			return [];
		}
		const value = omittedValue(param);
		if (value) {
			const skippedSlot = k < positional.length;
			out.push({
				param,
				value,
				passed: 0,
				skipped: new Set(),
				span: skippedSlot ? call.slotSpans?.[k] ?? call.nameSpan : call.nameSpan,
				key: JSON.stringify(value),
			});
		}
	}
	if (arrayAt >= 0 && named.length === 0) {
		const rest = positional.slice(fixed);
		const skipped = new Set(rest.flatMap((slot, j) => (slot.length === 0 ? [j] : [])));
		out.push({ param: params[arrayAt], passed: rest.length, skipped, span: call.nameSpan, key: `${rest.length}:${[...skipped].join(',')}` });
	}
	return out;
}

/** What an omitted Optional holds: Missing, its type's default, or the value it declares. */
function omittedValue(param: ParameterNode): OmittedValue | undefined {
	const type = parameterType(param);
	if (param.defaultRaw !== undefined) {
		const toks = rawExpressionTokens(param.defaultRaw).filter((tok) => tok.kind !== 'comment');
		const literal = literalValue(toks);
		if (!literal) {
			return undefined;
		}
		if (type === 'variant') {
			return literal;
		}
		if (type === 'string') {
			return literal.kind === 'string' ? literal : undefined;
		}
		return (isNumericType(type) || type === 'boolean') && literal.kind === 'number' ? literal : undefined;
	}
	if (type === 'variant') {
		return { kind: 'missing' };
	}
	if (type === 'string') {
		return { kind: 'string', value: '' };
	}
	return isNumericType(type) || type === 'boolean' ? { kind: 'number', value: 0 } : undefined;
}

function literalValue(toks: readonly VbaToken[]): OmittedValue | undefined {
	if (toks.length === 1 && toks[0].kind === 'stringLiteral') {
		return { kind: 'string', value: stringLiteralValue(toks[0].rawText) };
	}
	const word = toks.length === 1 ? tokenText(toks[0]) : '';
	if (word === 'true' || word === 'false') {
		return { kind: 'number', value: word === 'true' ? -1 : 0 };
	}
	const negative = toks.length === 2 && toks[0].rawText === '-';
	const number = toks[negative ? 1 : 0];
	if (toks.length !== (negative ? 2 : 1) || (number.kind !== 'integerLiteral' && number.kind !== 'floatLiteral')) {
		return undefined;
	}
	const value = Number(number.rawText.replace(/[%&^!#@]$/, ''));
	return Number.isFinite(value) ? { kind: 'number', value: negative ? -value : value } : undefined;
}

function parameterType(param: ParameterNode): string {
	if (param.typeSuffix) {
		return SUFFIX_TYPES[param.typeSuffix] ?? '';
	}
	return normalizeType(param.asType) ?? 'variant';
}

function report(source: string, proc: ProcedureNode, omitted: Omitted, use: RaisingUse | ParamArrayRead, push: PushFn): void {
	const name = omitted.param.name;
	const where = `line ${lineOf(source, use.span.start)}`;
	if (!('rule' in use)) {
		const bound = omitted.passed - 1;
		push(
			'arraySubscriptOutOfBounds',
			`'${proc.name}' reads ${name}(${use.index}) (${where}), and this call passes ${omitted.passed === 0 ? 'no values' : omitted.passed === 1 ? '1 value' : `${omitted.passed} values`} to that ParamArray, so its upper bound is ${bound}. This will raise Run-time error '9': Subscript out of range.`,
			omitted.span,
		);
		return;
	}
	const value = omitted.value;
	const holds = !value || value.kind === 'missing' ? 'Missing' : value.kind === 'string' ? JSON.stringify(value.value) : String(value.value);
	const subject = omitted.param.paramArray ? `skips an element of '${name}'` : `omits '${name}'`;
	push(
		use.rule,
		`This call ${subject}, so it is ${holds} in '${proc.name}', which ${use.does} (${where}). This will raise Run-time error ${use.error}.`,
		omitted.span,
	);
}

function lineOf(source: string, offset: number): number {
	let line = 1;
	for (let i = 0; i < offset; i++) {
		if (source.charCodeAt(i) === 10) {
			line++;
		}
	}
	return line;
}

/** Reads a callee from its first line to the first use of a parameter. */
class CalleeReader {
	constructor(
		private readonly source: string,
		private readonly symbols: ReturnType<typeof buildModuleSymbols>,
		private readonly activity: ConditionalActivityTracker | undefined,
		private readonly procedures: ReadonlyMap<string, ProcedureNode | null>,
	) {}

	firstRaisingUse(
		proc: ProcedureNode,
		param: ParameterNode,
		value: OmittedValue | undefined,
		passed: number,
		skipped: ReadonlySet<number>,
	): RaisingUse | ParamArrayRead | undefined {
		const lower = param.name.toLowerCase();
		const use = this.firstUse(proc.body, lower);
		if (!use) {
			return undefined;
		}
		const { toks, index, spanStart } = use;
		if (param.paramArray) {
			return this.paramArrayRead(proc, toks, index, spanStart, passed, skipped);
		}
		return value ? this.classify(proc, toks, index, index, spanStart, value) : undefined;
	}

	/**
	 * The first use of `lower` on the path every call runs, or undefined when
	 * something before it may leave or the first use is on only some paths.
	 */
	private firstUse(body: readonly BodyNode[], lower: string): { toks: readonly VbaToken[]; index: number; spanStart: number } | undefined {
		for (const node of body) {
			if (this.activity?.isInactive(node.span) || node.kind === 'ConditionalDirective') {
				continue;
			}
			if (node.kind === 'VariableGroup') {
				if (this.mentionIn(this.tokens(node.span), lower) >= 0) {
					return undefined;
				}
				continue;
			}
			if (!isLeafStatement(node)) {
				const header = blockHeaderLineSpan(this.source, node.span);
				const headToks = this.tokens(header);
				const at = this.mentionIn(headToks, lower);
				if (at >= 0) {
					return { toks: headToks, index: at, spanStart: header.start };
				}
				if (!('body' in node) || !Array.isArray(node.body) || this.blockStops(node.body as BodyNode[], lower)
					|| this.mentionIn(this.tokens(blockFooterLineSpan(this.source, node.span)), lower) >= 0) {
					return undefined;
				}
				continue;
			}
			const toks = this.tokens(node.span);
			if (LEAVING_HEADS.has(tokenText(toks[0])) || isErrRaise(toks)) {
				return undefined;
			}
			if (node.kind === 'Statement' && node.singleLineIfBranches) {
				// The condition runs every time; the arms do not.
				const then = toks.findIndex((tok) => tokenText(tok) === 'then');
				const at = this.mentionIn(then < 0 ? toks : toks.slice(0, then), lower);
				if (at >= 0) {
					return { toks, index: at, spanStart: node.span.start };
				}
				if (this.leafStops(node, lower)) {
					return undefined;
				}
				continue;
			}
			const at = this.mentionIn(toks, lower);
			if (at >= 0) {
				return { toks, index: at, spanStart: node.span.start };
			}
		}
		return undefined;
	}

	/** Whether a block's body uses the name or may leave. */
	private blockStops(body: readonly BodyNode[], lower: string): boolean {
		for (const node of body) {
			if (this.activity?.isInactive(node.span)) {
				continue;
			}
			if (isLeafStatement(node)) {
				if (this.leafStops(node, lower)) {
					return true;
				}
				continue;
			}
			if (this.mentionIn(this.tokens(blockHeaderLineSpan(this.source, node.span)), lower) >= 0
				|| this.mentionIn(this.tokens(blockFooterLineSpan(this.source, node.span)), lower) >= 0) {
				return true;
			}
			if ('body' in node && Array.isArray(node.body) && this.blockStops(node.body as BodyNode[], lower)) {
				return true;
			}
		}
		return false;
	}

	private leafStops(node: LeafStatementNode, lower: string): boolean {
		for (const span of statementAndBranchSpans(node)) {
			const toks = this.tokens(span);
			if (LEAVING_HEADS.has(tokenText(toks[0])) || isErrRaise(toks) || this.mentionIn(toks, lower) >= 0) {
				return true;
			}
		}
		return false;
	}

	private tokens(span: Span): VbaToken[] {
		return statementTokensAfterLeadingLabel(this.source, span).filter((tok) => tok.kind !== 'comment');
	}

	/** Where the statement names the parameter itself, not a member or a named argument. */
	private mentionIn(toks: readonly VbaToken[], lower: string): number {
		return toks.findIndex((tok, i) => tokenName(tok)?.toLowerCase() === lower
			&& toks[i - 1]?.rawText !== '.' && toks[i - 1]?.rawText !== '!' && toks[i + 1]?.rawText !== ':=');
	}

	private paramArrayRead(
		proc: ProcedureNode,
		toks: readonly VbaToken[],
		at: number,
		spanStart: number,
		passed: number,
		skipped: ReadonlySet<number>,
	): RaisingUse | ParamArrayRead | undefined {
		if (toks[at + 1]?.rawText !== '(') {
			return undefined;
		}
		const close = matchParenFrom(toks, at + 1);
		const inner = toks.slice(at + 2, close);
		const negative = inner.length === 2 && inner[0].rawText === '-';
		const literal = inner[negative ? 1 : 0];
		if (close < 0 || inner.length !== (negative ? 2 : 1) || literal.kind !== 'integerLiteral' || toks[close + 1]?.rawText === '(') {
			return undefined;
		}
		const index = (negative ? -1 : 1) * Number(literal.rawText.replace(/[%&^]$/, ''));
		const span = { start: spanStart + toks[at].start, end: spanStart + toks[close].end };
		if (index < 0 || index >= passed) {
			return { index, span };
		}
		return skipped.has(index) ? this.classify(proc, toks, at, close, spanStart, { kind: 'missing' }) : undefined;
	}

	/** Whether the use of the value at tokens first..last raises, and how. */
	private classify(proc: ProcedureNode, toks: readonly VbaToken[], first: number, last: number, spanStart: number, value: OmittedValue): RaisingUse | undefined {
		const prev = toks[first - 1];
		const next = toks[last + 1];
		const nextText = tokenText(next);
		if (nextText === '(' || nextText === '.' || nextText === '!') {
			return undefined;
		}
		const span = { start: spanStart + toks[first].start, end: spanStart + toks[last].end };
		const assignAt = assignmentIndex(toks);
		if (last + 1 === assignAt) {
			return undefined; // the statement assigns it
		}
		const prevOperator = first - 1 === assignAt ? undefined : operatorWord(prev, toks[first - 2]);
		const nextOperator = operatorWord(next, toks[last]);
		const operation = quote(this.operationText(toks, assignAt, spanStart));
		if (value.kind === 'missing') {
			if (prevOperator && (MISSING_RAISING_OPERATORS.has(prevOperator) || prevOperator === 'not' || prevOperator === 'unary')) {
				return missingUse(span, `uses it in ${operation}`);
			}
			if (nextOperator && MISSING_RAISING_OPERATORS.has(nextOperator)) {
				return missingUse(span, `uses it in ${operation}`);
			}
			if (first === assignAt + 1 && last === toks.length - 1 && assignAt === (tokenText(toks[0]) === 'let' ? 2 : 1)) {
				const targetTok = toks[assignAt - 1];
				const target = tokenName(targetTok)!.toLowerCase();
				const type = target === proc.name.toLowerCase()
					? (proc.procKind === 'Function' ? functionType(proc) : undefined)
					: typeEnvironmentFor(this.symbols, proc).get(target);
				const normalized = normalizeType(type);
				if (normalized && isKnownScalarType(normalized)) {
					return missingUse(span, `assigns it to '${targetTok.rawText}', a ${capitalize(normalized)}`);
				}
				return undefined;
			}
			if (this.isWholeCondition(toks, first, last)) {
				return missingUse(span, `tests it in ${quote(this.headText(toks))}`);
			}
			const argument = argumentOf(toks, first, last);
			if (argument) {
				const calleeLower = argument.name;
				const procedure = this.procedures.get(calleeLower);
				if (procedure) {
					const target = procedure.params[argument.position];
					const type = target ? parameterType(target) : 'variant';
					if (target && !target.paramArray && target.byVal && type !== 'variant' && isKnownScalarType(type)) {
						return missingUse(span, `passes it ByVal to the ${capitalize(type)} '${target.name}' of '${procedure.name}'`);
					}
					return undefined;
				}
				if (isBareOrVbaQualifiedIntrinsicCall(toks, argument.callee)
					&& ((argument.position === 0 && MISSING_RAISING_INTRINSICS.has(calleeLower)) || (argument.position === 1 && MISSING_RAISING_SECOND.has(calleeLower)))) {
					return missingUse(span, `passes it to ${argument.display}`);
				}
			}
			return undefined;
		}
		if (value.kind === 'number') {
			if (value.value === 0 && prevOperator && (prevOperator === '/' || prevOperator === '\\' || prevOperator === 'mod') && nextText !== '^') {
				return { rule: 'divisionByZero', does: `divides by it in ${operation}`, error: "'11': Division by zero", span };
			}
			return undefined;
		}
		if (!isInvalidNumericString(value.value)) {
			return undefined;
		}
		const coerces = (word: string | undefined): boolean => word !== undefined && ['-', '*', '/', '\\', '^', 'mod', 'unary', 'not'].includes(word);
		const numberLiteral = (tok: VbaToken | undefined): boolean => tok?.kind === 'integerLiteral' || tok?.kind === 'floatLiteral';
		if (coerces(prevOperator) || coerces(nextOperator)
			|| (prevOperator === '+' && numberLiteral(toks[first - 2])) || (nextOperator === '+' && numberLiteral(toks[last + 2]))) {
			return { rule: 'stringArithmeticCoercion', does: `uses it as a number in ${operation}`, error: "'13': Type mismatch", span };
		}
		const argument = argumentOf(toks, first, last);
		if (argument && argument.position === 0 && argument.count === 1 && NUMBER_CONVERSIONS.has(argument.name)
			&& isBareOrVbaQualifiedIntrinsicCall(toks, argument.callee)) {
			return { rule: 'runtimeConversionValue', does: `converts it with ${argument.display}`, error: "'13': Type mismatch", span };
		}
		return undefined;
	}

	/** `If x Then`, `ElseIf x Then`, `Do While x`, `Loop Until x`, `While x`, `Select Case x`. */
	private isWholeCondition(toks: readonly VbaToken[], first: number, last: number): boolean {
		const before = toks.slice(0, first).map((tok) => tokenText(tok)).join(' ');
		const after = tokenText(toks[last + 1]);
		if ((before === 'if' || before === 'elseif') && after === 'then') {
			return true;
		}
		return last === toks.length - 1 && ['do while', 'do until', 'loop while', 'loop until', 'while', 'select case'].includes(before);
	}

	private headText(toks: readonly VbaToken[]): string {
		return toks.map((tok) => tok.rawText).join(' ');
	}

	/** The statement's value: what follows an assignment's `=`, or the whole statement. */
	private operationText(toks: readonly VbaToken[], assignAt: number, spanStart: number): string {
		const then = toks.findIndex((tok) => tokenText(tok) === 'then');
		const end = then > 0 ? then - 1 : toks.length - 1;
		const from = assignAt >= 0 && assignAt < end ? assignAt + 1 : 0;
		return this.source.slice(spanStart + toks[from].start, spanStart + toks[end].end);
	}
}

function missingUse(span: Span, does: string): RaisingUse {
	return { rule: 'variantValueMisuse', does, error: "'13': Type mismatch", span };
}

function quote(text: string): string {
	return `'${text.length > 60 ? `${text.slice(0, 57)}...` : text}'`;
}

function capitalize(type: string): string {
	const names: Record<string, string> = { longlong: 'LongLong', longptr: 'LongPtr' };
	return names[type] ?? type.charAt(0).toUpperCase() + type.slice(1);
}

function functionType(proc: ProcedureNode): string | undefined {
	if (proc.typeSuffix) {
		return SUFFIX_TYPES[proc.typeSuffix];
	}
	return proc.returnType ?? 'variant';
}

/**
 * The operator a neighbouring token is, lowercased: a binary operator, `not`,
 * or 'unary' for a sign with no operand before it.
 */
function operatorWord(tok: VbaToken | undefined, before: VbaToken | undefined): string | undefined {
	if (!tok) {
		return undefined;
	}
	const word = tokenText(tok);
	if (word === 'not') {
		return 'not';
	}
	if (word === 'mod' || word === 'and' || word === 'or' || word === 'like') {
		return tok.kind === 'keyword' || tok.kind === 'operator' ? word : undefined;
	}
	if (tok.kind !== 'operator') {
		return undefined;
	}
	if ((word === '-' || word === '+') && before !== undefined && !endsOperand(before)) {
		return 'unary';
	}
	if ((word === '-' || word === '+') && before === undefined) {
		return 'unary';
	}
	return word;
}

function endsOperand(tok: VbaToken): boolean {
	return tok.kind === 'identifier' || tok.kind === 'integerLiteral' || tok.kind === 'floatLiteral' || tok.kind === 'stringLiteral'
		|| tok.kind === 'dateLiteral' || tok.rawText === ')' || (tok.kind === 'keyword' && ['true', 'false', 'nothing', 'empty', 'null', 'me'].includes(tokenText(tok)));
}

/** The call whose argument the tokens first..last are, standing alone in their slot. */
function argumentOf(toks: readonly VbaToken[], first: number, last: number): { callee: number; name: string; display: string; position: number; count: number } | undefined {
	const prev = toks[first - 1]?.rawText;
	const next = toks[last + 1]?.rawText;
	if ((prev !== '(' && prev !== ',') || (next !== ')' && next !== ',')) {
		return undefined;
	}
	let depth = 0;
	let position = 0;
	for (let i = first - 1; i >= 0; i--) {
		const raw = toks[i].rawText;
		if (raw === ')') {
			depth++;
		} else if (raw === '(') {
			if (depth === 0) {
				// `Left$(` lexes as `Left` and `$`.
				const callee = toks[i - 1]?.rawText === '$' ? i - 2 : i - 1;
				if (tokenName(toks[callee]) === undefined) {
					return undefined;
				}
				const suffix = callee === i - 2 ? '$' : '';
				const close = matchParenFrom(toks, i);
				let count = 1;
				let inner = 0;
				for (let k = i + 1; k < close; k++) {
					if (toks[k].rawText === '(') {
						inner++;
					} else if (toks[k].rawText === ')') {
						inner--;
					} else if (toks[k].rawText === ',' && inner === 0) {
						count++;
					}
				}
				return { callee, name: tokenText(toks[callee]) + suffix, display: toks[callee].rawText + suffix, position, count };
			}
			depth--;
		} else if (raw === ',' && depth === 0) {
			position++;
		}
	}
	return undefined;
}

/** Heads whose `=` compares; any other statement's first top-level `=` assigns. */
const COMPARING_HEADS: ReadonlySet<string> = new Set(['if', 'elseif', 'do', 'loop', 'while', 'select', 'case']);

function assignmentIndex(toks: readonly VbaToken[]): number {
	if (COMPARING_HEADS.has(tokenText(toks[0]))) {
		return -1;
	}
	let depth = 0;
	for (let i = 0; i < toks.length; i++) {
		const raw = toks[i].rawText;
		if (raw === '(') {
			depth++;
		} else if (raw === ')') {
			depth--;
		} else if (raw === '=' && depth === 0 && toks[i].kind === 'operator') {
			return i;
		}
	}
	return -1;
}

function isErrRaise(toks: readonly VbaToken[]): boolean {
	return tokenText(toks[0]) === 'err' && toks[1]?.rawText === '.' && tokenText(toks[2]) === 'raise';
}
