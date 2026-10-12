import { tokenize } from '../lexer/tokenize';
import { MAX_EXPRESSION_DEPTH } from '../parser/expressionLimits';
import { isKnownDirectiveFreeModule } from '../parser/moduleParseFacts';
import type { VbaToken } from '../lexer/tokenKinds';
import { relationalOperatorAt, tokenWord } from '../lexer/tokenHelpers';
import { bankersRound, parseVbaIntegerLiteral } from '../constants/integerConstantExpression';
import { dateLiteralSerial } from '../constants/dateLiteral';
import {
	isInvalidBooleanString,
	isInvalidDateString,
	isInvalidNumericString,
	numericStringReadings,
	numericStringVerdict,
} from '../diagnostics/stringConversion';
import type {
	BodyNode,
	ConditionalDirectiveNode,
	ModuleNode,
	ProcedureNode,
	Span,
} from '../parser/nodes';

/**
 * The values a directive expression has beside the plain ones: Empty, Null,
 * Nothing and a Date, which is its serial (issue #208).
 */
export type ConditionalSpecialValue =
	| { readonly kind: 'empty' }
	| { readonly kind: 'null' }
	| { readonly kind: 'nothing' }
	| { readonly kind: 'date'; readonly serial: number };
export type ConditionalValue = boolean | number | string | ConditionalSpecialValue;
export type ConditionalActivity = 'active' | 'inactive' | 'unknown';

export interface ConditionalCompilationEnvironment {
	compilerConstants?: Readonly<Record<string, ConditionalValue>>;
	projectConstants?: Readonly<Record<string, ConditionalValue>>;
}

export interface ConditionalDirectiveOccurrence {
	directive: ConditionalDirectiveNode;
	container:
		| { kind: 'module' }
		| { kind: 'procedure'; name: string; span: Span };
}

export interface ConditionalConstDefinition {
	name: string;
	nameSpan: Span;
	valueRaw?: string;
	value: ConditionalValue | undefined;
	directive: ConditionalDirectiveNode;
}

export interface ConditionalCompilationIndex {
	directives: ConditionalDirectiveOccurrence[];
	constants: ConditionalConstDefinition[];
}

export interface ConditionalActivityTracker {
	activityForSpan(span: Span): ConditionalActivity;
	isInactive(span: Span): boolean;
	/**
	 * Whether the two spans sit in different arms of one `#If` chain, and so
	 * are never compiled together however the constants evaluate.
	 */
	mutuallyExclusive(a: Span, b: Span): boolean;
	/**
	 * Whether the two spans sit under exactly the same arms, so every build
	 * either compiles both or neither. Stricter than "not mutually exclusive":
	 * spans in two SEPARATE chains are neither exclusive nor the same branch,
	 * because a build may take one and not the other. Rules that pair two
	 * pieces of one construct need this, since a pairing made across different
	 * chains is a guess about a build that may never exist.
	 */
	inSameBranch(a: Span, b: Span): boolean;
}

/**
 * The compiler constants of 64-bit Office on Windows. One that is on is 1,
 * not True (-1): `#If Not Win64` is true there, since Not 1 is -2, and
 * `#If VBA7 = True` is false (issue #214, measured in 64-bit Excel 16.0).
 * One that is off is 0.
 */
const DEFAULT_COMPILER_CONSTANTS: Readonly<Record<string, ConditionalValue>> = {
	VBA7: 1,
	VBA6: 1,
	Win64: 1,
	// Win32 is on in 64-bit Office as well: it means Windows, not a width
	// (issue #192, measured in 64-bit Excel 16.0). Win16 is off.
	Win32: 1,
	Win16: 0,
	Mac: 0,
	// `TWINBASIC` is a compiler auto-constant defined only by the twinBASIC
	// compiler; in Excel VBA it is undefined, and an undefined name is 0
	// (VBE-oracle verified). Modern VBA libraries gate twinBASIC-only
	// intrinsics behind `#If TWINBASIC Then ...`, so without this default those
	// (inactive) branches were analyzed and produced false positives. Unlike a
	// genuine unprovable host flag, TWINBASIC's value in VBA is known, so it
	// is a default, not left `unknown`.
	TWINBASIC: 0,
};

function effectiveConditionalCompilationEnvironment(
	env: ConditionalCompilationEnvironment = {},
): ConditionalCompilationEnvironment {
	return {
		compilerConstants: {
			...DEFAULT_COMPILER_CONSTANTS,
			...(env.compilerConstants ?? {}),
		},
		projectConstants: env.projectConstants,
	};
}

export function createConditionalActivityTracker(
	module: ModuleNode,
	env: ConditionalCompilationEnvironment = {},
): ConditionalActivityTracker | undefined {
	if (!moduleHasConditionalDirectives(module)) {
		return undefined;
	}
	const effectiveEnv = effectiveConditionalCompilationEnvironment(env);
	// One forward sweep at construction: replay the directive stack once and
	// record the activity in effect after each directive, so per-span queries
	// become a binary search instead of a full replay from offset 0.
	const events = collectConditionalActivityEvents(module, effectiveEnv);
	// Mirror conditionalActivityAtOffset: directives starting at or after the
	// queried offset are not applied.
	const eventForSpan = (span: Span): ConditionalActivityEvent | undefined => {
		let lo = -1;
		let hi = events.length - 1;
		while (lo < hi) {
			const mid = (lo + hi + 1) >> 1;
			if (events[mid].start < span.start) {
				lo = mid;
			} else {
				hi = mid - 1;
			}
		}
		return lo >= 0 ? events[lo] : undefined;
	};
	const activityForSpan = (span: Span): ConditionalActivity =>
		eventForSpan(span)?.activity ?? 'active';
	return {
		activityForSpan,
		isInactive: (span: Span): boolean => activityForSpan(span) === 'inactive',
		mutuallyExclusive: (a: Span, b: Span): boolean =>
			armsDiverge(eventForSpan(a)?.branch, eventForSpan(b)?.branch),
		// Arms are immutable and structurally shared, so one arm is one object:
		// identity IS the comparison, including after a nested chain has opened
		// and closed again between the two spans.
		inSameBranch: (a: Span, b: Span): boolean =>
			eventForSpan(a)?.branch === eventForSpan(b)?.branch,
	};
}

/**
 * One arm of one `#If` chain, as a persistent stack: `parent` is the enclosing
 * chain's arm. Immutable, so an event can keep the arm in effect when it was
 * recorded without copying the stack.
 */
interface ConditionalArm {
	/** Identifies the `#If` chain; every arm of one chain shares it. */
	chain: number;
	/** 0 for the `#If`, then one per `#ElseIf` / `#Else`. */
	index: number;
	parent: ConditionalArm | undefined;
}

/**
 * Whether the two arm stacks disagree about which arm of a shared chain they
 * are in - the branches then exclude each other, whatever the constants are
 * worth. Parent paths descend by source-order chain ID, allowing a linear walk.
 */
function armsDiverge(a: ConditionalArm | undefined, b: ConditionalArm | undefined): boolean {
	// Chain IDs increase in source order, so each parent path descends by ID.
	// Merge the paths to find their innermost shared chain without a cross scan.
	while (a && b) {
		if (a === b) {
			return false;
		}
		if (a.chain === b.chain) {
			return a.index !== b.index;
		}
		if (a.chain > b.chain) {
			a = a.parent;
		} else {
			b = b.parent;
		}
	}
	return false;
}

interface ConditionalActivityEvent {
	start: number;
	activity: ConditionalActivity;
	branch: ConditionalArm | undefined;
}

/** The environment's `#Const` values keyed by lower-cased name, the way lookups spell them. */
function projectConstantsOf(env: ConditionalCompilationEnvironment): Map<string, ConditionalValue> {
	const out = new Map<string, ConditionalValue>();
	for (const [name, value] of Object.entries(env.projectConstants ?? {})) {
		out.set(name.toLowerCase(), value);
	}
	return out;
}

function collectConditionalActivityEvents(
	module: ModuleNode,
	effectiveEnv: ConditionalCompilationEnvironment,
): ConditionalActivityEvent[] {
	const directives = collectConditionalDirectives(module);
	const projectConstants = projectConstantsOf(effectiveEnv);
	const evaluate = createConditionalExpressionEvaluator(effectiveEnv, projectConstants);
	const stack: ConditionalFrame[] = [];
	let current: ConditionalActivity = 'active';
	let branch: ConditionalArm | undefined;
	let chains = 0;
	const events: ConditionalActivityEvent[] = [];
	for (const { directive } of directives) {
		current = applyConditionalDirective(directive, evaluate, projectConstants, stack, current);
		switch (directive.directiveKind) {
			case 'If':
				branch = { chain: chains++, index: 0, parent: branch };
				break;
			case 'ElseIf':
			case 'Else':
				// An `#ElseIf` with no open `#If` is a parse-level error; leave
				// the stack alone rather than inventing an arm for it.
				if (branch) {
					branch = { chain: branch.chain, index: branch.index + 1, parent: branch.parent };
				}
				break;
			case 'EndIf':
				branch = branch?.parent;
				break;
			default:
				break;
		}
		events.push({ start: directive.span.start, activity: current, branch });
	}
	return events;
}

export function moduleHasConditionalDirectives(module: ModuleNode): boolean {
	if (isKnownDirectiveFreeModule(module)) { return false; }
	for (const member of module.members) {
		if (member.kind === 'ConditionalDirective') {
			return true;
		}
		if (member.kind === 'Procedure' && bodyHasConditionalDirectives(member.body)) {
			return true;
		}
		if (
			(member.kind === 'Enum' || member.kind === 'Type') &&
			(member.directives?.length ?? 0) > 0
		) {
			return true;
		}
	}
	return false;
}

export function indexConditionalCompilation(
	module: ModuleNode,
	env: ConditionalCompilationEnvironment = {},
): ConditionalCompilationIndex {
	const directives = collectConditionalDirectives(module);
	const constants = collectConditionalConstants(
		directives,
		effectiveConditionalCompilationEnvironment(env),
	);
	return { directives, constants };
}

export function collectConditionalDirectives(
	module: ModuleNode,
): ConditionalDirectiveOccurrence[] {
	const out: ConditionalDirectiveOccurrence[] = [];
	for (const member of module.members) {
		if (member.kind === 'ConditionalDirective') {
			out.push({ directive: member, container: { kind: 'module' } });
		} else if (member.kind === 'Procedure') {
			collectBodyDirectives(member.body, member, out);
		} else if (member.kind === 'Enum' || member.kind === 'Type') {
			for (const directive of member.directives ?? []) {
				out.push({ directive, container: { kind: 'module' } });
			}
		}
	}
	return out.sort((a, b) => a.directive.span.start - b.directive.span.start);
}

/**
 * Parses the VBE "Conditional Compilation Arguments" project property, which
 * MS-OVBA stores as `Name = Value : Name2 = Value2`.
 *
 * The VBE accepts integers here, and writes booleans as VBA does: -1 for True,
 * 0 for False. A value that is not an integer is kept as its raw text, so a
 * comparison against a string constant still works and an unreadable entry
 * cannot silently become a number. An entry with no `=` names nothing and is
 * skipped rather than guessed at. A name is a VBA identifier, whose letters
 * are any the project's code page holds, so a name opening with an E acute
 * is one (issue #207).
 */
export function parseProjectConditionalConstants(
	raw: string | undefined,
): Record<string, ConditionalValue> {
	const constants: Record<string, ConditionalValue> = {};
	for (const entry of (raw ?? '').split(':')) {
		const eq = entry.indexOf('=');
		if (eq < 0) {
			continue;
		}
		const name = entry.slice(0, eq).trim();
		const valueText = entry.slice(eq + 1).trim();
		if (!/^\p{L}[\p{L}\p{N}_]*$/u.test(name)) {
			continue;
		}
		constants[name] = /^[+-]?\d+$/.test(valueText) ? Number(valueText) : valueText;
	}
	return constants;
}

/**
 * The constants a directive sees, the defaults for 64-bit Office included:
 * what `conditionalCompilerConstants` gives for the environment the branch
 * activity is decided in (issue #215).
 */
export function compilerConstantsWithDefaults(
	env: ConditionalCompilationEnvironment = {},
): Map<string, ConditionalValue> {
	return conditionalCompilerConstants(effectiveConditionalCompilationEnvironment(env));
}

export function conditionalCompilerConstants(
	env: ConditionalCompilationEnvironment = {},
): Map<string, ConditionalValue> {
	const constants = new Map<string, ConditionalValue>();
	for (const [name, value] of Object.entries(env.compilerConstants ?? {})) {
		constants.set(name.toLowerCase(), value);
	}
	for (const [name, value] of Object.entries(env.projectConstants ?? {})) {
		constants.set(name.toLowerCase(), value);
	}
	return constants;
}

/**
 * A directive's expression as tokens. It is lexed after a throwaway `x=`, so a
 * `#` that opens it is a date literal and not the directive marker a `#` at
 * the start of a statement is: `#If #1/2/2000# > #1/1/2000# Then` (issue #208).
 */
function directiveExpressionTokens(expression: string): VbaToken[] {
	return tokenize(`x=${expression}`)
		.slice(2)
		.filter((t) => t.kind !== 'comment' && t.kind !== 'newline');
}

export function evaluateConditionalExpression(
	expression: string | undefined,
	env: ConditionalCompilationEnvironment = {},
): ConditionalValue | undefined {
	if (!expression?.trim()) {
		return undefined;
	}
	const parser = new ConditionalExpressionParser(
		directiveExpressionTokens(expression),
		conditionalCompilerConstants(env),
		env.projectConstants !== undefined,
	);
	return parser.parse();
}

export function conditionalActivityAtOffset(
	module: ModuleNode,
	offset: number,
	env: ConditionalCompilationEnvironment = {},
): ConditionalActivity {
	const effectiveEnv = effectiveConditionalCompilationEnvironment(env);
	const directives = collectConditionalDirectives(module);
	const projectConstants = projectConstantsOf(effectiveEnv);
	const evaluate = createConditionalExpressionEvaluator(effectiveEnv, projectConstants);
	const stack: ConditionalFrame[] = [];
	let current: ConditionalActivity = 'active';

	for (const { directive } of directives) {
		if (directive.span.start >= offset) {
			break;
		}
		current = applyConditionalDirective(directive, evaluate, projectConstants, stack, current);
	}
	return current;
}

/** An `#If`, `#ElseIf` or `#Const` expression the VBE refuses to compile, and why. */
export type InvalidConditionDirective =
	| { reason: 'null'; directive: ConditionalDirectiveNode }
	| { reason: 'type-mismatch'; directive: ConditionalDirectiveNode; value: string; as: StringConversion };

/**
 * The `#If` and `#ElseIf` lines whose condition the VBE refuses, in one replay:
 *
 * - Null: "Invalid use of Null" (issue #208, Excel 16.0). `#If Null`,
 *   `#If Null = 1`, `#If Not Null` and a `#Const N = Null` read by `#If N`
 *   all are. Only a line the VBE is sure to evaluate is listed: one in code
 *   that is compiled, and for `#ElseIf`, after arms that were all False.
 * - A string VBA cannot convert: "Type mismatch" (Excel 16.0). A condition
 *   that is a string no locale reads as True or False (`#If "abc"`, `#If ""`,
 *   `#If " True "`, `#If S` after `#Const S = "x"`), or that hands one to an
 *   operator needing a number or a Boolean (`#If Not "abc"`,
 *   `#If UNDEFINED_X - "abc"`, `#If "abc" = 1`). A `#Const` value that hands
 *   one to an operator is refused too (`#Const A = "abc" + 1`), even when
 *   nothing reads it; one that is a string is not (`#Const S = "x"`). Unlike
 *   Null, the VBE refuses these in an arm it skips as well, and an `#ElseIf`
 *   after a taken arm, so a line the VBE does not evaluate is read too, when it
 *   can hold a string at all: a quote, an `&`, or a constant that is one.
 */
export function invalidConditionDirectives(
	module: ModuleNode,
	env: ConditionalCompilationEnvironment = {},
): InvalidConditionDirective[] {
	if (!moduleHasConditionalDirectives(module)) {
		return [];
	}
	const effectiveEnv = effectiveConditionalCompilationEnvironment(env);
	const projectConstants = projectConstantsOf(effectiveEnv);
	// Strings the host supplied, until a module #Const gives the name a value.
	const hostStrings = new Set<string>();
	for (const table of [effectiveEnv.compilerConstants, effectiveEnv.projectConstants]) {
		for (const [name, value] of Object.entries(table ?? {})) {
			if (typeof value === 'string') {
				hostStrings.add(name.toLowerCase());
			}
		}
	}
	const read = createConditionalExpressionReader(effectiveEnv, projectConstants, undefined, hostStrings);
	const evaluate: ConditionalExpressionEvaluator = (expression) => read(expression).value;
	const stack: ConditionalFrame[] = [];
	let current: ConditionalActivity = 'active';
	const out: InvalidConditionDirective[] = [];
	// Whether any constant holds a string, so a condition with no quote and
	// no `&` can still be one.
	let stringConstant = [effectiveEnv.compilerConstants, effectiveEnv.projectConstants]
		.some((table) => Object.values(table ?? {}).some((value) => typeof value === 'string'));
	const mayHoldString = (raw: string): boolean => stringConstant || raw.includes('"') || raw.includes('&');
	for (const { directive } of collectConditionalDirectives(module)) {
		if (directive.directiveKind === 'Const' && directive.name && mayHoldString(directive.valueRaw ?? '')) {
			const mismatch = read(directive.valueRaw).operatorMismatch;
			if (mismatch !== undefined) {
				out.push({ reason: 'type-mismatch', directive, ...mismatch });
			}
		}
		if (directive.directiveKind === 'If' || directive.directiveKind === 'ElseIf') {
			const frame = stack[stack.length - 1];
			const evaluated = directive.directiveKind === 'If'
				? current === 'active'
				: frame?.parent === 'active' && !frame.seenTrue && !frame.seenUnknown;
			const raw = directive.conditionRaw ?? '';
			if (!evaluated && !mayHoldString(raw)) {
				current = applyConditionalDirective(directive, evaluate, projectConstants, stack, current);
				continue;
			}
			const reading = read(directive.conditionRaw);
			if (reading.mismatch !== undefined) {
				out.push({ reason: 'type-mismatch', directive, ...reading.mismatch });
			} else if (evaluated && reading.value !== undefined && isNull(reading.value)) {
				out.push({ reason: 'null', directive });
			}
		}
		// A #Const that gives its name a value replaces a host string; one that
		// fails (`#Const FLAG = 1 / 0`) leaves it.
		const replacesHost = directive.directiveKind === 'Const' && directive.name !== undefined
			&& hostStrings.has(directive.name.toLowerCase()) && evaluate(directive.valueRaw) !== undefined;
		current = applyConditionalDirective(directive, evaluate, projectConstants, stack, current);
		if (directive.directiveKind === 'Const' && directive.name) {
			const lower = directive.name.toLowerCase();
			if (replacesHost) {
				hostStrings.delete(lower);
			}
			if (typeof projectConstants.get(lower) === 'string') {
				stringConstant = true;
			}
		}
	}
	return out;
}

/** The lines `invalidConditionDirectives` refuses for Null. */
export function nullConditionDirectives(
	module: ModuleNode,
	env: ConditionalCompilationEnvironment = {},
): ConditionalDirectiveNode[] {
	return invalidConditionDirectives(module, env)
		.flatMap((item) => item.reason === 'null' ? [item.directive] : []);
}

function applyConditionalDirective(
	directive: ConditionalDirectiveNode,
	evaluate: ConditionalExpressionEvaluator,
	projectConstants: Map<string, ConditionalValue>,
	stack: ConditionalFrame[],
	current: ConditionalActivity,
): ConditionalActivity {
	switch (directive.directiveKind) {
		case 'Const': {
			// A #Const defines its constant even inside a #If False: the VBE
			// reads every #Const line (issue #192, measured in Excel 16.0).
			if (directive.name) {
				const value = evaluate(directive.valueRaw);
				if (value !== undefined) {
					projectConstants.set(directive.name.toLowerCase(), value);
				}
			}
			return current;
		}
		case 'If': {
			const condition = current === 'inactive'
				? 'inactive' : conditionActivity(directive, evaluate);
			const frame: ConditionalFrame = {
				parent: current,
				current: combineActivity(current, condition),
				seenTrue: condition === 'active',
				seenUnknown: condition === 'unknown',
			};
			stack.push(frame);
			return frame.current;
		}
		case 'ElseIf': {
			const frame = stack[stack.length - 1];
			if (!frame) {
				return current;
			}
			if (frame.seenTrue || frame.parent === 'inactive') {
				// Constants still replay, but this arm cannot change activity.
				frame.current = 'inactive';
				return frame.current;
			}
			const condition = conditionActivity(directive, evaluate);
			if (frame.seenUnknown && condition !== 'inactive') {
				frame.current = combineActivity(frame.parent, 'unknown');
			} else {
				frame.current = combineActivity(frame.parent, condition);
			}
			frame.seenTrue ||= condition === 'active';
			frame.seenUnknown ||= condition === 'unknown';
			return frame.current;
		}
		case 'Else': {
			const frame = stack[stack.length - 1];
			if (!frame) {
				return current;
			}
			if (frame.seenTrue) {
				frame.current = 'inactive';
			} else if (frame.seenUnknown) {
				frame.current = combineActivity(frame.parent, 'unknown');
			} else {
				frame.current = frame.parent;
			}
			frame.seenTrue = true;
			return frame.current;
		}
		case 'EndIf': {
			const frame = stack.pop();
			return frame?.parent ?? current;
		}
		case 'Unknown':
			return current;
	}
}

function collectBodyDirectives(
	body: BodyNode[],
	procedure: ProcedureNode,
	out: ConditionalDirectiveOccurrence[],
): void {
	for (const node of body) {
		if (node.kind === 'ConditionalDirective') {
			out.push({
				directive: node,
				container: {
					kind: 'procedure',
					name: procedure.name,
					span: procedure.span,
				},
			});
		} else if ('body' in node && Array.isArray(node.body)) {
			collectBodyDirectives(node.body, procedure, out);
		}
	}
}

function bodyHasConditionalDirectives(body: BodyNode[]): boolean {
	for (const node of body) {
		if (node.kind === 'ConditionalDirective') {
			return true;
		}
		if ('body' in node && Array.isArray(node.body) && bodyHasConditionalDirectives(node.body)) {
			return true;
		}
	}
	return false;
}

function collectConditionalConstants(
	directives: readonly ConditionalDirectiveOccurrence[],
	env: ConditionalCompilationEnvironment,
): ConditionalConstDefinition[] {
	const projectConstants = projectConstantsOf(env);
	const evaluate = createConditionalExpressionEvaluator(env, projectConstants, true);
	const constants: ConditionalConstDefinition[] = [];
	for (const { directive } of directives) {
		if (directive.directiveKind !== 'Const' || !directive.name || !directive.nameSpan) {
			continue;
		}
		// The index historically supplies a project table even when the caller
		// did not, so an absent name is zero on this path.
		const value = evaluate(directive.valueRaw);
		if (value !== undefined) {
			projectConstants.set(directive.name.toLowerCase(), value);
		}
		constants.push({
			name: directive.name,
			nameSpan: directive.nameSpan,
			valueRaw: directive.valueRaw,
			value,
			directive,
		});
	}
	return constants;
}

interface ConditionalFrame {
	parent: ConditionalActivity;
	current: ConditionalActivity;
	seenTrue: boolean;
	seenUnknown: boolean;
}

function conditionActivity(
	directive: ConditionalDirectiveNode,
	evaluate: ConditionalExpressionEvaluator,
): ConditionalActivity {
	const value = evaluate(directive.conditionRaw);
	const holds = value === undefined ? undefined : truthy(value);
	if (holds === undefined) {
		return 'unknown';
	}
	return holds ? 'active' : 'inactive';
}

type ConditionalExpressionEvaluator = (expression: string | undefined) => ConditionalValue | undefined;

/**
 * An expression's value, and the string it could not convert, if any: one an
 * operator needed (`operatorMismatch`), or, read as a condition, the value
 * itself (`mismatch`, which includes the first).
 */
interface ConditionalReading {
	value: ConditionalValue | undefined;
	mismatch?: StringMismatch;
	operatorMismatch?: StringMismatch;
}

/** What a string was needed as: True or False, a number, or a date. */
export type StringConversion = 'boolean' | 'number' | 'date';

/** A string an expression could not convert, and what it was needed as. */
export interface StringMismatch {
	value: string;
	as: StringConversion;
}

function createConditionalExpressionEvaluator(
	env: ConditionalCompilationEnvironment,
	projectConstants: ReadonlyMap<string, ConditionalValue>,
	undefinedIsZero = env.projectConstants !== undefined,
): ConditionalExpressionEvaluator {
	const read = createConditionalExpressionReader(env, projectConstants, undefinedIsZero);
	return (expression) => read(expression).value;
}

function createConditionalExpressionReader(
	env: ConditionalCompilationEnvironment,
	projectConstants: ReadonlyMap<string, ConditionalValue>,
	undefinedIsZero = env.projectConstants !== undefined,
	hostStrings?: ReadonlySet<string>,
): (expression: string | undefined) => ConditionalReading {
	// Compiler values stay fixed during a replay; module #Const values evolve.
	// Keep the lookup query-local and lazy for replays with no expressions.
	let compilerConstants: ReadonlyMap<string, ConditionalValue> | undefined;
	const constants = {
		get: (name: string): ConditionalValue | undefined => projectConstants.has(name)
			? projectConstants.get(name) : compilerConstants?.get(name),
	};
	return (expression) => {
		if (!expression?.trim()) {
			return { value: undefined };
		}
		compilerConstants ??= conditionalCompilerConstants({ compilerConstants: env.compilerConstants });
		return new ConditionalExpressionParser(
			directiveExpressionTokens(expression),
			constants,
			undefinedIsZero,
			hostStrings,
		).read();
	};
}

function combineActivity(
	parent: ConditionalActivity,
	condition: ConditionalActivity,
): ConditionalActivity {
	if (parent === 'inactive' || condition === 'inactive') {
		return 'inactive';
	}
	if (parent === 'unknown' || condition === 'unknown') {
		return 'unknown';
	}
	return 'active';
}

const EMPTY: ConditionalSpecialValue = { kind: 'empty' };
const NULL: ConditionalSpecialValue = { kind: 'null' };
const NOTHING: ConditionalSpecialValue = { kind: 'nothing' };

function isSpecial(value: ConditionalValue, kind: ConditionalSpecialValue['kind']): boolean {
	return typeof value === 'object' && value.kind === kind;
}

function isNull(value: ConditionalValue): boolean {
	return isSpecial(value, 'null');
}

/**
 * Whether a condition holds. Undefined for Null, which the VBE refuses as a
 * condition ("Invalid use of Null"), for Nothing ("Invalid use of object"),
 * and for a string VBA cannot read as a Boolean ("Type mismatch").
 */
function truthy(value: ConditionalValue): boolean | undefined {
	if (typeof value === 'boolean') {
		return value;
	}
	if (typeof value === 'number') {
		return value !== 0;
	}
	if (typeof value === 'string') {
		return stringTruthy(value);
	}
	switch (value.kind) {
		case 'empty': return false;
		case 'date': return value.serial !== 0;
		default: return undefined;
	}
}

/**
 * A string as the VBE reads it for a condition: as VBA converts a string to a
 * Boolean (issue #188), measured in a directive in Excel 16.0 (#If "0",
 * "0.0", "False" and #Const S = "0" then #If S are False; "1", "2.5",
 * "True", "true", "&H10" and "1e3" are True; "", "abc" and " True " are a
 * Type mismatch). "True" and "False" in any case, with nothing around them,
 * are those values; a number is True unless it is zero. A plain number whose
 * value depends on the locale ("2.5", "1,000") is zero in one reading
 * exactly when it is in the other, so it is decided too. Undefined for a
 * string no locale converts, and for any other number without a value here
 * (a hex string past a Long, a currency symbol, an exponent beside a
 * separator), which may not convert at all.
 */
function stringTruthy(value: string): boolean | undefined {
	if (/^true$/i.test(value)) {
		return true;
	}
	if (/^false$/i.test(value)) {
		return false;
	}
	const trimmed = value.replace(/^[ \t]+|[ \t]+$/g, '');
	if (RADIX_STRING.test(trimmed)) {
		const radix = radixStringValue(trimmed);
		return radix === undefined ? undefined : radix !== 0;
	}
	const verdict = numericStringVerdict(value);
	if (verdict.kind === 'invalid') {
		return undefined;
	}
	if (verdict.value !== undefined) {
		return verdict.value !== 0;
	}
	const readings = numericStringReadings(value);
	return readings === undefined ? undefined : readings[0] !== 0;
}

/**
 * Evaluates a #If or #Const expression as the VBE does (issues #192 and #208,
 * measured in Excel 16.0). The operators and their order are VBA's own,
 * loosest first: Imp, Eqv, Xor, Or, And, Not, the comparisons with Like and
 * Is, &, + and -, Mod, \, * and /, unary minus, ^. Not, And, Or, Xor, Eqv and
 * Imp are bitwise on numbers, as in code: `Not 1` is -2, which is True, and
 * `1 And 2` is 0. Two Booleans give a Boolean. Strings compare without regard
 * to case, and so does Like: `"A" = "a"` and `"ABC" Like "a*"` are True. Hex
 * and octal literals keep their width: &HFFFF is -1.
 *
 * Empty is 0 beside a number and "" beside a string. Null propagates through
 * arithmetic and comparisons, is "" to &, and follows VBA's three-valued
 * logic: `Null Or True` is True and `Null And False` is False. A date literal
 * is its serial, so `#12:00:00 AM#` is False. `Nothing Is Nothing` is True.
 */
class ConditionalExpressionParser {
	private index = 0;
	private parenthesisDepth = 0;
	/**
	 * The first string an operator needed as a number or a Boolean and could
	 * not convert, a compile error ("Type mismatch", measured in Excel 16.0).
	 */
	mismatch: { value: string; as: StringConversion } | undefined;
	/** A string came from the host's constants, whose reading was not measured. */
	private hostString = false;
	/**
	 * A value was missing where the expression needs one, a `(` or a string
	 * was not closed, or a keyword stood for a value. Only the mismatch
	 * depends on it.
	 */
	private broken = false;

	constructor(
		private readonly tokens: readonly VbaToken[],
		private readonly constants: Pick<ReadonlyMap<string, ConditionalValue>, 'get'>,
		/**
		 * Whether a name no constant defines evaluates as the VBE evaluates
		 * it, to 0. Not to Empty: `UNDEFINED & "x" = "x"` is False where
		 * `Empty & "x" = "x"` is True (issue #208). True only when the caller
		 * supplied the project's own conditional constants, so an absent name
		 * is provably undefined rather than unknown (issue #102); a module's
		 * `#Const` lines are folded into the same table before any `#If`
		 * reads them.
		 */
		private readonly undefinedIsZero: boolean,
		/**
		 * Names whose value is a string the host supplied, not a module
		 * `#Const`: how the VBE reads such an argument was not measured.
		 */
		private readonly hostStrings?: ReadonlySet<string>,
	) {}

	parse(): ConditionalValue | undefined {
		const value = this.parseLogical(0);
		return this.index >= this.tokens.length ? value : undefined;
	}

	/**
	 * The value, and the string the expression could not convert, if any. An
	 * expression the parser cannot read, missing a value or a `)` or with
	 * tokens left over, reports no mismatch.
	 */
	read(): ConditionalReading {
		const value = this.parseLogical(0);
		if (this.index < this.tokens.length) {
			return { value: undefined };
		}
		if (this.broken || this.hostString) {
			// The value is as `parse` gives it; only the mismatch is withheld.
			return { value };
		}
		const operatorMismatch = this.mismatch;
		const mismatch = operatorMismatch
			?? (typeof value === 'string' && isInvalidBooleanString(value)
				? { value, as: 'boolean' as const } : undefined);
		return { value, ...(mismatch === undefined ? {} : { mismatch }),
			...(operatorMismatch === undefined ? {} : { operatorMismatch }) };
	}

	/** Imp, Eqv, Xor, Or, And, loosest first; each level is left-associative. */
	private parseLogical(level: number): ConditionalValue | undefined {
		if (level === LOGICAL_LEVELS.length) {
			return this.parseNot();
		}
		let left = this.parseLogical(level + 1);
		while (this.matchWord(LOGICAL_LEVELS[level])) {
			const right = this.parseLogical(level + 1);
			this.requireConvertible(left, 'boolean');
			this.requireConvertible(right, 'boolean');
			left = left === undefined || right === undefined ? undefined : logical(LOGICAL_LEVELS[level], left, right);
		}
		return left;
	}

	/** `Not` binds looser than a comparison: `Not 1 = 2` is `Not (1 = 2)`. */
	private parseNot(): ConditionalValue | undefined {
		let count = 0;
		while (this.matchWord('not')) {
			count++;
		}
		let value = this.parseComparison();
		if (count > 0) {
			this.requireConvertible(value, 'boolean');
		}
		while (count-- > 0) {
			if (typeof value === 'boolean') {
				value = !value;
			} else if (value !== undefined && isNull(value)) {
				value = NULL;
			} else {
				const number = value === undefined ? undefined : wholeNumber(value);
				value = number === undefined ? undefined : ~number;
			}
		}
		return value;
	}

	private parseComparison(): ConditionalValue | undefined {
		let leftStart = this.index;
		let left = this.parseConcat();
		let leftEnd = this.index;
		for (;;) {
			const relational = relationalOperatorAt(this.tokens, this.index);
			const word = relational ? undefined : tokenWord(this.peek());
			if (!relational && word !== 'like' && word !== 'is') {
				return left;
			}
			this.index += relational ? relational.length : 1;
			const rightStart = this.index;
			const right = this.parseConcat();
			if (relational) {
				const forLeft = this.literalConversion(rightStart, this.index);
				const forRight = this.literalConversion(leftStart, leftEnd);
				if (forLeft) {
					this.requireConvertible(left, forLeft);
				}
				if (forRight) {
					this.requireConvertible(right, forRight);
				}
			}
			leftStart = this.index;
			leftEnd = this.index;
			if (left === undefined || right === undefined) {
				left = undefined;
			} else if (relational) {
				left = compare(relational.operator, left, right);
			} else {
				left = word === 'like' ? like(left, right) : is(left, right);
			}
		}
	}

	private parseConcat(): ConditionalValue | undefined {
		let left = this.parseAdditive();
		while (this.peek()?.rawText === '&') {
			this.index++;
			const right = this.parseAdditive();
			left = left === undefined || right === undefined ? undefined : concat(left, right);
		}
		return left;
	}

	private parseAdditive(): ConditionalValue | undefined {
		let leftStart = this.index;
		let left = this.parseMod();
		let leftEnd = this.index;
		while (this.peek()?.rawText === '+' || this.peek()?.rawText === '-') {
			const op = this.tokens[this.index++].rawText;
			const rightStart = this.index;
			const right = this.parseMod();
			// `-` reads a number (`UNDEFINED_X - "abc"` is a Type mismatch),
			// or a date beside a date literal; `+` joins two strings, and
			// beside a literal converts as `literalConversion` says
			// (`#1/1/2020# + "abc"` is a Type mismatch).
			const besideRight = this.literalConversion(rightStart, this.index);
			const besideLeft = this.literalConversion(leftStart, leftEnd);
			const forLeft = besideRight === 'date' || op !== '-' ? besideRight : 'number';
			const forRight = besideLeft === 'date' || op !== '-' ? besideLeft : 'number';
			if (forLeft) {
				this.requireConvertible(left, forLeft);
			}
			if (forRight) {
				this.requireConvertible(right, forRight);
			}
			if (left === undefined || right === undefined) {
				left = undefined;
			} else if (op === '+' && typeof left === 'string' && typeof right === 'string') {
				left = left + right;
			} else {
				left = arithmetic(op, left, right);
			}
			// What follows reads the whole sum so far, which is no literal.
			leftStart = this.index;
			leftEnd = this.index;
		}
		return left;
	}

	private parseMod(): ConditionalValue | undefined {
		let left = this.parseIntegerDivision();
		while (this.matchWord('mod')) {
			const right = this.parseIntegerDivision();
			this.requireConvertible(left, 'number');
			this.requireConvertible(right, 'number');
			left = left === undefined || right === undefined ? undefined : arithmetic('mod', left, right);
		}
		return left;
	}

	private parseIntegerDivision(): ConditionalValue | undefined {
		let left = this.parseProduct();
		while (this.peek()?.rawText === '\\') {
			this.index++;
			const right = this.parseProduct();
			this.requireConvertible(left, 'number');
			this.requireConvertible(right, 'number');
			left = left === undefined || right === undefined ? undefined : arithmetic('\\', left, right);
		}
		return left;
	}

	private parseProduct(): ConditionalValue | undefined {
		let left = this.parseNegation();
		while (this.peek()?.rawText === '*' || this.peek()?.rawText === '/') {
			const op = this.tokens[this.index++].rawText;
			const right = this.parseNegation();
			this.requireConvertible(left, 'number');
			this.requireConvertible(right, 'number');
			left = left === undefined || right === undefined ? undefined : arithmetic(op, left, right);
		}
		return left;
	}

	/** Unary minus binds looser than ^: `-2 ^ 2` is -4. */
	private parseNegation(): ConditionalValue | undefined {
		const start = this.index;
		while (this.peek()?.rawText === '-' || this.peek()?.rawText === '+') {
			this.index++;
		}
		const end = this.index;
		let value = this.parsePower();
		if (end > start) {
			this.requireConvertible(value, 'number');
		}
		for (let i = end - 1; i >= start; i--) {
			value = signed(this.tokens[i].rawText, value);
		}
		return value;
	}

	private parsePower(): ConditionalValue | undefined {
		let left = this.parsePrimary();
		while (this.peek()?.rawText === '^') {
			this.index++;
			const right = this.parseNegationOperand();
			this.requireConvertible(left, 'number');
			this.requireConvertible(right, 'number');
			left = left === undefined || right === undefined ? undefined : arithmetic('^', left, right);
		}
		return left;
	}

	/** An exponent may carry its own sign: `2 ^ -1`. */
	private parseNegationOperand(): ConditionalValue | undefined {
		const op = this.peek()?.rawText;
		if (op === '-' || op === '+') {
			this.index++;
			const value = this.parsePrimary();
			this.requireConvertible(value, 'number');
			return signed(op, value);
		}
		return this.parsePrimary();
	}

	private parsePrimary(): ConditionalValue | undefined {
		const token = this.peek();
		if (!token || token.rawText === ')' || token.kind === 'operator') {
			this.broken = true;
			return undefined;
		}
		if (token.rawText === '(') {
			if (this.parenthesisDepth >= MAX_EXPRESSION_DEPTH) {
				// Leave an over-nested directive unknown rather than exhaust the stack.
				this.index = this.tokens.length;
				return undefined;
			}
			this.index++;
			this.parenthesisDepth++;
			const value = this.parseLogical(0);
			this.parenthesisDepth--;
			if (this.peek()?.rawText !== ')') {
				this.broken = true;
				return undefined;
			}
			this.index++;
			return value;
		}
		this.index++;
		if (token.kind === 'integerLiteral') {
			return parseVbaIntegerLiteral(token.rawText);
		}
		if (token.kind === 'floatLiteral') {
			const number = Number(token.rawText.replace(/[!#@]$/, '').replace(/[dD]/, 'e'));
			return Number.isFinite(number) ? number : undefined;
		}
		if (token.kind === 'stringLiteral') {
			if (!/^"(?:[^"]|"")*"$/.test(token.rawText)) {
				// A string left open: the line cannot be read as written.
				this.broken = true;
			}
			return token.rawText.slice(1, -1).replace(/""/g, '"');
		}
		if (token.kind === 'dateLiteral') {
			const serial = dateLiteralSerial(token.rawText);
			return serial === undefined ? undefined : { kind: 'date', serial };
		}
		const word = tokenWord(token);
		switch (word) {
			case 'true': return true;
			case 'false': return false;
			case 'empty': return EMPTY;
			case 'null': return NULL;
			case 'nothing': return NOTHING;
			default: break;
		}
		if (token.kind !== 'identifier') {
			// A keyword where a value belongs (`"abc" - Not`).
			this.broken = true;
		}
		const value = this.constants.get(word);
		if (typeof value === 'string' && this.hostStrings?.has(word)) {
			this.hostString = true;
		}
		if (value === undefined && this.undefinedIsZero && token.kind === 'identifier') {
			return 0;
		}
		return value;
	}

	/**
	 * Notes a string an operator converts that no locale reads as a number,
	 * or, for the logical operators, as a Boolean ("True" Or 0 runs, #191),
	 * or, beside a Date, as a date (`"1/1/2020" = #1/1/2020#` is True).
	 */
	private requireConvertible(value: ConditionalValue | undefined, as: StringConversion): void {
		if (this.mismatch === undefined && typeof value === 'string' && (as === 'number'
			? isInvalidNumericString(value)
			: as === 'boolean' ? isInvalidBooleanString(value) : isInvalidDateString(value))) {
			this.mismatch = { value, as };
		}
	}

	/**
	 * What a string beside the operand written at `start`..`end` is converted
	 * to by `+` or a comparison: one written as a literal (a sign allowed)
	 * decides it, a number as a number, True or False as a Boolean, a date
	 * as a date (Excel 16.0: `1 = "True"` and `"abc" = True` are a Type
	 * mismatch, `"True" = True` is True). Beside a name the string compares
	 * as with a Variant and is not converted (`UNDEFINED_X = "abc"` and
	 * `UNDEFINED_X = ""` compile), nor beside anything else.
	 */
	private literalConversion(start: number, end: number): StringConversion | undefined {
		// A literal in brackets keeps its type: `(1) = "abc"`.
		while (end - start >= 3 && this.tokens[start].rawText === '(' && this.tokens[end - 1].rawText === ')'
			&& this.closingBracket(start) === end - 1) {
			start++;
			end--;
		}
		let at = start;
		while (at < end - 1 && (this.tokens[at].rawText === '-' || this.tokens[at].rawText === '+')) {
			at++;
		}
		if (at !== end - 1) {
			return undefined;
		}
		const token = this.tokens[at];
		if (token.kind === 'integerLiteral' || token.kind === 'floatLiteral') {
			return 'number';
		}
		if (token.kind === 'dateLiteral') {
			return 'date';
		}
		const word = tokenWord(token);
		return word === 'true' || word === 'false' ? 'boolean' : undefined;
	}

	/** The index of the `)` that closes the `(` at `open`. */
	private closingBracket(open: number): number {
		let depth = 0;
		for (let i = open; i < this.tokens.length; i++) {
			const text = this.tokens[i].rawText;
			depth += text === '(' ? 1 : text === ')' ? -1 : 0;
			if (depth === 0) {
				return i;
			}
		}
		return -1;
	}

	private matchWord(word: string): boolean {
		if (tokenWord(this.peek()) !== word) {
			return false;
		}
		this.index++;
		return true;
	}

	private peek(): VbaToken | undefined {
		return this.tokens[this.index];
	}
}

const LOGICAL_LEVELS = ['imp', 'eqv', 'xor', 'or', 'and'] as const;

/** The number a value converts to: True is -1, Empty 0, a Date its serial, a numeric string its number. */
function numberOf(value: ConditionalValue): number | undefined {
	if (typeof value === 'boolean') {
		return value ? -1 : 0;
	}
	if (typeof value === 'number') {
		return value;
	}
	if (typeof value === 'object') {
		return value.kind === 'empty' ? 0 : value.kind === 'date' ? value.serial : undefined;
	}
	// A hex or octal string as VBA reads it (#188): "&H10" is 16, so
	// "&H10" And 1 is 0 (measured in a directive in Excel 16.0). Any other
	// string is read as before, with "." as the decimal point.
	const trimmed = value.trim();
	if (RADIX_STRING.test(trimmed)) {
		return radixStringValue(trimmed);
	}
	const parsed = trimmed.length === 0 ? Number.NaN : Number(trimmed);
	return Number.isFinite(parsed) ? parsed : undefined;
}

/** A hex or octal string: "&H10", "&O17" or "&17". */
const RADIX_STRING = /^&[Hh][0-9A-Fa-f]+$|^&[Oo]?[0-7]+$/;

/**
 * A hex or octal string's value as #188 reads it, or undefined for one past
 * a Long, whose conversion was not measured in a directive.
 */
function radixStringValue(trimmed: string): number | undefined {
	const hex = /^&[Hh]([0-9A-Fa-f]+)$/.exec(trimmed);
	const digits = hex ? hex[1] : trimmed.replace(/^&[Oo]?/, '');
	if (Number.parseInt(digits, hex ? 16 : 8) > 0xFFFFFFFF) {
		return undefined;
	}
	const verdict = numericStringVerdict(trimmed);
	return verdict.kind === 'number' ? verdict.value : undefined;
}

/** A value as the whole number a bitwise operator reads. */
function wholeNumber(value: ConditionalValue): number | undefined {
	const number = numberOf(value);
	return number === undefined ? undefined : bankersRound(number);
}

/**
 * A value as the text & and Like read. Empty and Null are "". A Date's text
 * is the locale's, and a number's is left alone where JavaScript would spell
 * it otherwise than VBA (1E+20, 0.1 + 0.2), so neither is guessed.
 */
function text(value: ConditionalValue): string | undefined {
	if (typeof value === 'boolean') {
		return value ? 'True' : 'False';
	}
	if (typeof value === 'number') {
		return Number.isSafeInteger(value) ? String(value) : undefined;
	}
	if (typeof value === 'string') {
		return value;
	}
	return value.kind === 'empty' || value.kind === 'null' ? '' : undefined;
}

/** `&`: Null is "" beside anything but another Null. */
function concat(left: ConditionalValue, right: ConditionalValue): ConditionalValue | undefined {
	if (isNull(left) && isNull(right)) {
		return NULL;
	}
	const a = text(left);
	const b = text(right);
	return a === undefined || b === undefined ? undefined : a + b;
}

function signed(op: string, value: ConditionalValue | undefined): ConditionalValue | undefined {
	if (value === undefined || isNull(value)) {
		return value;
	}
	const number = numberOf(value);
	return number === undefined ? undefined : op === '-' ? -number : number;
}

/**
 * The bitwise and Boolean operators, with Null as VBA treats it: unknown, so
 * `Null And False` is False and `Null Or True` is True, and anything that
 * depends on the Null is Null.
 */
function logical(op: typeof LOGICAL_LEVELS[number], left: ConditionalValue, right: ConditionalValue): ConditionalValue | undefined {
	if (isNull(left) || isNull(right)) {
		return logicalWithNull(op, left, right);
	}
	if (typeof left === 'boolean' && typeof right === 'boolean') {
		switch (op) {
			case 'and': return left && right;
			case 'or': return left || right;
			case 'xor': return left !== right;
			case 'eqv': return left === right;
			default: return !left || right;
		}
	}
	const a = wholeNumber(left);
	const b = wholeNumber(right);
	if (a === undefined || b === undefined) {
		return undefined;
	}
	switch (op) {
		case 'and': return a & b;
		case 'or': return a | b;
		case 'xor': return a ^ b;
		case 'eqv': return ~(a ^ b);
		default: return ~a | b;
	}
}

function logicalWithNull(op: typeof LOGICAL_LEVELS[number], left: ConditionalValue, right: ConditionalValue): ConditionalValue | undefined {
	const known = isNull(left) ? right : left;
	if (isNull(known)) {
		return NULL;
	}
	// Which value of the known side decides the result alone: every bit clear
	// for And, every bit set for Or. Imp is decided by a False left side or a
	// True right side.
	const bits = typeof known === 'boolean' ? (known ? -1 : 0) : wholeNumber(known);
	if (bits === undefined) {
		return undefined;
	}
	const decided = (result: number): ConditionalValue => (typeof known === 'boolean' ? result !== 0 : result);
	switch (op) {
		case 'and': return bits === 0 ? decided(0) : NULL;
		case 'or': return bits === -1 ? decided(-1) : NULL;
		case 'imp':
			if (known === left) {
				return bits === 0 ? decided(-1) : NULL;
			}
			return bits === -1 ? decided(-1) : NULL;
		default: return NULL;
	}
}

function arithmetic(op: string, left: ConditionalValue, right: ConditionalValue): ConditionalValue | undefined {
	if (isNull(left) || isNull(right)) {
		return NULL;
	}
	const a = numberOf(left);
	const b = numberOf(right);
	if (a === undefined || b === undefined) {
		return undefined;
	}
	const result = numericResult(op, a, b);
	// A Date plus or minus a number is a Date; two Dates subtracted are days.
	const leftDate = isSpecial(left, 'date');
	const rightDate = isSpecial(right, 'date');
	if (result !== undefined && (op === '+' || op === '-') && leftDate !== rightDate && (leftDate || op === '+')) {
		return { kind: 'date', serial: result };
	}
	return result;
}

function numericResult(op: string, a: number, b: number): number | undefined {
	switch (op) {
		case '+': return a + b;
		case '-': return a - b;
		case '*': return a * b;
		case '/': return b === 0 ? undefined : a / b;
		case '\\': {
			const divisor = bankersRound(b);
			return divisor === 0 ? undefined : Math.trunc(bankersRound(a) / divisor);
		}
		case 'mod': {
			const divisor = bankersRound(b);
			return divisor === 0 ? undefined : bankersRound(a) % divisor;
		}
		default: {
			const result = Math.pow(a, b);
			return Number.isFinite(result) ? result : undefined;
		}
	}
}

/**
 * A comparison: two strings compare as text, without regard to case, and so
 * does a string against Empty, which is ""; anything else compares as
 * numbers. Null against anything is Null.
 */
function compare(op: string, left: ConditionalValue, right: ConditionalValue): ConditionalValue | undefined {
	if (isNull(left) || isNull(right)) {
		return NULL;
	}
	const asText = (value: ConditionalValue): string | undefined =>
		typeof value === 'string' ? value : isSpecial(value, 'empty') ? '' : undefined;
	let order: number;
	const textLeft = asText(left);
	const textRight = asText(right);
	if (textLeft !== undefined && textRight !== undefined && (typeof left === 'string' || typeof right === 'string')) {
		const a = textLeft.toLowerCase();
		const b = textRight.toLowerCase();
		order = a < b ? -1 : a > b ? 1 : 0;
	} else {
		const a = numberOf(left);
		const b = numberOf(right);
		if (a === undefined || b === undefined) {
			return undefined;
		}
		order = a < b ? -1 : a > b ? 1 : 0;
	}
	switch (op) {
		case '=': return order === 0;
		case '<>': return order !== 0;
		case '<': return order < 0;
		case '>': return order > 0;
		case '<=': return order <= 0;
		default: return order >= 0;
	}
}

/** `Like`, without regard to case: `"ABC" Like "a*"` and `"a" Like "[A-C]"` are True. */
function like(left: ConditionalValue, right: ConditionalValue): ConditionalValue | undefined {
	if (isNull(left) || isNull(right)) {
		return NULL;
	}
	const subject = text(left);
	const pattern = text(right);
	if (subject === undefined || pattern === undefined) {
		return undefined;
	}
	const regex = likePatternRegex(pattern);
	return regex === undefined ? undefined : regex.test(subject);
}

/**
 * A Like pattern as a regular expression: `?` one character, `*` any run, `#`
 * a digit, and `[...]` a character list, which `!` negates and `a-z` spans.
 * Undefined for a pattern VBA refuses at run time, a `[` never closed.
 */
function likePatternRegex(pattern: string): RegExp | undefined {
	let source = '';
	for (let i = 0; i < pattern.length; i++) {
		const ch = pattern[i];
		if (ch === '?') {
			source += '[\\s\\S]';
		} else if (ch === '*') {
			source += '[\\s\\S]*';
		} else if (ch === '#') {
			source += '[0-9]';
		} else if (ch === '[') {
			const close = pattern.indexOf(']', i + 1);
			if (close < 0) {
				return undefined;
			}
			let list = pattern.slice(i + 1, close);
			const negated = list.startsWith('!');
			if (negated) {
				list = list.slice(1);
			}
			if (list.length === 0) {
				// `[]` matches nothing at all, `[!]` any one character.
				source += negated ? '[\\s\\S]' : '(?!)';
			} else {
				const escaped = list.replace(/[\\\]^]/g, (c) => `\\${c}`);
				source += negated ? `[^${escaped}]` : `[${escaped}]`;
			}
			i = close;
		} else {
			source += ch.replace(/[.*+?^${}()|[\]\\/]/g, (c) => `\\${c}`);
		}
	}
	try {
		return new RegExp(`^${source}$`, 'iu');
	} catch {
		// A range written backwards, `[z-a]`, is a run-time error in VBA too.
		return undefined;
	}
}

/** `Is` compares object references, and the only one a directive can name is Nothing. */
function is(left: ConditionalValue, right: ConditionalValue): ConditionalValue | undefined {
	return isSpecial(left, 'nothing') && isSpecial(right, 'nothing') ? true : undefined;
}
