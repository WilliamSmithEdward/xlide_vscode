import { tokenize } from '../lexer/tokenize';
import { MAX_EXPRESSION_DEPTH } from '../parser/expressionLimits';
import { isKnownDirectiveFreeModule } from '../parser/moduleParseFacts';
import type { VbaToken } from '../lexer/tokenKinds';
import { relationalOperatorAt, tokenWord } from '../lexer/tokenHelpers';
import { bankersRound, parseVbaIntegerLiteral } from '../constants/integerConstantExpression';
import { dateLiteralSerial } from '../constants/dateLiteral';
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
 * worth. Stacks are as deep as the source nests directives, so the walk is
 * short.
 */
function armsDiverge(a: ConditionalArm | undefined, b: ConditionalArm | undefined): boolean {
	for (let outer = a; outer; outer = outer.parent) {
		for (let inner = b; inner; inner = inner.parent) {
			if (outer.chain === inner.chain) {
				return outer.index !== inner.index;
			}
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
	const stack: ConditionalFrame[] = [];
	let current: ConditionalActivity = 'active';
	let branch: ConditionalArm | undefined;
	let chains = 0;
	const events: ConditionalActivityEvent[] = [];
	for (const { directive } of directives) {
		current = applyConditionalDirective(directive, effectiveEnv, projectConstants, stack, current);
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
	const stack: ConditionalFrame[] = [];
	let current: ConditionalActivity = 'active';

	for (const { directive } of directives) {
		if (directive.span.start >= offset) {
			break;
		}
		current = applyConditionalDirective(directive, effectiveEnv, projectConstants, stack, current);
	}
	return current;
}

/**
 * The `#If` and `#ElseIf` lines whose condition is Null, which the VBE refuses
 * to compile: "Invalid use of Null" (issue #208, Excel 16.0). `#If Null`,
 * `#If Null = 1`, `#If Not Null` and a `#Const N = Null` read by `#If N` all
 * are. Only a line the VBE is sure to evaluate is listed: one in code that
 * is compiled, and for `#ElseIf`, after arms that were all False.
 */
export function nullConditionDirectives(
	module: ModuleNode,
	env: ConditionalCompilationEnvironment = {},
): ConditionalDirectiveNode[] {
	if (!moduleHasConditionalDirectives(module)) {
		return [];
	}
	const effectiveEnv = effectiveConditionalCompilationEnvironment(env);
	const projectConstants = projectConstantsOf(effectiveEnv);
	const stack: ConditionalFrame[] = [];
	let current: ConditionalActivity = 'active';
	const out: ConditionalDirectiveNode[] = [];
	for (const { directive } of collectConditionalDirectives(module)) {
		const frame = stack[stack.length - 1];
		const evaluated = directive.directiveKind === 'If'
			? current === 'active'
			: directive.directiveKind === 'ElseIf'
				&& frame?.parent === 'active' && !frame.seenTrue && !frame.seenUnknown;
		if (evaluated) {
			const value = evaluateWithProjectConstants(directive.conditionRaw, effectiveEnv, projectConstants);
			if (value !== undefined && isNull(value)) {
				out.push(directive);
			}
		}
		current = applyConditionalDirective(directive, effectiveEnv, projectConstants, stack, current);
	}
	return out;
}

function applyConditionalDirective(
	directive: ConditionalDirectiveNode,
	env: ConditionalCompilationEnvironment,
	projectConstants: Map<string, ConditionalValue>,
	stack: ConditionalFrame[],
	current: ConditionalActivity,
): ConditionalActivity {
	switch (directive.directiveKind) {
		case 'Const': {
			// A #Const defines its constant even inside a #If False: the VBE
			// reads every #Const line (issue #192, measured in Excel 16.0).
			if (directive.name) {
				const value = evaluateWithProjectConstants(directive.valueRaw, env, projectConstants);
				if (value !== undefined) {
					projectConstants.set(directive.name.toLowerCase(), value);
				}
			}
			return current;
		}
		case 'If': {
			const condition = conditionActivity(directive, env, projectConstants);
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
			const condition = conditionActivity(directive, env, projectConstants);
			if (frame.seenTrue) {
				frame.current = 'inactive';
			} else if (frame.seenUnknown && condition !== 'inactive') {
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
	const projectConstants = new Map<string, ConditionalValue>();
	for (const [name, value] of Object.entries(env.projectConstants ?? {})) {
		projectConstants.set(name.toLowerCase(), value);
	}
	const constants: ConditionalConstDefinition[] = [];
	for (const { directive } of directives) {
		if (directive.directiveKind !== 'Const' || !directive.name || !directive.nameSpan) {
			continue;
		}
		// The index historically supplies a project table even when the caller
		// did not, so an absent name is zero on this path.
		const value = evaluateWithProjectConstants(directive.valueRaw, env, projectConstants, true);
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
	env: ConditionalCompilationEnvironment,
	projectConstants: ReadonlyMap<string, ConditionalValue>,
): ConditionalActivity {
	const value = evaluateWithProjectConstants(directive.conditionRaw, env, projectConstants);
	const holds = value === undefined ? undefined : truthy(value);
	if (holds === undefined) {
		return 'unknown';
	}
	return holds ? 'active' : 'inactive';
}

function evaluateWithProjectConstants(
	expression: string | undefined,
	env: ConditionalCompilationEnvironment,
	projectConstants: ReadonlyMap<string, ConditionalValue>,
	undefinedIsZero = env.projectConstants !== undefined,
): ConditionalValue | undefined {
	if (!expression?.trim()) {
		return undefined;
	}
	// The module's own `#Const` values ride in `projectConstants` whether or
	// not the caller supplied the project's; only the caller's presence says
	// an absent name is provably undefined (issue #102).
	const compilerConstants = conditionalCompilerConstants({ compilerConstants: env.compilerConstants });
	// The parser only needs lookup. Copying all preceding #Const values here
	// for every directive makes a forward replay quadratic.
	const constants = {
		get: (name: string): ConditionalValue | undefined => projectConstants.has(name)
			? projectConstants.get(name) : compilerConstants.get(name),
	};
	return new ConditionalExpressionParser(
		directiveExpressionTokens(expression),
		constants,
		undefinedIsZero,
	).parse();
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
 * condition ("Invalid use of Null"), and for Nothing ("Invalid use of object").
 */
function truthy(value: ConditionalValue): boolean | undefined {
	if (typeof value === 'boolean') {
		return value;
	}
	if (typeof value === 'number') {
		return value !== 0;
	}
	if (typeof value === 'string') {
		return value.length > 0;
	}
	switch (value.kind) {
		case 'empty': return false;
		case 'date': return value.serial !== 0;
		default: return undefined;
	}
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
	) {}

	parse(): ConditionalValue | undefined {
		const value = this.parseLogical(0);
		return this.index >= this.tokens.length ? value : undefined;
	}

	/** Imp, Eqv, Xor, Or, And, loosest first; each level is left-associative. */
	private parseLogical(level: number): ConditionalValue | undefined {
		if (level === LOGICAL_LEVELS.length) {
			return this.parseNot();
		}
		let left = this.parseLogical(level + 1);
		while (this.matchWord(LOGICAL_LEVELS[level])) {
			const right = this.parseLogical(level + 1);
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
		let left = this.parseConcat();
		for (;;) {
			const relational = relationalOperatorAt(this.tokens, this.index);
			const word = relational ? undefined : tokenWord(this.peek());
			if (!relational && word !== 'like' && word !== 'is') {
				return left;
			}
			this.index += relational ? relational.length : 1;
			const right = this.parseConcat();
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
		let left = this.parseMod();
		while (this.peek()?.rawText === '+' || this.peek()?.rawText === '-') {
			const op = this.tokens[this.index++].rawText;
			const right = this.parseMod();
			if (left === undefined || right === undefined) {
				left = undefined;
			} else if (op === '+' && typeof left === 'string' && typeof right === 'string') {
				left = left + right;
			} else {
				left = arithmetic(op, left, right);
			}
		}
		return left;
	}

	private parseMod(): ConditionalValue | undefined {
		let left = this.parseIntegerDivision();
		while (this.matchWord('mod')) {
			const right = this.parseIntegerDivision();
			left = left === undefined || right === undefined ? undefined : arithmetic('mod', left, right);
		}
		return left;
	}

	private parseIntegerDivision(): ConditionalValue | undefined {
		let left = this.parseProduct();
		while (this.peek()?.rawText === '\\') {
			this.index++;
			const right = this.parseProduct();
			left = left === undefined || right === undefined ? undefined : arithmetic('\\', left, right);
		}
		return left;
	}

	private parseProduct(): ConditionalValue | undefined {
		let left = this.parseNegation();
		while (this.peek()?.rawText === '*' || this.peek()?.rawText === '/') {
			const op = this.tokens[this.index++].rawText;
			const right = this.parseNegation();
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
			left = left === undefined || right === undefined ? undefined : arithmetic('^', left, right);
		}
		return left;
	}

	/** An exponent may carry its own sign: `2 ^ -1`. */
	private parseNegationOperand(): ConditionalValue | undefined {
		const op = this.peek()?.rawText;
		if (op === '-' || op === '+') {
			this.index++;
			return signed(op, this.parsePrimary());
		}
		return this.parsePrimary();
	}

	private parsePrimary(): ConditionalValue | undefined {
		const token = this.peek();
		if (!token) {
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
		const value = this.constants.get(word);
		if (value === undefined && this.undefinedIsZero && token.kind === 'identifier') {
			return 0;
		}
		return value;
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
	const trimmed = value.trim();
	const parsed = trimmed.length === 0 ? Number.NaN : Number(trimmed);
	return Number.isFinite(parsed) ? parsed : undefined;
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
