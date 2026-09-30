import { tokenize } from '../lexer/tokenize';
import type { VbaToken } from '../lexer/tokenKinds';
import { relationalOperatorAt, tokenWord } from '../lexer/tokenHelpers';
import { bankersRound, parseVbaIntegerLiteral } from '../constants/integerConstantExpression';
import type {
	BodyNode,
	ConditionalDirectiveNode,
	ModuleNode,
	ProcedureNode,
	Span,
} from '../parser/nodes';

export type ConditionalValue = boolean | number | string;
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

const DEFAULT_COMPILER_CONSTANTS: Readonly<Record<string, ConditionalValue>> = {
	VBA7: true,
	Win64: true,
	// Win32 is True in 64-bit Office as well: it means Windows, not a width
	// (issue #192, measured in 64-bit Excel 16.0). Win16 is False.
	Win32: true,
	Win16: false,
	Mac: false,
	// `TWINBASIC` is a compiler auto-constant defined only by the twinBASIC
	// compiler; in Excel VBA it is undefined and therefore False (VBE-oracle
	// verified). Modern VBA libraries gate twinBASIC-only intrinsics behind
	// `#If TWINBASIC Then ...`, so without this default those (inactive) branches
	// were analyzed and produced false positives. Unlike a genuine unprovable
	// host flag, TWINBASIC's value in VBA is known, so it is a default, not left
	// `unknown`.
	TWINBASIC: false,
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

export function evaluateConditionalExpression(
	expression: string | undefined,
	env: ConditionalCompilationEnvironment = {},
): ConditionalValue | undefined {
	if (!expression?.trim()) {
		return undefined;
	}
	const parser = new ConditionalExpressionParser(
		tokenize(expression).filter((t) => t.kind !== 'comment' && t.kind !== 'newline'),
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
		const value = evaluateConditionalExpression(directive.valueRaw, {
			...env,
			projectConstants: Object.fromEntries(projectConstants),
		});
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
	if (value === undefined) {
		return 'unknown';
	}
	return truthy(value) ? 'active' : 'inactive';
}

function evaluateWithProjectConstants(
	expression: string | undefined,
	env: ConditionalCompilationEnvironment,
	projectConstants: ReadonlyMap<string, ConditionalValue>,
): ConditionalValue | undefined {
	if (!expression?.trim()) {
		return undefined;
	}
	// The module's own `#Const` values ride in `projectConstants` whether or
	// not the caller supplied the project's; only the caller's presence says
	// an absent name is provably undefined (issue #102).
	const constants = conditionalCompilerConstants({ compilerConstants: env.compilerConstants });
	for (const [name, value] of projectConstants) {
		constants.set(name, value);
	}
	return new ConditionalExpressionParser(
		tokenize(expression).filter((t) => t.kind !== 'comment' && t.kind !== 'newline'),
		constants,
		env.projectConstants !== undefined,
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

function truthy(value: ConditionalValue): boolean {
	if (typeof value === 'boolean') {
		return value;
	}
	if (typeof value === 'number') {
		return value !== 0;
	}
	return value.length > 0;
}

/**
 * Evaluates a #If or #Const expression as the VBE does (issue #192, measured
 * in Excel 16.0). The operators and their order are VBA's own, loosest
 * first: Imp, Eqv, Xor, Or, And, Not, the comparisons, &, + and -, Mod, \,
 * * and /, unary minus, ^. Not, And, Or, Xor, Eqv and Imp are bitwise on
 * numbers, as in code: `Not 1` is -2, which is True, and `1 And 2` is 0.
 * Two Booleans give a Boolean. Strings compare without regard to case:
 * `"A" = "a"` is True. Hex and octal literals keep their width: &HFFFF is -1.
 */
class ConditionalExpressionParser {
	private index = 0;

	constructor(
		private readonly tokens: readonly VbaToken[],
		private readonly constants: ReadonlyMap<string, ConditionalValue>,
		/**
		 * Whether a name no constant defines evaluates as the VBE evaluates
		 * it, to Empty (0 here, since Empty compares as 0 and is False). True
		 * only when the caller supplied the project's own conditional
		 * constants, so an absent name is provably undefined rather than
		 * unknown (issue #102); a module's `#Const` lines are folded into the
		 * same table before any `#If` reads them.
		 */
		private readonly undefinedIsEmpty: boolean,
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
		if (this.matchWord('not')) {
			const value = this.parseNot();
			if (typeof value === 'boolean') {
				return !value;
			}
			const number = value === undefined ? undefined : wholeNumber(value);
			return number === undefined ? undefined : ~number;
		}
		return this.parseComparison();
	}

	private parseComparison(): ConditionalValue | undefined {
		let left = this.parseConcat();
		for (;;) {
			const relational = relationalOperatorAt(this.tokens, this.index);
			if (!relational) {
				return left;
			}
			this.index += relational.length;
			const right = this.parseConcat();
			left = left === undefined || right === undefined ? undefined : compare(relational.operator, left, right);
		}
	}

	private parseConcat(): ConditionalValue | undefined {
		let left = this.parseAdditive();
		while (this.peek()?.rawText === '&') {
			this.index++;
			const right = this.parseAdditive();
			left = left === undefined || right === undefined ? undefined : `${text(left)}${text(right)}`;
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
		const op = this.peek()?.rawText;
		if (op === '-' || op === '+') {
			this.index++;
			const value = this.parseNegation();
			const number = value === undefined ? undefined : numberOf(value);
			return number === undefined ? undefined : op === '-' ? -number : number;
		}
		return this.parsePower();
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
			const value = this.parsePrimary();
			const number = value === undefined ? undefined : numberOf(value);
			return number === undefined ? undefined : op === '-' ? -number : number;
		}
		return this.parsePrimary();
	}

	private parsePrimary(): ConditionalValue | undefined {
		const token = this.peek();
		if (!token) {
			return undefined;
		}
		if (token.rawText === '(') {
			this.index++;
			const value = this.parseLogical(0);
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
		const word = tokenWord(token);
		if (word === 'true') {
			return true;
		}
		if (word === 'false') {
			return false;
		}
		const value = this.constants.get(word);
		if (value === undefined && this.undefinedIsEmpty && token.kind === 'identifier') {
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

/** The number a value converts to: True is -1, a numeric string its number. */
function numberOf(value: ConditionalValue): number | undefined {
	if (typeof value === 'boolean') {
		return value ? -1 : 0;
	}
	if (typeof value === 'number') {
		return value;
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

function text(value: ConditionalValue): string {
	if (typeof value === 'boolean') {
		return value ? 'True' : 'False';
	}
	return String(value);
}

function logical(op: typeof LOGICAL_LEVELS[number], left: ConditionalValue, right: ConditionalValue): ConditionalValue | undefined {
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

function arithmetic(op: string, left: ConditionalValue, right: ConditionalValue): number | undefined {
	const a = numberOf(left);
	const b = numberOf(right);
	if (a === undefined || b === undefined) {
		return undefined;
	}
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

/** A comparison: two strings compare as text, without regard to case; anything else as numbers. */
function compare(op: string, left: ConditionalValue, right: ConditionalValue): boolean | undefined {
	let order: number;
	if (typeof left === 'string' && typeof right === 'string') {
		const a = left.toLowerCase();
		const b = right.toLowerCase();
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
