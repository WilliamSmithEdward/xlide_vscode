import { tokenize } from '../lexer/tokenize';
import type { VbaToken } from '../lexer/tokenKinds';
import { tokenWord } from '../lexer/tokenHelpers';
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
	Win32: false,
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

function collectConditionalActivityEvents(
	module: ModuleNode,
	effectiveEnv: ConditionalCompilationEnvironment,
): ConditionalActivityEvent[] {
	const directives = collectConditionalDirectives(module);
	const projectConstants = new Map<string, ConditionalValue>();
	for (const [name, value] of Object.entries(effectiveEnv.projectConstants ?? {})) {
		projectConstants.set(name.toLowerCase(), value);
	}
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
 * skipped rather than guessed at.
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
		if (!name || !/^[A-Za-z_][A-Za-z0-9_]*$/.test(name)) {
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
	const projectConstants = new Map<string, ConditionalValue>();
	for (const [name, value] of Object.entries(effectiveEnv.projectConstants ?? {})) {
		projectConstants.set(name.toLowerCase(), value);
	}
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
			if (current === 'active' && directive.name) {
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
	return evaluateConditionalExpression(expression, {
		compilerConstants: env.compilerConstants,
		projectConstants: Object.fromEntries(projectConstants),
	});
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

class ConditionalExpressionParser {
	private index = 0;

	constructor(
		private readonly tokens: readonly VbaToken[],
		private readonly constants: ReadonlyMap<string, ConditionalValue>,
	) {}

	parse(): ConditionalValue | undefined {
		const value = this.parseOr();
		return this.index >= this.tokens.length ? value : undefined;
	}

	private parseOr(): ConditionalValue | undefined {
		let left = this.parseAnd();
		while (this.matchWord('or')) {
			const right = this.parseAnd();
			if (left === undefined || right === undefined) {
				return undefined;
			}
			left = truthy(left) || truthy(right);
		}
		return left;
	}

	private parseAnd(): ConditionalValue | undefined {
		let left = this.parseComparison();
		while (this.matchWord('and')) {
			const right = this.parseComparison();
			if (left === undefined || right === undefined) {
				return undefined;
			}
			left = truthy(left) && truthy(right);
		}
		return left;
	}

	private parseComparison(): ConditionalValue | undefined {
		const left = this.parseUnary();
		const op = this.peek()?.rawText;
		// Relational operators (<, >, <=, >=) join the existing equality (=, <>)
		// handling so `#If Win64 >= 1 Then` and friends evaluate. Anything else
		// (Like, etc.) is left to the caller as an unmodeled remainder.
		if (op !== '=' && op !== '<>' && op !== '<' && op !== '>' && op !== '<=' && op !== '>=') {
			return left;
		}
		this.index++;
		const right = this.parseUnary();
		if (left === undefined || right === undefined) {
			return undefined;
		}
		if (op === '=' || op === '<>') {
			const same = normalizedComparisonValue(left) === normalizedComparisonValue(right);
			return op === '=' ? same : !same;
		}
		// Relational comparisons operate on the operands' numeric values (VBA
		// coerces booleans to -1/0); a non-numeric operand stays unmodeled.
		const leftNumber = relationalNumber(left);
		const rightNumber = relationalNumber(right);
		if (leftNumber === undefined || rightNumber === undefined) {
			return undefined;
		}
		switch (op) {
			case '<':
				return leftNumber < rightNumber;
			case '>':
				return leftNumber > rightNumber;
			case '<=':
				return leftNumber <= rightNumber;
			default:
				return leftNumber >= rightNumber;
		}
	}

	private parseUnary(): ConditionalValue | undefined {
		if (this.matchWord('not')) {
			const value = this.parseUnary();
			return value === undefined ? undefined : !truthy(value);
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
			const value = this.parseOr();
			if (this.peek()?.rawText !== ')') {
				return undefined;
			}
			this.index++;
			return value;
		}
		this.index++;
		if (token.kind === 'integerLiteral' || token.kind === 'floatLiteral') {
			// VBA hex (&H10) / octal (&O17) literals are not parseable by Number();
			// strip any trailing type-suffix char and parse with the correct radix
			// before falling back to a decimal/float Number() parse.
			const raw = token.rawText.replace(/[!#@%&^]$/, '');
			const hex = /^&[hH]([0-9A-Fa-f]+)$/.exec(raw);
			if (hex) {
				const value = Number.parseInt(hex[1], 16);
				return Number.isFinite(value) ? value : undefined;
			}
			const octal = /^&[oO]([0-7]+)$/.exec(raw);
			if (octal) {
				const value = Number.parseInt(octal[1], 8);
				return Number.isFinite(value) ? value : undefined;
			}
			const number = Number(raw);
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
		return this.constants.get(word);
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

/** Numeric value used for relational comparisons; undefined for non-numeric strings. */
function relationalNumber(value: ConditionalValue): number | undefined {
	if (typeof value === 'boolean') {
		// VBA coerces True -> -1, False -> 0 for numeric comparison.
		return value ? -1 : 0;
	}
	if (typeof value === 'number') {
		return value;
	}
	const parsed = Number(value);
	return Number.isFinite(parsed) ? parsed : undefined;
}

function normalizedComparisonValue(value: ConditionalValue): string {
	if (typeof value === 'boolean') {
		// VBA numeric values: True = -1, False = 0. A boolean #Const therefore
		// compares equal to its numeric form, so `Mac = 0`, `TWINBASIC = 0`, and
		// `VBA7 = -1` all behave as VBE evaluates them.
		return value ? '-1' : '0';
	}
	return String(value).toLowerCase();
}
