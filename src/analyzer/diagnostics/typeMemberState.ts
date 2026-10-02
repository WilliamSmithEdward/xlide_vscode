// What a local of a user-defined type holds in its members, statement by
// statement (issues #248 and #253, each measured in Excel 16.0): a dynamic
// array field has no elements until a ReDim gives it bounds, and none again
// after Erase; an object field is Nothing until a Set; a numeric field is 0
// until an assignment stores a literal.
//
// Only a local of a module Type is followed, not Static, from its Dim. A
// whole use of the local or of a field that holds more than one value
// (`Fill t`, `t = u`, `FillArr t.dyn`, `Take t.o`, `Bump t.a`), LSet, RSet,
// Input, Get, a label and a GoSub end what is known. Blocks are entered as
// issue #237 enters them; entering a With leaves its subject alone.

import type { ConditionalActivityTracker } from '../conditional/conditionalCompilation';
import type { VbaToken } from '../lexer/tokenKinds';
import type { BodyNode, LeafStatementNode, ProcedureNode } from '../parser/nodes';
import { isLeafStatement } from '../parser/nodes';
import type { buildModuleSymbols } from '../symbols/buildModuleSymbols';
import { jumpTargetLabelDeclaration } from '../flow/procedureLabels';
import { procedureSymbolFor } from './analysisContext';
import { walkEnteringBlocks } from './dataflow';
import { literalDimensions, type FixedArrayBound } from './rules/arrays';
import { fieldChain, isFixedArrayField, isLeadingDot, typeKey, withSubject, withSubjectsIn, type FieldStep, type ModuleTypes, type WithSubject } from './typeFields';
import { isInactiveNode, statementTokensAfterLeadingLabel, tokenName, tokenText } from './walker';

/** A numeric field's value. */
export interface KnownNumber {
	number: number;
}

/** A dynamic array field with no elements, an object field that is Nothing, a dynamic array's bounds, or a number. */
export type MemberState = 'unallocated' | 'nothing' | FixedArrayBound | KnownNumber;

export function isArrayBounds(state: MemberState | undefined): state is FixedArrayBound {
	return typeof state === 'object' && 'dims' in state;
}

export function isKnownNumber(state: MemberState | undefined): state is KnownNumber {
	return typeof state === 'object' && 'number' in state;
}

/** The member states a statement sees; a single-line If's branch sees them less what its condition names. */
export type MemberStatesAt = (stmt: LeafStatementNode, offset: number) => ReadonlyMap<string, MemberState>;

const NUMBER_TYPES: ReadonlySet<string> = new Set(['byte', 'integer', 'long', 'longlong', 'longptr', 'currency', 'single', 'double', 'decimal']);

const NONE: ReadonlyMap<string, MemberState> = new Map();

export function typeMemberStatesAt(
	source: string,
	symbols: ReturnType<typeof buildModuleSymbols>,
	proc: ProcedureNode,
	types: ModuleTypes,
	activity: ConditionalActivityTracker | undefined,
	optionBase: number,
	isObjectType: (type: string) => boolean = () => false,
): MemberStatesAt {
	const out = new Map<LeafStatementNode, ReadonlyMap<string, MemberState>>();
	const branches = new Map<LeafStatementNode, { then: number; states: ReadonlyMap<string, MemberState> }>();
	const at: MemberStatesAt = (stmt, offset) => {
		const branch = branches.get(stmt);
		return branch && offset > branch.then ? branch.states : out.get(stmt) ?? NONE;
	};
	const roots = new Map<string, string>();
	let states = new Map<string, MemberState>();
	for (const child of procedureSymbolFor(symbols, proc)?.children ?? []) {
		const type = typeKey(child.asType);
		if (child.kind !== 'localVariable' || child.visibility === 'Static' || child.isArray || !type || !types.has(type)) {
			continue;
		}
		const lower = child.name.toLowerCase();
		roots.set(lower, type);
		for (const [key, state] of initialStates(types, type, lower, isObjectType, 0)) {
			states.set(key, state);
		}
	}
	if (states.size === 0) {
		return at;
	}
	const subjects = withSubjectsIn(source, proc, activity, symbols, types);
	// True while a recorded statement or a snapshot holds the current map, which is then copied before a change.
	let shared = false;
	const set = (key: string, state: MemberState | undefined): void => {
		if (shared) {
			states = new Map(states);
			shared = false;
		}
		if (state === undefined) {
			states.delete(key);
		} else {
			states.set(key, state);
		}
	};
	const forget = (names: Iterable<string>): void => {
		for (const name of names) {
			for (const key of [...states.keys()]) {
				if (key === name || key.startsWith(`${name}.`)) {
					set(key, undefined);
				}
			}
		}
	};
	/** The tracked locals a statement names, and its With's local where it reaches the subject with a `.`. */
	const rootsIn = (toks: readonly VbaToken[], subject: WithSubject | undefined): Set<string> => {
		const named = new Set<string>();
		for (let i = 0; i < toks.length; i++) {
			const lower = tokenName(toks[i])?.toLowerCase();
			if (lower && roots.has(lower) && toks[i - 1]?.rawText !== '.' && toks[i - 1]?.rawText !== '!') {
				named.add(lower);
			}
			const subjectRoot = subject?.path?.split('.')[0];
			if (toks[i].rawText === '.' && subjectRoot && roots.has(subjectRoot) && isLeadingDot(toks, i)) {
				named.add(subjectRoot);
			}
		}
		return named;
	};
	/** The fields a chain at `i` reaches from a tracked local, or from the With's subject. */
	const chainAt = (toks: readonly VbaToken[], i: number, subject: WithSubject | undefined): { steps: FieldStep[]; root?: string } | undefined => {
		if (toks[i].rawText === '.') {
			const subjectRoot = subject?.path?.split('.')[0];
			return subject?.type && subjectRoot && roots.has(subjectRoot) && isLeadingDot(toks, i)
				? { steps: fieldChain(toks, { type: subject.type, path: subject.path, display: subject.display, dot: i }, types) }
				: undefined;
		}
		const lower = tokenName(toks[i])?.toLowerCase();
		const type = lower ? roots.get(lower) : undefined;
		if (!lower || !type || toks[i - 1]?.rawText === '.' || toks[i - 1]?.rawText === '!') {
			return undefined;
		}
		return toks[i + 1]?.rawText === '.'
			? { steps: fieldChain(toks, { type, path: lower, display: toks[i].rawText, dot: i + 1 }, types), root: lower }
			: { steps: [], root: lower };
	};
	const visit = (node: BodyNode): void => {
		if (!isLeafStatement(node)) {
			return;
		}
		if (jumpTargetLabelDeclaration(source, node.span)) {
			forget(roots.keys());
		}
		const toks = statementTokensAfterLeadingLabel(source, node.span);
		const subject = subjects.get(node.span.start);
		out.set(node, states);
		shared = true;
		const head = tokenText(toks[0]);
		if (head === 'gosub') {
			forget(roots.keys());
			return;
		}
		// A single-line If runs its branch only sometimes, and its condition may
		// guard it: `If Not t.o Is Nothing Then t.o.Add 1`.
		const singleLineIf = (node.kind === 'Statement' && node.singleLineIfBranches !== undefined) || node.singleLineIfTail === true;
		if (singleLineIf) {
			const then = toks.findIndex((tok) => tokenText(tok) === 'then');
			if (then > 0) {
				forget(rootsIn(toks.slice(0, then), subject));
				branches.set(node, { then: node.span.start + toks[then].start, states });
				shared = true;
			}
			forget(rootsIn(toks, subject));
			return;
		}
		if (head === 'redim' || head === 'erase') {
			const start = tokenText(toks[1]) === 'preserve' ? 2 : 1;
			for (const group of splitGroups(toks.slice(start))) {
				const chain = chainAt(group, 0, subject);
				const step = chain?.steps.at(-1);
				const tracked = step?.path !== undefined && states.has(step.path);
				if (tracked && head === 'erase' && step!.open === undefined && step!.at === group.length - 1) {
					set(step!.path!, 'unallocated');
					continue;
				}
				// A literal ReDim of a field under Option Base 0; Option Base 1 was not measured.
				const dims = tracked && head === 'redim' && step!.open !== undefined && optionBase === 0 ? literalDimensions(group.slice(step!.open + 1, step!.close), 0) : undefined;
				if (dims) {
					set(step!.path!, { name: step!.path!, dims, origin: 'ReDim' });
					continue;
				}
				forget(tracked ? [step!.path!] : rootsIn(group, subject));
			}
			return;
		}
		if (['lset', 'rset', 'input', 'get', 'line', 'mid', 'mid$'].includes(head)) {
			forget(rootsIn(toks, subject));
			return;
		}
		// `Set t.o = ...` and `t.a = 5`: what the target holds after the statement.
		const isSet = head === 'set';
		const targetAt = isSet || head === 'let' ? 1 : 0;
		const target = toks[targetAt] ? chainAt(toks, targetAt, subject)?.steps.at(-1) : undefined;
		const targetEnd = target ? target.close ?? target.at : -1;
		const assigns = target !== undefined && toks[targetEnd + 1]?.rawText === '=' && target.open === undefined;
		for (let i = assigns ? targetEnd + 2 : 0; i < toks.length; i++) {
			const chain = chainAt(toks, i, subject);
			if (!chain) {
				continue;
			}
			const last = chain.steps.at(-1);
			if (!last) {
				if (chain.root) {
					forget([chain.root]); // the local used whole
				}
				continue;
			}
			const end = last.close ?? last.at;
			const whole = last.open === undefined && toks[end + 1]?.rawText !== '.' && toks[end + 1]?.rawText !== '!' && toks[end + 1]?.rawText !== '(';
			const bound = ['ubound', 'lbound'].includes(tokenText(toks[i - 2])) && toks[i - 1]?.rawText === '(';
			const isTypeValue = !last.field.isArray && last.field.type !== undefined && types.has(last.field.type);
			if (last.path && whole && !bound && (isTypeValue || last.field.isArray || passedWhole(toks, i, end))) {
				forget([last.path]);
			}
		}
		if (!assigns || !target.path || !states.has(target.path) && !isTrackedScalar(target, isObjectType)) {
			return;
		}
		const value = toks.slice(targetEnd + 2).filter((tok) => tok.kind !== 'comment');
		if (isSet) {
			set(target.path, value.length === 1 && tokenText(value[0]) === 'nothing' && isObjectType(target.field.type ?? '') ? 'nothing' : undefined);
			return;
		}
		const number = target.field.type && NUMBER_TYPES.has(target.field.type) && !target.field.isArray ? literalNumber(value) : undefined;
		set(target.path, number === undefined ? undefined : { number });
	};
	walkEnteringBlocks(source, proc.body, (node) => isInactiveNode(activity, node), visit, {
		snapshot: () => {
			shared = true;
			return states;
		},
		restore: (saved) => {
			states = new Map(saved);
			shared = false;
		},
		forget,
		touches: (stmt) => {
			const toks = statementTokensAfterLeadingLabel(source, stmt.span);
			const subject = subjects.get(stmt.span.start);
			// Entering `With t` evaluates t, and changes nothing.
			if (tokenText(toks[0]) === 'with' && withSubject(toks, symbols, proc, types, subject)) {
				return [];
			}
			return rootsIn(toks, subject);
		},
	});
	return at;
}

/** A scalar or object field an assignment can give a state to. */
function isTrackedScalar(step: FieldStep, isObjectType: (type: string) => boolean): boolean {
	const type = step.field.type;
	return !step.field.isArray && type !== undefined && (NUMBER_TYPES.has(type) || isObjectType(type));
}

/** Each member a local starts with a state for: dynamic arrays, objects and numbers, through fields that are not arrays. */
function initialStates(
	types: ModuleTypes,
	type: string,
	prefix: string,
	isObjectType: (type: string) => boolean,
	depth: number,
): Array<[string, MemberState]> {
	const out: Array<[string, MemberState]> = [];
	for (const [lower, field] of types.get(type) ?? []) {
		const key = `${prefix}.${lower}`;
		if (field.isArray) {
			if (!isFixedArrayField(field)) {
				out.push([key, 'unallocated']);
			}
		} else if (field.type && types.has(field.type)) {
			if (depth < 8) {
				out.push(...initialStates(types, field.type, key, isObjectType, depth + 1));
			}
		} else if (field.type && NUMBER_TYPES.has(field.type)) {
			out.push([key, { number: 0 }]);
		} else if (field.type && isObjectType(field.type)) {
			out.push([key, 'nothing']);
		}
	}
	return out;
}

/**
 * Whether the chain from `start` to `end` stands alone in an argument slot,
 * where a call may take it ByRef: `Bump t.a`, `Take(t.o)`, `x = F(t.a)`.
 */
function passedWhole(toks: readonly VbaToken[], start: number, end: number): boolean {
	const prev = toks[start - 1];
	const next = toks[end + 1];
	const opens = prev === undefined || prev.rawText === '(' || prev.rawText === ',' || prev.kind === 'identifier' || (prev.kind === 'keyword' && tokenText(prev) === 'call');
	const closes = next === undefined || next.rawText === ')' || next.rawText === ',' || next.rawText === ':' || next.kind === 'comment';
	return opens && closes;
}

/** The value of a plain number literal, signed or not, or undefined. */
function literalNumber(value: readonly VbaToken[]): number | undefined {
	const signed = value.length === 2 && (value[0].rawText === '-' || value[0].rawText === '+');
	const literal = value.length === 1 ? value[0] : signed ? value[1] : undefined;
	if (!literal || (literal.kind !== 'integerLiteral' && literal.kind !== 'floatLiteral')) {
		return undefined;
	}
	const raw = literal.rawText.replace(/[%&^!#@]$/, '');
	const number = /^&[hHoO]/.test(raw) ? undefined : Number(raw.replace(/[dD]/, 'E'));
	return number === undefined || !Number.isFinite(number) ? undefined : value[0].rawText === '-' ? -number : number;
}

/** A statement's comma-separated groups outside parentheses. */
function splitGroups(toks: readonly VbaToken[]): VbaToken[][] {
	const out: VbaToken[][] = [[]];
	let depth = 0;
	for (const tok of toks) {
		if (tok.rawText === '(') {
			depth++;
		} else if (tok.rawText === ')') {
			depth--;
		} else if (tok.rawText === ',' && depth === 0) {
			out.push([]);
			continue;
		}
		if (tok.kind !== 'comment') {
			out[out.length - 1].push(tok);
		}
	}
	return out.filter((group) => group.length > 0);
}
