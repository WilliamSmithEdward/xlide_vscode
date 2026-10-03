// What a procedure of the module does with an argument passed to it whole
// (issue #449, each measured in Excel 16.0). A variable passed ByRef may come
// back changed, so the state rules forget it at the call. Two kinds of
// callee cannot change it: one whose parameter is ByVal, and one that never
// writes the parameter - no assignment, Set, ReDim, Erase, Input, Get, LSet,
// RSet or Mid statement on it, no For over it, and no call it is passed on
// to. `InitV c` with `ByVal c As Collection`, and `Touch c` that only reads
// c, both leave c Nothing.

import type { VbaToken } from '../lexer/tokenKinds';
import { parseModule } from '../parser/parseModule';
import { resolveRuntimeFunction } from '../runtime/vbaRuntime';
import type { BodyNode, ProcedureNode } from '../parser/nodes';
import { isLeafStatement } from '../parser/nodes';
import { statementAndBranchSpans, statementTokensAfterLeadingLabel, tokenName, tokenText } from './walker';

/** Whether the module's own procedure `callee` leaves an argument in slot `index`, or named `named`, as it was. */
export type CalleeKeepsArgument = (callee: string, index: number, named?: string) => boolean;

const KEEPS_CACHE = new WeakMap<object, Map<string, ProcedureNode | null>>();

const WRITING_HEADS: ReadonlySet<string> = new Set(['set', 'let', 'redim', 'erase', 'input', 'line', 'get', 'lset', 'rset', 'mid', 'mid$', 'midb', 'midb$', 'for']);

/** The module's procedure of that name, when there is exactly one and it is a Sub or Function. */
function procedureNamed(source: string, lower: string): ProcedureNode | undefined {
	const module = parseModule(source);
	let byName = KEEPS_CACHE.get(module);
	if (!byName) {
		byName = new Map();
		for (const member of module.members) {
			if (member.kind !== 'Procedure') {
				continue;
			}
			const key = member.name.toLowerCase();
			byName.set(key, byName.has(key) || (member.procKind !== 'Sub' && member.procKind !== 'Function') ? null : member);
		}
		KEEPS_CACHE.set(module, byName);
	}
	return byName.get(lower) ?? undefined;
}

/** Whether a procedure's body writes the parameter, or passes it on whole to a call. */
function writesParameter(source: string, proc: ProcedureNode, lower: string): boolean {
	let writes = false;
	const visit = (body: readonly BodyNode[]): void => {
		for (const node of body) {
			if (writes) {
				return;
			}
			if (node.kind === 'ForBlock' && node.controlVariable?.toLowerCase() === lower) {
				writes = true;
				return;
			}
			if (isLeafStatement(node)) {
				for (const span of statementAndBranchSpans(node)) {
					if (statementWrites(statementTokensAfterLeadingLabel(source, span), lower, (name) => procedureNamed(source, name) === undefined && resolveRuntimeFunction(name) !== undefined)) {
						writes = true;
						return;
					}
				}
			}
			if ('body' in node && Array.isArray(node.body)) {
				visit(node.body as BodyNode[]);
			}
			if (node.kind === 'IfBlock') {
				for (const branch of node.branches) {
					visit(branch.body);
				}
			}
		}
	};
	visit(proc.body);
	return writes;
}

/**
 * Whether one statement may write `lower`: as its target, under a writing
 * statement, or passed on whole. A VBA function only reads what it is given:
 * `Debug.Print TypeName(p)` leaves p alone (issue #685, measured in Excel 16.0).
 */
function statementWrites(toks: readonly VbaToken[], lower: string, builtin: (lower: string) => boolean = () => false): boolean {
	const mentions = toks.some((tok) => tokenName(tok)?.toLowerCase() === lower);
	if (!mentions) {
		return false;
	}
	const head = tokenText(toks[0]);
	if (WRITING_HEADS.has(head)) {
		return true;
	}
	// `p = 5` and `p(1) = 5`: the parameter or an element of it.
	if (tokenName(toks[0])?.toLowerCase() === lower && (toks[1]?.rawText === '=' || toks[1]?.rawText === '(')) {
		return true;
	}
	// Passed on whole, `Other p`, `Other x, p` or `x = F(p)`: the next callee
	// may write it. A call statement opens with the callee's name.
	const callStatement = toks[0]?.kind === 'identifier' && !toks.some((tok) => tok.rawText === '=');
	let depth = 0;
	for (let i = 1; i < toks.length; i++) {
		depth += toks[i].rawText === '(' ? 1 : toks[i].rawText === ')' ? -1 : 0;
		if (tokenName(toks[i])?.toLowerCase() !== lower) {
			continue;
		}
		const prev = toks[i - 1]?.rawText;
		const next = toks[i + 1]?.rawText;
		if (prev === '.' || prev === '!' || next === '.' || next === '!' || next === '(' || tokenText(toks[i + 1]) === 'is') {
			continue;
		}
		// After a comma inside a call's parentheses too: `F = G(a, p)` passes p
		// on (issue #665).
		if (prev === '(' || prev === ':=' || (prev === ',' && depth > 0) || (callStatement && (prev === ',' || i === 1))) {
			// The call whose parentheses hold it, when that is a VBA function.
			let open = i - 1;
			for (let level = 0; open >= 0; open--) {
				const raw = toks[open].rawText;
				if (raw === ')') {
					level++;
				} else if (raw === '(' && level-- === 0) {
					break;
				}
			}
			const callee = open > 0 && toks[open - 2]?.rawText !== '.' ? tokenName(toks[open - 1])?.toLowerCase() : undefined;
			if (depth > 0 && callee !== undefined && builtin(callee)) {
				continue;
			}
			return true;
		}
	}
	return false;
}

const ANSWERS = new WeakMap<object, Map<string, boolean>>();

/** A keeps-argument test over the module's own procedures (issue #449). */
export function calleeKeepsArgument(source: string): CalleeKeepsArgument {
	return (callee, index, named) => {
		const proc = procedureNamed(source, callee.toLowerCase());
		if (!proc) {
			return false;
		}
		const param = named !== undefined
			? proc.params.find((p) => p.name.toLowerCase() === named.toLowerCase())
			: proc.params[index];
		if (!param || param.paramArray) {
			return false;
		}
		if (param.byVal) {
			return true;
		}
		let answers = ANSWERS.get(proc);
		if (!answers) {
			answers = new Map();
			ANSWERS.set(proc, answers);
		}
		const key = param.name.toLowerCase();
		let keeps = answers.get(key);
		if (keeps === undefined) {
			keeps = !writesParameter(source, proc, key);
			answers.set(key, keeps);
		}
		return keeps;
	};
}
