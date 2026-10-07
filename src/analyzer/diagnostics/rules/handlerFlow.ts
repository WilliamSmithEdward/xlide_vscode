// Rule family: runtime errors that control flow itself raises (issue #117).
//
// Each construct below compiles and raises every time it runs, measured in
// Excel 16.0 (build 20326, 2026-09-26):
//
//  - handler-fall-through: a procedure whose normal path runs off the end
//    into its error handler, where `Err.Raise Err.Number` re-raises with no
//    error pending. Err.Number is 0 there, and `Err.Raise 0` is error 5. A
//    handler that ends with Resume Next does not raise; it just runs once too
//    often, so only the re-raising form is reported.
//  - resume-without-error: a `Resume` statement in a procedure that never
//    installs a handler with `On Error GoTo <label>`. Nothing can be pending
//    when it runs, and Resume then raises error 20.
//  - return-without-gosub: a GoSub target entered by falling into it from
//    the statement above, whose `Return` then has no GoSub to return to,
//    error 3.
//  - recursive-property-accessor: a Property Get that reads `Me.Name` for its
//    own Name, or a Property Let/Set that assigns `Name = value` to its own
//    Name (in a Let, the name is the property, not a return variable). Each
//    calls itself without end, error 28.
//  - unbounded-recursion: a Sub or Function that calls itself, or another
//    procedure of the module that calls it back, before anything that could
//    leave (issue #240): `Sub S(): S: End Sub`, `F = F() + 1`, and
//    `F = F(n - 1)` with no base case. Error 28.
//
// Fall-through is proven only from a plain statement directly above the
// label: a block above it may or may not leave the procedure, and nothing is
// reported for it.

import type { ConditionalActivityTracker } from '../../conditional/conditionalCompilation';
import {
	collectProcedureLabelReferences,
	statementLabelDeclarations,
	type VbaProcedureLabel,
	type VbaProcedureLabelReference,
} from '../../flow/procedureLabels';
import type { VbaToken } from '../../lexer/tokenKinds';
import type { BodyNode, LeafStatementNode, ModuleNode, ProcedureNode, Span } from '../../parser/nodes';
import { isLeafStatement } from '../../parser/nodes';
import { splitTopLevelTokenGroups } from '../../lexer/tokenHelpers';
import type { PushFn } from '../analysisContext';
import { straightLineUnreachable } from '../straightLineValues';
import {
	activeModuleMembers,
	bareAssignmentTarget,
	blockHeaderLineSpan,
	firstExecutableTokenIndex,
	matchParenFrom,
	statementAndBranchSpans,
	statementTokens,
	statementTokensAfterLeadingLabel,
	tokenName,
	tokenText,
} from '../walker';

/** One top-level entry of a procedure body: a leaf statement or an opaque block. */
interface TopLevelEntry {
	node: BodyNode;
	leaf: LeafStatementNode | undefined;
	/** The labels this statement declares: none, one, or a line number and a name (`10 L1:`). */
	labels: VbaProcedureLabel[];
}

export function checkHandlerFlow(
	source: string,
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
	className?: string,
	procedureFilter?: (member: ProcedureNode) => boolean,
): void {
	// `New Class1` inside Class1 makes another of this class (issue #613).
	const ownClass = className?.toLowerCase();
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind === 'Procedure' && procedureFilter && !procedureFilter(member)) { continue; }
		if (member.kind !== 'Procedure') {
			continue;
		}
		const entries = topLevelEntries(source, member.body, activity);
		checkResumeWithoutError(source, member, activity, push);
		// A label in a block falls in from the statement above it there, the
		// way one at the top level does (issue #237). The procedure's label
		// references are read once, not once per block (issue #322).
		const references = collectProcedureLabelReferences(source, member, activity);
		for (const body of bodyLists(member.body)) {
			checkFallThroughIntoTargets(source, member, body === member.body ? entries : topLevelEntries(source, body, activity), references, push);
		}
		checkRecursiveProperty(source, member, entries, ownClass, push);
	}
	checkUnboundedRecursion(source, mod, activity, ownClass, push);
}

/**
 * The names a statement at the top of a body makes another way to this
 * object, or to a new one of its class that does the same: `Set o = Me`,
 * `Set o = New Class1` (issue #613). Me is always one.
 */
function noteSelfAlias(toks: readonly VbaToken[], selves: Set<string>, ownClass: string | undefined): void {
	const words = toks.filter((tok) => tok.kind !== 'comment');
	const name = tokenText(words[0]) === 'set' && words[2]?.rawText === '=' ? tokenName(words[1])?.toLowerCase() : undefined;
	if (!name) {
		return;
	}
	const value = words.slice(3);
	const self = (value.length === 1 && selves.has(tokenText(value[0])))
		|| (ownClass !== undefined && value.length === 2 && tokenText(value[0]) === 'new' && tokenText(value[1]) === ownClass);
	if (self) {
		selves.add(name);
	} else {
		selves.delete(name);
	}
}

/** A procedure's body, and every body a block in it holds: each If arm alone. */
function bodyLists(body: readonly BodyNode[]): (readonly BodyNode[])[] {
	const out: (readonly BodyNode[])[] = [body];
	for (const node of body) {
		if (node.kind === 'IfBlock') {
			for (const branch of node.branches) {
				out.push(...bodyLists(branch.body));
			}
		} else if ('body' in node && Array.isArray(node.body)) {
			out.push(...bodyLists(node.body as BodyNode[]));
		}
	}
	return out;
}

function topLevelEntries(
	source: string,
	body: readonly BodyNode[],
	activity: ConditionalActivityTracker | undefined,
): TopLevelEntry[] {
	const out: TopLevelEntry[] = [];
	for (const node of body) {
		if (activity?.isInactive(node.span)) {
			continue;
		}
		if (!isLeafStatement(node)) {
			out.push({ node, leaf: undefined, labels: [] });
			continue;
		}
		out.push({ node, leaf: node, labels: statementLabelDeclarations(source, node.span) });
	}
	return out;
}

/** Whether a plain statement always leaves the place it stands: Exit, GoTo, Return, Resume, End, Err.Raise, Error, Stop. */
function leavesUnconditionally(source: string, stmt: LeafStatementNode): boolean {
	const toks = statementTokensAfterLeadingLabel(source, stmt.span);
	const head = tokenText(toks[0]);
	if (head === 'exit' || head === 'goto' || head === 'return' || head === 'resume' || head === 'end' || head === 'stop' || head === 'error') {
		// `End If` and the like never reach here (they close blocks), so `End`
		// alone is the End statement.
		return true;
	}
	if (head === 'err' && toks[1]?.rawText === '.' && tokenText(toks[2]) === 'raise') {
		return true;
	}
	return false;
}

/** The statements of a label's body: from the label to the next label or the end. */
function labelBody(entries: readonly TopLevelEntry[], index: number): TopLevelEntry[] {
	const out: TopLevelEntry[] = [];
	for (let k = index; k < entries.length; k++) {
		if (k > index && entries[k].labels.length > 0) {
			break;
		}
		out.push(entries[k]);
	}
	return out;
}

function checkFallThroughIntoTargets(
	source: string,
	proc: ProcedureNode,
	entries: readonly TopLevelEntry[],
	references: readonly VbaProcedureLabelReference[],
	push: PushFn,
): void {
	const handlerLabels = new Set(references.filter((ref) => ref.statementKind === 'on-error-goto').map((ref) => ref.key));
	const gosubLabels = new Set(references.filter((ref) => ref.statementKind === 'gosub' || ref.statementKind === 'on-gosub').map((ref) => ref.key));
	const named = new Set(references.map((ref) => ref.key));
	for (let i = 0; i < entries.length; i++) {
		const entry = entries[i];
		if (entry.labels.length === 0) {
			continue;
		}
		// The flow above must be a plain statement that does not leave, and
		// that itself runs; a block above proves nothing. The first statement
		// of the body is entered directly.
		const above = entries[i - 1];
		const fallsIn = above === undefined
			|| (above.leaf !== undefined && !leavesUnconditionally(source, above.leaf) && runs(source, entries, i - 1, named));
		if (!fallsIn) {
			continue;
		}
		const body = labelBody(entries, i);
		const handler = entry.labels.find((label) => handlerLabels.has(label.key));
		if (handler) {
			const reraise = body.find((one) => one.leaf && reraisesPendingError(source, one.leaf));
			const exitsFirst = body.findIndex((one) => one.leaf && leavesUnconditionally(source, one.leaf) && !reraisesPendingError(source, one.leaf));
			if (reraise && (exitsFirst < 0 || body.indexOf(reraise) < exitsFirst)) {
				push(
					'handlerFallThrough',
					`Execution falls into error handler '${handler.text}' with no error pending, and 'Err.Raise Err.Number' then raises with Err.Number 0. This will raise Run-time error '5': Invalid procedure call or argument. Put an Exit ${procedureWord(proc)} before the label.`,
					handler.span,
				);
			}
		}
		const target = entry.labels.find((label) => gosubLabels.has(label.key));
		if (target) {
			const returns = body.find((one) => one.leaf && tokenText(statementTokensAfterLeadingLabel(source, one.leaf.span)[0]) === 'return');
			const exitsFirst = body.findIndex((one) => one.leaf && leavesUnconditionally(source, one.leaf));
			if (returns && body.indexOf(returns) === exitsFirst) {
				push(
					'returnWithoutGosub',
					`Execution falls into GoSub target '${target.text}' from the statement above it, and its Return then has no GoSub to return to. This will raise Run-time error '3': Return without GoSub. Put an Exit ${procedureWord(proc)} before the label.`,
					target.span,
				);
			}
		}
	}
}

/**
 * Whether the top-level entry at `index` can run: false when a statement above
 * it leaves unconditionally with nothing between that a statement can jump
 * to. A label no statement names is no way in: in `Exit Function`,
 * `Skip:`, `x = 1`, the assignment is dead, and falls into nothing (issue
 * #203, measured in Excel 16.0).
 */
function runs(source: string, entries: readonly TopLevelEntry[], index: number, named: ReadonlySet<string>): boolean {
	for (let k = index; k >= 0; k--) {
		const entry = entries[k];
		if (entry.labels.some((label) => named.has(label.key))) {
			return true;
		}
		if (k < index && entry.leaf && leavesUnconditionally(source, entry.leaf)) {
			return false;
		}
		if (!entry.leaf) {
			return true; // a block may or may not leave
		}
	}
	return true;
}

/** `Err.Raise Err.Number` (with or without further arguments). */
function reraisesPendingError(source: string, stmt: LeafStatementNode): boolean {
	const toks = statementTokensAfterLeadingLabel(source, stmt.span);
	return tokenText(toks[0]) === 'err' && toks[1]?.rawText === '.' && tokenText(toks[2]) === 'raise'
		&& tokenText(toks[3]) === 'err' && toks[4]?.rawText === '.' && tokenText(toks[5]) === 'number';
}

function procedureWord(proc: ProcedureNode): string {
	return proc.procKind === 'Sub' ? 'Sub' : proc.procKind === 'Function' ? 'Function' : 'Property';
}

function checkResumeWithoutError(
	source: string,
	proc: ProcedureNode,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	const resumes: Span[] = [];
	let installsHandler = false;
	// A handler below Exit Function that nothing jumps to never runs, as when
	// its On Error line is commented out (issue #421, measured in Excel 16.0).
	let dead: ReadonlySet<BodyNode> | undefined;
	const visit = (body: readonly BodyNode[]): void => {
		for (const node of body) {
			if (activity?.isInactive(node.span)) {
				continue;
			}
			if ('body' in node && Array.isArray(node.body)) {
				visit(node.body as BodyNode[]);
				continue;
			}
			if (!isLeafStatement(node)) {
				continue;
			}
			for (const span of statementAndBranchSpans(node)) {
				const toks = statementTokensAfterLeadingLabel(source, span);
				const head = tokenText(toks[0]);
				if (head === 'on' && toks.some((tok) => tokenText(tok) === 'error') && toks.some((tok) => tokenText(tok) === 'goto')) {
					// `On Error GoTo 0` and `On Error GoTo -1` install nothing; a
					// label does. The lexer gives `-1` as two tokens (issue #142).
					const target = toks[toks.length - 1];
					const zero = target.kind === 'integerLiteral' && /^0+$/.test(target.rawText);
					const minusOne = target.kind === 'integerLiteral' && /^0*1$/.test(target.rawText) && toks[toks.length - 2]?.rawText === '-';
					if (!zero && !minusOne) {
						installsHandler = true;
					}
				}
				if (head === 'resume' && !(dead ??= straightLineUnreachable(source, proc.body, activity)).has(node)) {
					resumes.push({ start: span.start + toks[0].start, end: span.start + toks[0].end });
				}
			}
		}
	};
	visit(proc.body);
	if (installsHandler) {
		return;
	}
	for (const span of resumes) {
		push(
			'resumeWithoutError',
			`'Resume' runs with no error handler installed in this procedure, so no error is pending. This will raise Run-time error '20': Resume without error.`,
			span,
		);
	}
}

/** How an `On Error` statement sets error handling. */
export type OnErrorMode = 'resume-next' | 'goto-label' | 'goto-0' | 'goto-minus-1';

/** The mode an `On [Local] Error` statement sets, from its tokens after any line label. */
export function onErrorMode(toks: readonly VbaToken[]): OnErrorMode | undefined {
	const words = toks.filter((tok) => tok.kind !== 'comment');
	let i = tokenText(words[0]) === 'on' ? 1 : -1;
	if (i < 0) {
		return undefined;
	}
	if (tokenText(words[i]) === 'local') {
		i++;
	}
	if (tokenText(words[i]) !== 'error') {
		return undefined;
	}
	if (tokenText(words[i + 1]) === 'resume' && tokenText(words[i + 2]) === 'next') {
		return 'resume-next';
	}
	if (tokenText(words[i + 1]) !== 'goto') {
		return undefined;
	}
	// The lexer gives `-1` as two tokens (issue #142).
	const target = words.slice(i + 2);
	if (target.length === 1 && target[0].kind === 'integerLiteral' && /^0+$/.test(target[0].rawText)) {
		return 'goto-0';
	}
	if (target.length === 2 && target[0].rawText === '-' && target[1].kind === 'integerLiteral' && /^0*1$/.test(target[1].rawText)) {
		return 'goto-minus-1';
	}
	return 'goto-label';
}

/**
 * The stretches of a procedure where one of its error handlers is running.
 * There, `On Error Resume Next` and `On Error GoTo label` do not take effect,
 * and the next error goes to the caller (issue #199, measured in Excel 16.0).
 *
 * A stretch starts at a label that only `On Error GoTo` names, below a
 * statement that leaves, so that nothing but an error reaches it: a label
 * execution can fall into, or that a GoTo names, runs with no handler active.
 * It ends at an `On Error GoTo -1` anywhere, which ends the handler, or at
 * the next label a statement names, which may be entered from outside; a
 * label nothing names, like the number on every line of numbered code, does
 * not end it. Nor does an Exit inside an If: the code after the If runs only
 * when the If did not leave, and the handler is still running there. Code
 * after a Resume or Exit at the top level is reached only through a named
 * label, so those need no rule of their own.
 */
export function errorHandlerExtents(source: string, proc: ProcedureNode): Span[] {
	// A handler inside a block runs to the procedure's end the same way
	// (issue #237).
	const kindsByLabel = new Map<string, Set<string>>();
	for (const ref of collectProcedureLabelReferences(source, proc, undefined)) {
		const kinds = kindsByLabel.get(ref.key) ?? new Set<string>();
		kinds.add(ref.statementKind);
		kindsByLabel.set(ref.key, kinds);
	}
	return bodyLists(proc.body).flatMap((body) => handlerExtentsIn(source, proc, body, kindsByLabel));
}

function handlerExtentsIn(source: string, proc: ProcedureNode, body: readonly BodyNode[], kindsByLabel: ReadonlyMap<string, ReadonlySet<string>>): Span[] {
	const entries = topLevelEntries(source, body, undefined);
	const out: Span[] = [];
	for (let i = 0; i < entries.length; i++) {
		const entry = entries[i];
		// What names any of the line's labels: `10 H:` is entered by GoTo 10 as well.
		const kinds = new Set(entry.labels.flatMap((label) => [...(kindsByLabel.get(label.key) ?? [])]));
		if (kinds.size !== 1 || !kinds.has('on-error-goto')) {
			continue;
		}
		const above = entries[i - 1];
		if (!above?.leaf || !leavesUnconditionally(source, above.leaf)) {
			continue;
		}
		let end = proc.span.end;
		for (let k = i; k < entries.length; k++) {
			const one = entries[k];
			const entered = k > i && one.labels.some((label) => kindsByLabel.has(label.key));
			if (entered || resetsHandler(source, one.node)) {
				end = one.node.span.start;
				break;
			}
		}
		out.push({ start: entry.node.span.start, end });
	}
	return out;
}

/** Whether a statement, or any statement in a block, is `On Error GoTo -1`. */
function resetsHandler(source: string, node: BodyNode): boolean {
	if ('body' in node && Array.isArray(node.body)) {
		return (node.body as BodyNode[]).some((child) => resetsHandler(source, child));
	}
	if (!isLeafStatement(node)) {
		return false;
	}
	return statementAndBranchSpans(node).some((span) =>
		onErrorMode(statementTokensAfterLeadingLabel(source, span)) === 'goto-minus-1');
}

function checkRecursiveProperty(
	source: string,
	proc: ProcedureNode,
	entries: readonly TopLevelEntry[],
	ownClass: string | undefined,
	push: PushFn,
): void {
	if (proc.procKind !== 'PropertyGet' && proc.procKind !== 'PropertyLet' && proc.procKind !== 'PropertySet') {
		return;
	}
	const lower = proc.name.toLowerCase();
	// Me, and a local set to it (issue #613).
	const selves = new Set(['me']);
	// The statements that run once each: the top level, and the body of a
	// `With Me` there, whose `.Value` is Me's (issue #613).
	const leaves: Array<{ leaf: LeafStatementNode; within: boolean }> = [];
	for (const entry of entries) {
		if (entry.leaf) {
			leaves.push({ leaf: entry.leaf, within: false });
			continue;
		}
		if (entry.node.kind === 'WithBlock') {
			const header = statementTokensAfterLeadingLabel(source, blockHeaderLineSpan(source, entry.node.span));
			if (header.length === 2 && selves.has(tokenText(header[1]))) {
				for (const child of entry.node.body) {
					if (isLeafStatement(child) && !(child.kind === 'Statement' && child.singleLineIfBranches)) {
						leaves.push({ leaf: child, within: true });
					}
				}
			}
		}
	}
	for (const { leaf, within } of leaves) {
		const own = statementTokensAfterLeadingLabel(source, leaf.span);
		if (!within) {
			noteSelfAlias(own, selves, ownClass);
		}
		// Inside `With Me`, a leading dot is Me's.
		const toks = within && own[0]?.rawText === '.' ? [{ ...own[0], kind: 'identifier' as const, rawText: 'Me', end: own[0].start }, ...own]
			: within ? own.flatMap((tok, i) => (tok.rawText === '.' && i > 0 && !['identifier', 'bracketedIdentifier'].includes(own[i - 1].kind) && own[i - 1].rawText !== ')'
				? [{ ...tok, kind: 'identifier' as const, rawText: 'Me', end: tok.start }, tok] : [tok]))
			: own;
		const entry = { leaf };
		if (proc.procKind === 'PropertyGet') {
			// `Name = Me.Name`: the bare Name is the return variable, `Me.Name`
			// the property, which is this procedure. `Me.Name = 7` assigns,
			// and calls the Let (issue #613).
			for (let i = 0; i + 2 < toks.length; i++) {
				const assigned = toks[i + 3]?.rawText === '=' && (i === 0 || ['then', 'else', 'set', ':'].includes(tokenText(toks[i - 1]) || toks[i - 1].rawText));
				// `Item = Me.Item(i)`: the same arguments again (issue #613).
				const sameArguments = proc.params.length > 0 && toks[i + 3]?.rawText === '(' && sameArgumentsAt(toks, i + 3, proc);
				if (selves.has(tokenText(toks[i])) && toks[i + 1].rawText === '.' && tokenName(toks[i + 2])?.toLowerCase() === lower && !assigned
					&& ((proc.params.length === 0 && toks[i + 3]?.rawText !== '(') || sameArguments)) {
					push(
						'recursivePropertyAccessor',
						`Property Get '${proc.name}' reads 'Me.${proc.name}', which is itself: the call never returns. This will raise Run-time error '28': Out of stack space.`,
						absoluteRange(entry.leaf.span, toks[i], toks[i + 2]),
					);
				}
				// `Value = Value() + 1`: with its parentheses the name calls the
				// property, where bare it is the return variable (issue #338).
				if (tokenName(toks[i])?.toLowerCase() === lower && toks[i - 1]?.rawText !== '.' && toks[i + 1].rawText === '(' && toks[i + 2].rawText === ')'
					&& proc.params.length === 0) {
					push(
						'recursivePropertyAccessor',
						`Property Get '${proc.name}' calls '${proc.name}()', which is itself: the call never returns. This will raise Run-time error '28': Out of stack space.`,
						absoluteRange(entry.leaf.span, toks[i], toks[i + 2]),
					);
				}
			}
			continue;
		}
		// `Me.Value = v` in the Let, `Set Me.Items = v` in the Set: the
		// property assigned through Me is this procedure (issue #338).
		const meAt = tokenText(toks[0]) === 'set' ? 1 : 0;
		const setForm = meAt === 1;
		if (selves.has(tokenText(toks[meAt])) && toks[meAt + 1]?.rawText === '.' && tokenName(toks[meAt + 2])?.toLowerCase() === lower
			&& toks[meAt + 3]?.rawText === '=' && setForm === (proc.procKind === 'PropertySet') && proc.params.length === 1) {
			push(
				'recursivePropertyAccessor',
				`${proc.procKind === 'PropertyLet' ? 'Property Let' : 'Property Set'} '${proc.name}' assigns 'Me.${proc.name}', which is itself: the call never returns. This will raise Run-time error '28': Out of stack space.`,
				absoluteRange(entry.leaf.span, toks[meAt], toks[meAt + 2]),
			);
			continue;
		}
		// Property Let/Set: `Name = value` assigns the property, which is this
		// procedure, and Property Let has no return variable to mean instead.
		const target = bareAssignmentTarget(source, entry.leaf.span);
		const first = firstExecutableTokenIndex(toks);
		if (target && target.name.toLowerCase() === lower && proc.params.length === 1) {
			push(
				'recursivePropertyAccessor',
				`${proc.procKind === 'PropertyLet' ? 'Property Let' : 'Property Set'} '${proc.name}' assigns '${proc.name}', which is itself: the call never returns. This will raise Run-time error '28': Out of stack space.`,
				{ start: entry.leaf.span.start + toks[first].start, end: entry.leaf.span.start + toks[first].end },
			);
		}
	}
}

/** Words that may leave a procedure, or hand an error to a handler, before a call is reached. */
const LEAVING_WORDS: ReadonlySet<string> = new Set(['exit', 'goto', 'gosub', 'return', 'resume', 'end', 'stop', 'error', 'raise', 'on']);

/** The words after End that close a block rather than end the program. */
const BLOCK_ENDS: ReadonlySet<string> = new Set(['if', 'select', 'with', 'sub', 'function', 'property', 'type', 'enum']);

/** Whether any of these tokens may leave the procedure or install a handler. */
function mayLeave(toks: readonly VbaToken[]): boolean {
	return toks.some((tok, i) => LEAVING_WORDS.has(tokenText(tok)) && !(tokenText(tok) === 'end' && BLOCK_ENDS.has(tokenText(toks[i + 1]))));
}

/** A call one procedure makes before anything could leave it. */
interface FirstCall {
	callee: string;
	span: Span;
}

/**
 * A Sub or Function whose every run calls itself, directly or through other
 * procedures of the module that do the same, never returns (issue #240,
 * measured in Excel 16.0): error 28. A call counts only at the top of the
 * body, ahead of any statement or block that could leave or install a
 * handler, and outside a single-line If.
 */
function checkUnboundedRecursion(
	source: string,
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
	ownClass: string | undefined,
	push: PushFn,
): void {
	const procedures = new Map<string, ProcedureNode>();
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind === 'Procedure' && (member.procKind === 'Sub' || member.procKind === 'Function')) {
			procedures.set(member.name.toLowerCase(), member);
		}
	}
	const firstCalls = new Map<string, FirstCall>();
	for (const [lower, proc] of procedures) {
		const call = firstUnconditionalCall(source, proc, procedures, activity, ownClass);
		if (call) {
			firstCalls.set(lower, call);
		}
	}
	for (const [lower, call] of firstCalls) {
		// Follow the first calls until one repeats; report each procedure on the cycle.
		const path = [lower];
		let next = call.callee;
		while (firstCalls.has(next) && !path.includes(next)) {
			path.push(next);
			next = firstCalls.get(next)!.callee;
		}
		if (next !== lower) {
			continue;
		}
		const name = procedures.get(lower)!.name;
		// 'A' calls 'B', which calls 'C', which calls 'A'.
		const chain = [...path.slice(1), lower].map((next) => `'${procedures.get(next)!.name}'`);
		const through = path.length === 1 ? 'calls itself' : `calls ${chain.join(', which calls ')},`;
		push(
			'unboundedRecursion',
			`'${name}' ${through} before anything could make it return: the calls never end. This will raise Run-time error '28': Out of stack space.`,
			call.span,
		);
	}
}

function firstUnconditionalCall(
	source: string,
	proc: ProcedureNode,
	procedures: ReadonlyMap<string, ProcedureNode>,
	activity: ConditionalActivityTracker | undefined,
	ownClass?: string,
): FirstCall | undefined {
	const selves = new Set(['me']);
	for (const entry of topLevelEntries(source, proc.body, activity)) {
		if (entry.node.kind === 'VariableGroup') {
			continue;
		}
		if (!entry.leaf) {
			// A block may leave inside.
			if (mayLeave(statementTokens(source, entry.node.span))) {
				return undefined;
			}
			continue;
		}
		const toks = statementTokensAfterLeadingLabel(source, entry.leaf.span);
		if (mayLeave(toks)) {
			return undefined;
		}
		if (tokenText(toks[0]) === 'if') {
			continue; // a single-line If runs its call on some paths only
		}
		noteSelfAlias(toks, selves, ownClass);
		const call = procedureCallIn(toks, procedures, selves);
		if (call) {
			return { callee: call.callee, span: absoluteRange(entry.leaf.span, call.first, call.last) };
		}
	}
	return undefined;
}

/**
 * The first call a statement makes to a Sub or Function of the module: a
 * call statement (`S`, `S 1`, `Call S(1)`), or a Function with its
 * parentheses in an expression (`F()`, `F(n - 1)`). Inside F, a bare `F` is
 * the return value. `F(1)`, when F takes no arguments, calls a Variant F and
 * indexes what comes back (measured in Excel 16.0).
 */
function procedureCallIn(
	toks: readonly VbaToken[],
	procedures: ReadonlyMap<string, ProcedureNode>,
	selves: ReadonlySet<string> = new Set(['me']),
): { callee: string; first: VbaToken; last: VbaToken } | undefined {
	// `CallByName Me, "Go", VbMethod` calls Go (issue #613, measured in Excel
	// 16.0).
	const byName = toks.findIndex((tok, i) => tokenText(tok) === 'callbyname' && toks[i - 1]?.rawText !== '.');
	if (byName >= 0) {
		const open = toks[byName + 1]?.rawText === '(' ? byName + 1 : -1;
		const close = open > 0 ? matchParenFrom(toks, open) : toks.length;
		const args = splitTopLevelTokenGroups(toks, open > 0 ? open + 1 : byName + 1, ',', close);
		const target = args[1]?.length === 1 && args[1][0].kind === 'stringLiteral' ? args[1][0].rawText.slice(1, -1).toLowerCase() : undefined;
		const callee = target ? procedures.get(target) : undefined;
		const callType = args[2]?.map((tok) => tokenText(tok)).join('');
		if (args.length === 3 && args[0].length === 1 && selves.has(tokenText(args[0][0])) && callee && (callType === 'vbmethod' || callType === '1')
			&& !callee.modifiers.some((modifier) => modifier.toLowerCase() === 'private') && callee.params.length === 0) {
			return { callee: target!, first: toks[byName], last: args[1][0] };
		}
	}
	const head = tokenText(toks[0]) === 'call' ? 1 : 0;
	const headName = tokenName(toks[head])?.toLowerCase();
	const headProc = headName ? procedures.get(headName) : undefined;
	const assigns = toks.some((tok) => tok.rawText === '=') && head === 0;
	if (headProc && !assigns && toks[head + 1]?.rawText !== '.' && toks[head + 1]?.rawText !== '!') {
		return { callee: headName!, first: toks[0], last: toks[head] };
	}
	// `Me.Go`, `Twice = Me.Twice`: through Me a member is always called, never
	// the return variable (issue #338, measured in Excel 16.0). Me reaches a
	// Public or Friend member only.
	for (let i = 2; i < toks.length; i++) {
		const lower = tokenName(toks[i])?.toLowerCase();
		const callee = lower ? procedures.get(lower) : undefined;
		if (callee && toks[i - 1].rawText === '.' && selves.has(tokenText(toks[i - 2])) && toks[i - 3]?.rawText !== '.'
			&& !callee.modifiers.some((modifier) => modifier.toLowerCase() === 'private')
			&& (callee.params.length === 0 || toks[i + 1]?.rawText === '(')) {
			return { callee: lower!, first: toks[i - 2], last: toks[i] };
		}
	}
	for (let i = 1; i < toks.length - 1; i++) {
		const lower = tokenName(toks[i])?.toLowerCase();
		const callee = lower ? procedures.get(lower) : undefined;
		if (!callee || toks[i + 1].rawText !== '(' || toks[i - 1].rawText === '.' || toks[i - 1].rawText === '!') {
			continue;
		}
		// With arguments it takes none of, a Variant Function calls itself and
		// indexes what comes back; a typed one does not compile.
		const variant = !callee.returnType || callee.returnType.trim().toLowerCase() === 'variant';
		if (callee.params.length === 0 && toks[i + 2]?.rawText !== ')' && !variant) {
			continue;
		}
		return { callee: lower!, first: toks[i], last: toks[i] };
	}
	return undefined;
}

/** Whether the argument list opening at `open` passes the procedure's own parameters, in order. */
function sameArgumentsAt(toks: readonly VbaToken[], open: number, proc: ProcedureNode): boolean {
	const close = matchParenFrom(toks, open);
	const args = close > open + 1 ? splitTopLevelTokenGroups(toks, open + 1, ',', close) : [];
	return args.length === proc.params.length && args.every((arg, k) => arg.length === 1 && tokenName(arg[0])?.toLowerCase() === proc.params[k].name.toLowerCase());
}

function absoluteRange(base: Span, first: VbaToken, last: VbaToken): Span {
	return { start: base.start + first.start, end: base.start + last.end };
}
