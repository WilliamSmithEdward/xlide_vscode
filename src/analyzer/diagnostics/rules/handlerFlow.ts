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
//
// Fall-through is proven only from a plain statement directly above the
// label: a block above it may or may not leave the procedure, and nothing is
// reported for it.

import type { ConditionalActivityTracker } from '../../conditional/conditionalCompilation';
import {
	collectProcedureLabelReferences,
	statementLabelDeclaration,
} from '../../flow/procedureLabels';
import type { VbaToken } from '../../lexer/tokenKinds';
import type { BodyNode, LeafStatementNode, ModuleNode, ProcedureNode, Span } from '../../parser/nodes';
import { isLeafStatement } from '../../parser/nodes';
import type { PushFn } from '../analysisContext';
import {
	activeModuleMembers,
	bareAssignmentTarget,
	firstExecutableTokenIndex,
	statementAndBranchSpans,
	statementTokensAfterLeadingLabel,
	tokenName,
	tokenText,
} from '../walker';

/** One top-level entry of a procedure body: a leaf statement or an opaque block. */
interface TopLevelEntry {
	node: BodyNode;
	leaf: LeafStatementNode | undefined;
	/** The label this statement declares, lower-cased, when it does. */
	label?: string;
	labelSpan?: Span;
}

export function checkHandlerFlow(
	source: string,
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		const entries = topLevelEntries(source, member.body, activity);
		checkResumeWithoutError(source, member, activity, push);
		checkFallThroughIntoTargets(source, member, entries, activity, push);
		checkRecursiveProperty(source, member, entries, push);
	}
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
			out.push({ node, leaf: undefined });
			continue;
		}
		const label = statementLabelDeclaration(source, node.span);
		out.push({ node, leaf: node, label: label?.key, labelSpan: label?.span });
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
		if (k > index && entries[k].label !== undefined) {
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
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	const references = collectProcedureLabelReferences(source, proc, activity);
	const handlerLabels = new Set(references.filter((ref) => ref.statementKind === 'on-error-goto').map((ref) => ref.key));
	const gosubLabels = new Set(references.filter((ref) => ref.statementKind === 'gosub' || ref.statementKind === 'on-gosub').map((ref) => ref.key));
	for (let i = 0; i < entries.length; i++) {
		const entry = entries[i];
		if (entry.label === undefined || !entry.labelSpan) {
			continue;
		}
		// The flow above must be a plain statement that does not leave; a
		// block above proves nothing. The first statement of the body is
		// entered directly.
		const above = entries[i - 1];
		const fallsIn = above === undefined || (above.leaf !== undefined && !leavesUnconditionally(source, above.leaf));
		if (!fallsIn) {
			continue;
		}
		const body = labelBody(entries, i);
		if (handlerLabels.has(entry.label)) {
			const reraise = body.find((one) => one.leaf && reraisesPendingError(source, one.leaf));
			const exitsFirst = body.findIndex((one) => one.leaf && leavesUnconditionally(source, one.leaf) && !reraisesPendingError(source, one.leaf));
			if (reraise && (exitsFirst < 0 || body.indexOf(reraise) < exitsFirst)) {
				push(
					'handlerFallThrough',
					`Execution falls into error handler '${labelText(source, entry)}' with no error pending, and 'Err.Raise Err.Number' then raises with Err.Number 0. This will raise Run-time error '5': Invalid procedure call or argument. Put an Exit ${procedureWord(proc)} before the label.`,
					entry.labelSpan,
				);
			}
		}
		if (gosubLabels.has(entry.label)) {
			const returns = body.find((one) => one.leaf && tokenText(statementTokensAfterLeadingLabel(source, one.leaf.span)[0]) === 'return');
			const exitsFirst = body.findIndex((one) => one.leaf && leavesUnconditionally(source, one.leaf));
			if (returns && body.indexOf(returns) === exitsFirst) {
				push(
					'returnWithoutGosub',
					`Execution falls into GoSub target '${labelText(source, entry)}' from the statement above it, and its Return then has no GoSub to return to. This will raise Run-time error '3': Return without GoSub. Put an Exit ${procedureWord(proc)} before the label.`,
					entry.labelSpan,
				);
			}
		}
	}
}

/** `Err.Raise Err.Number` (with or without further arguments). */
function reraisesPendingError(source: string, stmt: LeafStatementNode): boolean {
	const toks = statementTokensAfterLeadingLabel(source, stmt.span);
	return tokenText(toks[0]) === 'err' && toks[1]?.rawText === '.' && tokenText(toks[2]) === 'raise'
		&& tokenText(toks[3]) === 'err' && toks[4]?.rawText === '.' && tokenText(toks[5]) === 'number';
}

function labelText(source: string, entry: TopLevelEntry): string {
	return entry.labelSpan ? source.slice(entry.labelSpan.start, entry.labelSpan.end) : entry.label ?? '';
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
				if (head === 'resume') {
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
	const entries = topLevelEntries(source, proc.body, undefined);
	const kindsByLabel = new Map<string, Set<string>>();
	for (const ref of collectProcedureLabelReferences(source, proc, undefined)) {
		const kinds = kindsByLabel.get(ref.key) ?? new Set<string>();
		kinds.add(ref.statementKind);
		kindsByLabel.set(ref.key, kinds);
	}
	const out: Span[] = [];
	for (let i = 0; i < entries.length; i++) {
		const entry = entries[i];
		const kinds = entry.label === undefined ? undefined : kindsByLabel.get(entry.label);
		if (!kinds || kinds.size !== 1 || !kinds.has('on-error-goto')) {
			continue;
		}
		const above = entries[i - 1];
		if (!above?.leaf || !leavesUnconditionally(source, above.leaf)) {
			continue;
		}
		let end = proc.span.end;
		for (let k = i; k < entries.length; k++) {
			const one = entries[k];
			const entered = k > i && one.label !== undefined && kindsByLabel.has(one.label);
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
	push: PushFn,
): void {
	if (proc.procKind !== 'PropertyGet' && proc.procKind !== 'PropertyLet' && proc.procKind !== 'PropertySet') {
		return;
	}
	const lower = proc.name.toLowerCase();
	for (const entry of entries) {
		if (!entry.leaf) {
			continue;
		}
		const toks = statementTokensAfterLeadingLabel(source, entry.leaf.span);
		if (proc.procKind === 'PropertyGet') {
			// `Name = Me.Name`: the bare Name is the return variable, `Me.Name`
			// the property, which is this procedure.
			for (let i = 0; i + 2 < toks.length; i++) {
				if (tokenText(toks[i]) === 'me' && toks[i + 1].rawText === '.' && tokenName(toks[i + 2])?.toLowerCase() === lower
					&& !(proc.params.length > 0) && toks[i + 3]?.rawText !== '(') {
					push(
						'recursivePropertyAccessor',
						`Property Get '${proc.name}' reads 'Me.${proc.name}', which is itself: the call never returns. This will raise Run-time error '28': Out of stack space.`,
						absoluteRange(entry.leaf.span, toks[i], toks[i + 2]),
					);
				}
			}
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

function absoluteRange(base: Span, first: VbaToken, last: VbaToken): Span {
	return { start: base.start + first.start, end: base.start + last.end };
}
