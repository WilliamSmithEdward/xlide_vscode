// Rule family: what earlier statements of a procedure leave in Excel, and
// the 1004 a later one then raises (issue #308). Measured in Excel 16.0
// (build 20430, 2026-10-02); each compiles and raises every time it runs.
//
//  - `Application.CutCopyMode = False` empties Excel's clipboard, and a
//    PasteSpecial onto a range before anything is copied again raises 1004,
//    "PasteSpecial method of Range class failed". A Copy, a Cut, or a call
//    into code that may copy ends what is known.
//  - Two sheets the procedure added, `Set w = Worksheets.Add`, cannot share
//    a name: sheet names ignore case, so `w1.Name = "Aa"` then
//    `w2.Name = "aa"` raises 1004, "That name is already taken".

import type { ConditionalActivityTracker } from '../../conditional/conditionalCompilation';
import type { MemberCompletionContext } from '../../completion/memberAccess';
import { resolveReceiverTypeAt } from '../../completion/memberAccess';
import { statementLabelDeclaration } from '../../flow/procedureLabels';
import type { VbaToken } from '../../lexer/tokenKinds';
import type { BodyNode, ModuleNode } from '../../parser/nodes';
import { isLeafStatement } from '../../parser/nodes';
import type { PushFn } from '../analysisContext';
import { walkEnteringBlocks } from '../dataflow';
import { stringLiteralValue } from '../typeInference';
import {
	activeModuleMembers,
	setAssignmentTarget,
	statementTokensAfterLeadingLabel,
	tokenName,
	tokenText,
} from '../walker';
import { namesIn } from './shared';

/** The state key for an empty clipboard; no VBA name holds a `#`. */
const CLIPBOARD = '#clipboard';

/** Words that may put something on the clipboard, or run code that does. */
const CLIPBOARD_WORDS: ReadonlySet<string> = new Set(['copy', 'cut', 'copypicture', 'run', 'sendkeys', 'execute', 'doevents', 'call']);

export function checkExcelSessionState(
	source: string,
	mod: ModuleNode,
	callables: ReadonlySet<string>,
	memberCtx: MemberCompletionContext,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	const host = memberCtx.model?.hostName;
	if (host !== undefined && host !== 'Excel') {
		return;
	}
	// Anything that may copy, or a call into the project's own code.
	const mayCopy = (toks: readonly VbaToken[]): boolean => toks.some((tok) => {
		const word = tokenText(tok);
		return CLIPBOARD_WORDS.has(word) || (tokenName(tok) !== undefined && callables.has(word));
	});
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		// A sheet the procedure added, by variable, and the name it gave it ('' unknown).
		const state = new Map<string, string>();
		const forget = (names: Iterable<string>): void => {
			for (const lower of names) {
				state.delete(lower);
			}
		};
		const visit = (node: BodyNode): void => {
			if (!isLeafStatement(node)) {
				return;
			}
			const toks = statementTokensAfterLeadingLabel(source, node.span).filter((tok) => tok.kind !== 'comment');
			if (statementLabelDeclaration(source, node.span) || tokenText(toks[0]) === 'gosub') {
				state.clear();
			}
			if (node.kind === 'Statement' && node.singleLineIfBranches) {
				forget(namesIn(source, node.span));
				state.delete(CLIPBOARD);
				return;
			}
			const at = (first: VbaToken, last: VbaToken) => ({ start: node.span.start + first.start, end: node.span.start + last.end });
			// `Application.CutCopyMode = False` (or 0).
			if (toks.length === 5 && tokenText(toks[0]) === 'application' && toks[1].rawText === '.' && tokenText(toks[2]) === 'cutcopymode'
				&& toks[3].rawText === '=' && (tokenText(toks[4]) === 'false' || toks[4].rawText === '0')) {
				state.set(CLIPBOARD, '');
				return;
			}
			if (mayCopy(toks)) {
				state.delete(CLIPBOARD);
			}
			const paste = toks.findIndex((tok, k) => tokenText(tok) === 'pastespecial' && toks[k - 1]?.rawText === '.');
			if (paste > 0 && state.has(CLIPBOARD) && resolveReceiverTypeAt(source, node.span.start + toks[paste - 1].end, memberCtx) === 'Excel.Range') {
				push('pasteWithNothingCopied', `Nothing is copied here: Application.CutCopyMode = False emptied the clipboard, and no Copy or Cut came after it. This will raise Run-time error '1004': PasteSpecial method of Range class failed.`, at(toks[paste], toks[paste]));
			}
			// `Set w = Worksheets.Add`: a sheet of its own.
			const set = setAssignmentTarget(source, node.span);
			if (set) {
				const lower = set.name.toLowerCase();
				forget(namesIn(source, node.span));
				const value = set.valueTokens.filter((tok) => tok.kind !== 'comment').map((tok) => tokenText(tok)).join('');
				if (/^(?:(?:activeworkbook|thisworkbook|application)\.)?(?:worksheets|sheets)\.add(?:\(\))?$/.test(value)) {
					state.set(lower, '');
				}
				return;
			}
			// `w.Name = "Aa"`.
			const sheet = toks.length === 5 && toks[1].rawText === '.' && tokenText(toks[2]) === 'name' && toks[3].rawText === '='
				? tokenName(toks[0])?.toLowerCase() : undefined;
			if (sheet !== undefined && state.has(sheet)) {
				const name = toks[4].kind === 'stringLiteral' ? stringLiteralValue(toks[4].rawText) : '';
				const taken = name === '' ? undefined : [...state].find(([other, held]) => other !== sheet && other !== CLIPBOARD && held.toLowerCase() === name.toLowerCase());
				if (taken) {
					push('sheetNameInvalid', `"${name}" is the name the code gave '${taken[0]}', another sheet it added, and sheet names ignore case. This will raise Run-time error '1004': That name is already taken.`, at(toks[4], toks[4]));
				}
				state.set(sheet, name);
				return;
			}
			forget([...namesIn(source, node.span)].filter((lower) => lower !== CLIPBOARD));
		};
		walkEnteringBlocks(source, member.body, (node) => activity?.isInactive(node.span) === true, visit, {
			snapshot: () => new Map(state),
			restore: (saved) => {
				state.clear();
				for (const [lower, held] of saved) {
					state.set(lower, held);
				}
			},
			forget,
			touches: (stmt) => {
				const names = new Set(namesIn(source, stmt.span));
				if (mayCopy(statementTokensAfterLeadingLabel(source, stmt.span))) {
					names.add(CLIPBOARD);
				}
				return names;
			},
		});
	}
}
