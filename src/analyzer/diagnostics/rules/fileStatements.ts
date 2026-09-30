// Rule family: file statements whose failure the code proves (issue #123).
//
// Each case was measured in Excel 16.0 (build 20326, 2026-09-26): it compiles
// and raises every time it runs.
//
//  - file-number-zero (52, Bad file name or number): `As #0`, `LOF(0)` - a
//    file number is 1 to 511, and 0 is never one.
//  - file-used-after-close (52): `Close #f` and then `Print #f, ...` on the
//    same number with no Open between.
//  - file-mode-mismatch (54, Bad file mode): `Print #f`/`Write #f` on a file
//    opened For Input; `Input #f`/`Line Input #f` on one opened For Output or
//    Append.
//  - file-already-open (55, File already open): two Opens As the same number
//    with no Close between.
//  - file-record-zero (63, Bad record number): `Seek #f, 0`, `Get #f, 0, x`,
//    `Put #f, 0, x` - records and Binary positions start at 1.
//
// A file number is a literal, or a local that FreeFile fills once. The rule
// follows the top-level statements of a procedure in order; a block between two
// statements that could touch the number ends what is known about it, and a
// number named inside a block is never followed.

import { bareCallStatementTarget } from '../../call/callContext';
import type { ConditionalActivityTracker } from '../../conditional/conditionalCompilation';
import { parseVbaIntegerLiteral } from '../../constants/integerConstantExpression';
import { statementLabelDeclaration } from '../../flow/procedureLabels';
import type { VbaToken } from '../../lexer/tokenKinds';
import type { BodyNode, LeafStatementNode, ModuleNode, Span } from '../../parser/nodes';
import { isLeafStatement } from '../../parser/nodes';
import { walkEnteringBlocks } from '../dataflow';
import type { PushFn } from '../analysisContext';
import {
	activeModuleMembers,
	bareAssignmentTarget,
	statementTokensAfterLeadingLabel,
	tokenName,
	tokenText,
} from '../walker';

type FileMode = 'input' | 'output' | 'append' | 'random' | 'binary';

interface OpenFile {
	mode: FileMode;
	span: Span;
}

/** What is known about each file number key as the statements run. */
type FileStates = Map<string, OpenFile | 'closed'>;

const FILE_STATEMENTS: ReadonlySet<string> = new Set([
	'print', 'write', 'input', 'line', 'get', 'put', 'seek', 'close', 'lock', 'unlock', 'width',
]);

export function checkFileStatements(
	source: string,
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		const states: FileStates = new Map();
		// Blocks are entered with the state they start with; a block may open,
		// close or reopen anything it names (issue #237).
		const visit = (node: BodyNode): void => {
			if (!isLeafStatement(node)) {
				return; // a Dim inside the body declares, and runs nothing
			}
			const toks = statementTokensAfterLeadingLabel(source, node.span);
			if (toks.length === 0) {
				return;
			}
			if (node.kind === 'Statement' && node.singleLineIfBranches) {
				// A single-line If runs its statement on one path only.
				for (const key of fileNumberKeysIn(toks)) {
					states.delete(key);
				}
				return;
			}
			// A label may be reached from anywhere, an error handler's included,
			// so nothing is known there; and a call to a procedure may open or
			// close any file (issue #146).
			if (statementLabelDeclaration(source, node.span) || tokenText(toks[0]) === 'gosub') {
				states.clear();
			}
			if (!isFileStatementHead(tokenText(toks[0])) && bareCallStatementTarget(source, node.span)) {
				states.clear();
				return;
			}
			// `f = FreeFile` again names a new file: what was known about f ends.
			// The value may still name a file number: `Main = LOF(0)`.
			const assigned = bareAssignmentTarget(source, node.span);
			if (assigned) {
				states.delete(assigned.name.toLowerCase());
			}
			checkStatement(node.span, toks, states, push);
		};
		walkEnteringBlocks(source, member.body, (node) => activity?.isInactive(node.span) === true, visit, {
			snapshot: () => new Map(states),
			restore: (saved) => {
				states.clear();
				for (const [key, state] of saved) {
					states.set(key, state);
				}
			},
			forget: (keys) => {
				if (keys.has('*')) {
					states.clear();
				}
				for (const key of keys) {
					states.delete(key);
				}
			},
			touches: (stmt) => fileKeysTouchedBy(source, stmt),
		});
	}
}

function checkStatement(base: Span, toks: readonly VbaToken[], states: FileStates, push: PushFn): void {
	const at = (tok: VbaToken): Span => ({ start: base.start + tok.start, end: base.start + tok.end });
	const head = tokenText(toks[0]);
	// Function forms: LOF(0), EOF(0), Loc(0), FileAttr(0, 1), Seek(0).
	for (let i = 0; i + 2 < toks.length; i++) {
		const name = tokenText(toks[i]);
		if ((name === 'lof' || name === 'eof' || name === 'loc' || name === 'fileattr' || name === 'seek')
			&& toks[i + 1].rawText === '(' && toks[i - 1]?.rawText !== '.' && isZeroLiteral(toks[i + 2])
			&& (toks[i + 3]?.rawText === ')' || toks[i + 3]?.rawText === ',')) {
			push('fileNumberZero', `File number 0 is never open: file numbers run from 1 to 511. This will raise Run-time error '52': Bad file name or number.`, at(toks[i + 2]));
		}
	}
	if (head === 'open') {
		const opened = parseOpen(toks);
		if (!opened) {
			return;
		}
		if (opened.numberToken && isZeroLiteral(opened.numberToken)) {
			push('fileNumberZero', `File number 0 cannot be opened: file numbers run from 1 to 511. This will raise Run-time error '52': Bad file name or number.`, at(opened.numberToken));
			return;
		}
		if (opened.key === undefined) {
			return;
		}
		const previous = states.get(opened.key);
		if (previous !== undefined && previous !== 'closed') {
			push('fileAlreadyOpen', `File number ${describeKey(opened.key)} is still open from the Open statement above; opening it again raises Run-time error '55': File already open. Close it first.`, at(opened.numberToken!));
		}
		states.set(opened.key, { mode: opened.mode, span: base });
		return;
	}
	if (head === 'close' || head === 'reset') {
		const keys = head === 'reset' ? [] : fileNumberKeysIn(toks.slice(1));
		if (keys.length === 0) {
			// `Close` with no number, and `Reset`, close every open file: a
			// later `Print #1` raises 52 in Excel (issue #146).
			for (const [key, state] of states) {
				if (state !== 'closed') {
					states.set(key, 'closed');
				}
			}
			return;
		}
		for (const key of keys) {
			states.set(key, 'closed');
		}
		return;
	}
	if (!FILE_STATEMENTS.has(head)) {
		return;
	}
	// `Line Input #f, x`: the statement word is two tokens.
	const numberIndex = head === 'line' ? (tokenText(toks[1]) === 'input' ? 2 : -1) : 1;
	if (numberIndex < 0) {
		return;
	}
	const numberStart = toks[numberIndex]?.rawText === '#' ? numberIndex + 1 : numberIndex;
	const numberToken = toks[numberStart];
	if (!numberToken) {
		return;
	}
	if (isZeroLiteral(numberToken)) {
		push('fileNumberZero', `File number 0 is never open: file numbers run from 1 to 511. This will raise Run-time error '52': Bad file name or number.`, at(numberToken));
		return;
	}
	const key = fileNumberKey(numberToken);
	if (key === undefined) {
		return;
	}
	const state = states.get(key);
	if (state === 'closed') {
		push('fileUsedAfterClose', `File number ${describeKey(key)} was closed above and not opened again. This will raise Run-time error '52': Bad file name or number.`, at(numberToken));
		return;
	}
	if (state === undefined) {
		return;
	}
	const statement = head === 'line' ? 'line input' : head;
	const writes = statement === 'print' || statement === 'write';
	const reads = statement === 'input' || statement === 'line input';
	if ((writes && state.mode === 'input') || (reads && (state.mode === 'output' || state.mode === 'append'))) {
		const word = statement === 'line input' ? 'Line Input' : head.charAt(0).toUpperCase() + head.slice(1);
		push('fileModeMismatch', `'${word} #' on a file opened For ${modeWord(state.mode)} raises Run-time error '54': Bad file mode.`, at(toks[0]));
		return;
	}
	if ((statement === 'seek' || statement === 'get' || statement === 'put') && (state.mode === 'binary' || state.mode === 'random')) {
		// `Seek #f, 0` / `Get #f, 0, x`: the record or position after the comma.
		const comma = toks.findIndex((tok, index) => index > numberStart && tok.rawText === ',');
		const record = comma > 0 ? toks[comma + 1] : undefined;
		if (record && isZeroLiteral(record) && (toks[comma + 2] === undefined || toks[comma + 2].rawText === ',')) {
			push('fileRecordZero', `${statement.charAt(0).toUpperCase() + statement.slice(1)} with record number 0: records and Binary positions start at 1. This will raise Run-time error '63': Bad record number.`, at(record));
		}
	}
}

function parseOpen(toks: readonly VbaToken[]): { mode: FileMode; key: string | undefined; numberToken: VbaToken | undefined } | undefined {
	let mode: FileMode = 'random';
	let numberToken: VbaToken | undefined;
	let depth = 0;
	for (let i = 1; i < toks.length; i++) {
		const raw = toks[i].rawText;
		if (raw === '(') {
			depth++;
		} else if (raw === ')') {
			depth--;
		}
		if (depth !== 0) {
			continue;
		}
		const word = tokenText(toks[i]);
		if (word === 'for') {
			const next = tokenText(toks[i + 1]);
			if (next === 'input' || next === 'output' || next === 'append' || next === 'random' || next === 'binary') {
				mode = next;
			}
		} else if (word === 'as') {
			numberToken = toks[i + 1]?.rawText === '#' ? toks[i + 2] : toks[i + 1];
			break;
		}
	}
	if (!numberToken) {
		return undefined;
	}
	return { mode, key: fileNumberKey(numberToken), numberToken };
}

/** The key a file number token identifies: its literal value, or the variable's name. */
function fileNumberKey(tok: VbaToken): string | undefined {
	if (tok.kind === 'integerLiteral') {
		// `#&H1` is file 1, not #NaN (issue #146).
		const value = parseVbaIntegerLiteral(tok.rawText);
		return value === undefined ? undefined : `#${value}`;
	}
	const name = tokenName(tok);
	return name ? name.toLowerCase() : undefined;
}

/** True for the statement words this rule follows: Open, Close, Reset and the file I/O statements. */
function isFileStatementHead(head: string): boolean {
	return head === 'open' || head === 'close' || head === 'reset' || FILE_STATEMENTS.has(head);
}

function describeKey(key: string): string {
	return key.startsWith('#') ? key : `'${key}'`;
}

function isZeroLiteral(tok: VbaToken | undefined): boolean {
	return tok?.kind === 'integerLiteral' && /^0+[%&^]?$/.test(tok.rawText);
}

function modeWord(mode: FileMode): string {
	return mode.charAt(0).toUpperCase() + mode.slice(1);
}

/** Every file-number key a statement's tokens name after `#` or `As`. */
function fileNumberKeysIn(toks: readonly VbaToken[]): string[] {
	const out: string[] = [];
	for (let i = 0; i < toks.length; i++) {
		const tok = toks[i];
		if (tok.rawText === '#' && toks[i + 1]) {
			const key = fileNumberKey(toks[i + 1]);
			if (key) {
				out.push(key);
			}
		} else if (i === 0 || toks[i - 1].rawText === ',') {
			const key = fileNumberKey(tok);
			if (key && (toks[i + 1] === undefined || toks[i + 1].rawText === ',')) {
				out.push(key);
			}
		}
	}
	return out;
}

/**
 * The file number keys a statement may open, close or reopen, and the name it
 * assigns; `*` when it may close anything: a procedure call, Reset, or a
 * Close with no number.
 */
function fileKeysTouchedBy(source: string, node: LeafStatementNode): Set<string> {
	const out = new Set<string>();
	const toks = statementTokensAfterLeadingLabel(source, node.span);
	const head = tokenText(toks[0]);
	if (!isFileStatementHead(head) && bareCallStatementTarget(source, node.span)) {
		out.add('*');
	}
	if (isFileStatementHead(head)) {
		for (const key of fileNumberKeysIn(toks)) {
			out.add(key);
		}
		const opened = head === 'open' ? parseOpen(toks) : undefined;
		if (opened?.key) {
			out.add(opened.key);
		}
		if (head === 'reset' || (head === 'close' && fileNumberKeysIn(toks.slice(1)).length === 0)) {
			out.add('*');
		}
	}
	const target = bareAssignmentTarget(source, node.span);
	if (target) {
		out.add(target.name.toLowerCase());
	}
	return out;
}
