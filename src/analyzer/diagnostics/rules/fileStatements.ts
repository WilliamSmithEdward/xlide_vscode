// Rule family: file statements whose failure the code proves (issue #123).
//
// Each case was measured in Excel 16.0 (build 20326, 2026-09-26): it compiles
// and raises every time it runs.
//
//  - file-number-zero (52, Bad file name or number): `As #0`, `LOF(0)`, and
//    any literal past the numbers Open takes, 1 to 512: `As #513`,
//    `Close #-1` (issue #262).
//  - file-used-after-close (52): `Close #f` and then `Print #f, ...` on the
//    same number with no Open between.
//  - file-mode-mismatch (54, Bad file mode): `Print #f`/`Write #f` on a file
//    opened For Input; `Input #f`/`Line Input #f` on one opened For Output or
//    Append.
//  - file-already-open (55, File already open): two Opens As the same number
//    with no Close between.
//  - file-record-zero (63, Bad record number): `Seek #f, 0`, `Get #f, 0, x`,
//    `Put #f, 0, x` - records and Binary positions start at 1. Seek raises
//    in any mode, and on a number nothing opened (issue #262).
//  - file-read-past-end (62, Input past end of file): reading a file this
//    procedure created empty - opened For Output, closed with nothing
//    written, and opened For Input from the same path - before anything
//    checks EOF or LOF (issue #262).
//  - Open's `Len = 0` raises 5 in every mode, and a Len literal past 32767
//    raises 6 (issue #262); those report as runtime-argument-value and
//    arithmetic-overflow.
//
// A file number is a literal, or a local that FreeFile fills once. The rule
// follows the top-level statements of a procedure in order; a block between two
// statements that could touch the number ends what is known about it, and a
// number named inside a block is never followed.

import { bareCallStatementTarget } from '../../call/callContext';
import type { ConditionalActivityTracker } from '../../conditional/conditionalCompilation';
import { parseVbaIntegerLiteral } from '../../constants/integerConstantExpression';
import { jumpTargetLabelDeclaration } from '../../flow/procedureLabels';
import type { VbaToken } from '../../lexer/tokenKinds';
import type { BodyNode, LeafStatementNode, ModuleNode, Span } from '../../parser/nodes';
import { isLeafStatement } from '../../parser/nodes';
import { trackedLocalsNamedWhole, walkEnteringBlocks } from '../dataflow';
import type { PushFn } from '../analysisContext';
import { mergeOpenedFileNumbers, openedFileNumbersIn, type OpenedFileNumbers } from '../openedFileNumbers';
import { stringLiteralValue } from '../typeInference';
import {
	activeModuleMembers,
	bareAssignmentTarget,
	blockHeaderLineSpan,
	forEachStatement,
	statementAndBranchSpans,
	statementTokensAfterLeadingLabel,
	tokenName,
	tokenText,
} from '../walker';

type FileMode = 'input' | 'output' | 'append' | 'random' | 'binary';

interface OpenFile {
	mode: FileMode;
	span: Span;
	/** The `path:` key of the path it was opened from, when that is a name or a literal. */
	path?: string;
	/** For Output: whether anything was written. */
	written?: boolean;
	/** For Input: the file is known to be empty, and nothing has checked EOF or LOF yet. */
	emptyUnchecked?: boolean;
	/** The `path:` key it was opened from, in any mode, while that name still holds it. */
	openPath?: string;
}

/**
 * The modes each statement works in; any other raises 54, Bad file mode
 * (issue #419, measured in Excel 16.0). Input and Line Input read a Binary
 * file too; Get and Put need Binary or Random.
 */
const STATEMENT_MODES: Readonly<Record<string, readonly FileMode[]>> = {
	print: ['output', 'append'],
	write: ['output', 'append'],
	input: ['input', 'binary'],
	'line input': ['input', 'binary'],
	get: ['binary', 'random'],
	put: ['binary', 'random'],
};

/**
 * What is known about each file number key as the statements run, and under
 * `path:` keys the paths known to name an empty file.
 */
type FileStates = Map<string, OpenFile | 'closed' | 'empty'>;

const FILE_STATEMENTS: ReadonlySet<string> = new Set([
	'print', 'write', 'input', 'line', 'get', 'put', 'seek', 'close', 'lock', 'unlock', 'width',
]);

/** The highest file number Open takes: As #512 runs and As #513 raises 52 (issue #262). */
const MAX_FILE_NUMBER = 512;

/** The functions that read a file's state without reading from it. */
const FILE_STATE_FUNCTIONS: ReadonlySet<string> = new Set(['lof', 'eof', 'loc', 'fileattr', 'seek']);

/** VBA functions that only read a name passed to them. */
const READ_ONLY_INTRINSICS: ReadonlySet<string> = new Set(['len', 'lenb', 'dir', 'filelen', 'filedatetime', 'getattr']);

export function checkFileStatements(
	source: string,
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
	projectOpened?: OpenedFileNumbers,
): void {
	if (projectOpened) {
		checkUnopenedNumbers(source, mod, activity, push, mergeOpenedFileNumbers([projectOpened, openedFileNumbersIn(source)]));
	}
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
				forgetPathsNamedIn(states, toks);
				return;
			}
			// A label may be reached from anywhere, an error handler's included,
			// so nothing is known there; and a call to a procedure may open or
			// close any file (issue #146).
			if (jumpTargetLabelDeclaration(source, node.span) || tokenText(toks[0]) === 'gosub') {
				states.clear();
			}
			if (checkOpenPathUse(node.span, toks, states, push)) {
				return;
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
				forgetPath(states, `path:${assigned.name.toLowerCase()}`);
			}
			// A path passed whole to a procedure may come back changed; a file
			// statement only reads it.
			if (!isFileStatementHead(tokenText(toks[0]))) {
				const tracked = (name: string): boolean => states.has(`path:${name}`) || [...states.values()].some((state) => typeof state === 'object' && state.openPath === `path:${name}`);
				for (const lower of trackedLocalsNamedWhole(toks, node.span.start, tracked, READ_ONLY_INTRINSICS).keys()) {
					forgetPath(states, `path:${lower}`);
				}
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
					states.delete(`path:${key}`);
				}
			},
			touches: (stmt) => fileKeysTouchedBy(source, stmt),
			// `Do Until EOF(f)` checks before its body reads.
			enter: (node) => markChecked(states, statementTokensAfterLeadingLabel(source, blockHeaderLineSpan(source, node.span))),
		});
	}
}

/** File statements that raise 52 on a number nothing opened; Close runs (issue #419). */
const NUMBERED_STATEMENTS: ReadonlySet<string> = new Set(['print', 'write', 'input', 'line', 'get', 'put', 'seek', 'lock', 'unlock', 'width']);

/**
 * A literal file number no Open in the project names, while none names a
 * variable or FreeFile: `Print #1, "x"` and `EOF(1)` raise 52, "Bad file
 * name or number", wherever they run (issue #419, measured in Excel 16.0).
 * The project's Opens come from the index, and this module's from its text
 * as it stands.
 */
function checkUnopenedNumbers(source: string, mod: ModuleNode, activity: ConditionalActivityTracker | undefined, push: PushFn, opened: OpenedFileNumbers): void {
	if (opened.any) {
		return;
	}
	const unopened = (tok: VbaToken | undefined): number | undefined => {
		const value = tok?.kind === 'integerLiteral' && /^\d+$/.test(tok.rawText) ? Number(tok.rawText) : undefined;
		return value !== undefined && value >= 1 && value <= MAX_FILE_NUMBER && !opened.numbers.has(value) ? value : undefined;
	};
	const report = (base: Span, tok: VbaToken, value: number): void => {
		push('fileNumberZero', `File number ${value} is opened by no Open statement in this project, so nothing can be open on it. This will raise Run-time error '52': Bad file name or number.`, { start: base.start + tok.start, end: base.start + tok.end });
	};
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		forEachStatement(member.body, (stmt) => {
			for (const span of statementAndBranchSpans(stmt)) {
				const toks = statementTokensAfterLeadingLabel(source, span).filter((tok) => tok.kind !== 'comment' && tokenText(tok) !== 'else');
				const head = tokenText(toks[0]);
				const numberAt = head === 'line' ? (tokenText(toks[1]) === 'input' ? 2 : -1) : NUMBERED_STATEMENTS.has(head) ? 1 : -1;
				if (numberAt > 0 && toks[numberAt]?.rawText === '#') {
					const value = unopened(toks[numberAt + 1]);
					if (value !== undefined) {
						report(span, toks[numberAt + 1], value);
						continue;
					}
				}
				// `EOF(1)`, `LOF(1)`, `Input(1, #1)`.
				for (let i = 0; i + 2 < toks.length; i++) {
					const name = tokenText(toks[i]);
					if (toks[i + 1].rawText !== '(' || toks[i - 1]?.rawText === '.') {
						continue;
					}
					const number = FILE_STATE_FUNCTIONS.has(name) ? toks[i + 2]
						: name === 'input' && toks[i + 3]?.rawText === ',' && toks[i + 4]?.rawText === '#' ? toks[i + 5]
						: undefined;
					const value = toks[toks.indexOf(number!) + 1]?.rawText === ')' || toks[toks.indexOf(number!) + 1]?.rawText === ',' ? unopened(number) : undefined;
					if (value !== undefined) {
						report(span, number!, value);
					}
				}
			}
		}, activity);
	}
}

function checkStatement(base: Span, toks: readonly VbaToken[], states: FileStates, push: PushFn): void {
	const at = (tok: VbaToken): Span => ({ start: base.start + tok.start, end: base.start + tok.end });
	const range = (first: VbaToken, last: VbaToken): Span => ({ start: base.start + first.start, end: base.start + last.end });
	const head = tokenText(toks[0]);
	// Function forms: LOF(0), EOF(600), Loc(-1), FileAttr(0, 1), Seek(0).
	for (let i = 0; i + 2 < toks.length; i++) {
		const name = tokenText(toks[i]);
		if (!FILE_STATE_FUNCTIONS.has(name) || toks[i + 1].rawText !== '(' || toks[i - 1]?.rawText === '.') {
			continue;
		}
		const impossible = impossibleFileNumber(toks, i + 2);
		const after = impossible ? toks[toks.indexOf(impossible.last) + 1]?.rawText : undefined;
		if (impossible && (after === ')' || after === ',')) {
			push('fileNumberZero', `File number ${impossible.value} is never open: file numbers run from 1 to ${MAX_FILE_NUMBER}. This will raise Run-time error '52': Bad file name or number.`, range(impossible.first, impossible.last));
		}
	}
	markChecked(states, toks);
	reportEmptyInputFunction(toks, states, push, at);
	if (head === 'open') {
		const opened = parseOpen(toks);
		if (!opened) {
			return;
		}
		const impossible = impossibleFileNumber(toks, opened.numberIndex);
		if (impossible) {
			push('fileNumberZero', `File number ${impossible.value} cannot be opened: file numbers run from 1 to ${MAX_FILE_NUMBER}. This will raise Run-time error '52': Bad file name or number.`, range(impossible.first, impossible.last));
			return;
		}
		if (opened.len) {
			const value = opened.len.value;
			if (value === 0) {
				push('runtimeArgumentValue', `Argument 'Len' of 'Open' is 0; this will raise Run-time error '5': Invalid procedure call or argument.`, at(opened.len.token));
				return;
			}
			if (value > 32767) {
				push('arithmeticOverflow', `Open's Len of ${value} does not fit an Integer. This will raise Run-time error '6': Overflow.`, at(opened.len.token));
				return;
			}
		}
		const empty = opened.path !== undefined && states.get(opened.path) === 'empty';
		if (opened.path !== undefined && opened.mode !== 'input') {
			// Output empties it, and the other modes may write to it.
			states.delete(opened.path);
		}
		if (opened.key === undefined) {
			return;
		}
		const previous = states.get(opened.key);
		if (previous !== undefined && previous !== 'closed' && previous !== 'empty') {
			push('fileAlreadyOpen', `File number ${describeKey(opened.key)} is still open from the Open statement above; opening it again raises Run-time error '55': File already open. Close it first.`, at(toks[opened.numberIndex]));
		}
		states.set(opened.key, {
			mode: opened.mode,
			span: base,
			...(opened.mode === 'output' && opened.path !== undefined ? { path: opened.path, written: false } : {}),
			...(opened.mode === 'input' && empty ? { emptyUnchecked: true } : {}),
			...(opened.path !== undefined ? { openPath: opened.path } : {}),
		});
		return;
	}
	if (head === 'close' || head === 'reset') {
		for (const index of closeNumberIndexes(toks)) {
			const impossible = impossibleFileNumber(toks, index);
			if (impossible) {
				push('fileNumberZero', `File number ${impossible.value} is never open: file numbers run from 1 to ${MAX_FILE_NUMBER}. This will raise Run-time error '52': Bad file name or number.`, range(impossible.first, impossible.last));
				return;
			}
		}
		const keys = head === 'reset' ? [] : fileNumberKeysIn(toks.slice(1));
		if (keys.length === 0) {
			// `Close` with no number, and `Reset`, close every open file: a
			// later `Print #1` raises 52 in Excel (issue #146).
			for (const [key, state] of [...states]) {
				if (state !== 'closed' && state !== 'empty') {
					closeFile(states, key);
				}
			}
			return;
		}
		for (const key of keys) {
			closeFile(states, key);
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
	const impossible = impossibleFileNumber(toks, numberStart);
	if (impossible) {
		push('fileNumberZero', `File number ${impossible.value} is never open: file numbers run from 1 to ${MAX_FILE_NUMBER}. This will raise Run-time error '52': Bad file name or number.`, range(impossible.first, impossible.last));
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
	const statement = head === 'line' ? 'line input' : head;
	// `Seek #f, 0` raises 63 whatever f is open for, and when nothing opened
	// it (issue #262); Get and Put only reach the record in Binary and Random.
	if (statement === 'seek' || ((statement === 'get' || statement === 'put') && state !== undefined && state !== 'empty' && (state.mode === 'binary' || state.mode === 'random'))) {
		const comma = toks.findIndex((tok, index) => index > numberStart && tok.rawText === ',');
		const record = comma > 0 ? recordBelowOne(toks, comma + 1) : undefined;
		if (record && (statement === 'seek' || record.value === 0)) {
			push('fileRecordZero', `${statement.charAt(0).toUpperCase() + statement.slice(1)} with record number ${record.value}: records and Binary positions start at 1. This will raise Run-time error '63': Bad record number.`, range(record.first, record.last));
			return;
		}
	}
	if (state === undefined || state === 'empty') {
		return;
	}
	const writes = statement === 'print' || statement === 'write';
	const reads = statement === 'input' || statement === 'line input';
	const modes = STATEMENT_MODES[statement];
	if (modes && !modes.includes(state.mode)) {
		const word = statement === 'line input' ? 'Line Input' : head.charAt(0).toUpperCase() + head.slice(1);
		push('fileModeMismatch', `'${word} #' on a file opened For ${modeWord(state.mode)} raises Run-time error '54': Bad file mode.`, at(toks[0]));
		return;
	}
	if ((writes || statement === 'put') && state.mode === 'output' && !state.written && !writesNothing(toks, numberStart)) {
		states.set(key, { ...state, written: true });
	}
	if (reads && state.emptyUnchecked) {
		const word = statement === 'line input' ? 'Line Input' : 'Input';
		push('fileReadPastEnd', `'${word} #' reads file ${describeKey(key)}, which this procedure created empty and reopened For Input without checking EOF. This will raise Run-time error '62': Input past end of file.`, at(toks[0]));
		states.set(key, { ...state, emptyUnchecked: false });
	}
}

/** `Input(n, #f)` and `InputB$(n, f)` read the file too. */
function reportEmptyInputFunction(toks: readonly VbaToken[], states: FileStates, push: PushFn, at: (tok: VbaToken) => Span): void {
	for (let i = 0; i + 1 < toks.length; i++) {
		const name = tokenText(toks[i]);
		if ((name !== 'input' && name !== 'inputb') || toks[i - 1]?.rawText === '.' || i === 0) {
			continue;
		}
		const open = toks[i + 1].rawText === '$' ? i + 2 : i + 1;
		if (toks[open]?.rawText !== '(') {
			continue;
		}
		const comma = toks.findIndex((tok, k) => k > open && tok.rawText === ',');
		const numberTok = toks[comma + 1]?.rawText === '#' ? toks[comma + 2] : toks[comma + 1];
		const key = comma > 0 && numberTok ? fileNumberKey(numberTok) : undefined;
		const state = key ? states.get(key) : undefined;
		// `Input(1, #1)` reads an Input or Binary file only (issue #419).
		if (key && state && state !== 'closed' && state !== 'empty' && state.mode !== 'input' && state.mode !== 'binary') {
			push('fileModeMismatch', `'${toks[i].rawText}' reads file ${describeKey(key)}, opened For ${modeWord(state.mode)}. This will raise Run-time error '54': Bad file mode.`, at(toks[i]));
			continue;
		}
		if (key && state && state !== 'closed' && state !== 'empty' && state.emptyUnchecked) {
			push('fileReadPastEnd', `'${toks[i].rawText}' reads file ${describeKey(key)}, which this procedure created empty and reopened For Input without checking EOF. This will raise Run-time error '62': Input past end of file.`, at(toks[i]));
			states.set(key, { ...state, emptyUnchecked: false });
		}
	}
}

/** EOF, LOF, Loc and Seek on a file: a read after them may be guarded. */
function markChecked(states: FileStates, toks: readonly VbaToken[]): void {
	for (let i = 0; i + 2 < toks.length; i++) {
		if (!FILE_STATE_FUNCTIONS.has(tokenText(toks[i])) || toks[i + 1].rawText !== '(') {
			continue;
		}
		const key = fileNumberKey(toks[i + 2].rawText === '#' ? toks[i + 3] : toks[i + 2]);
		const state = key ? states.get(key) : undefined;
		if (key && state && state !== 'closed' && state !== 'empty' && state.emptyUnchecked) {
			states.set(key, { ...state, emptyUnchecked: false });
		}
	}
	if (tokenText(toks[0]) === 'seek') {
		const numberTok = toks[1]?.rawText === '#' ? toks[2] : toks[1];
		const key = numberTok ? fileNumberKey(numberTok) : undefined;
		const state = key ? states.get(key) : undefined;
		if (key && state && state !== 'closed' && state !== 'empty' && state.emptyUnchecked) {
			states.set(key, { ...state, emptyUnchecked: false });
		}
	}
}

/** Closes a number; an Output file closed with nothing written leaves its path empty. */
function closeFile(states: FileStates, key: string): void {
	const state = states.get(key);
	if (state && state !== 'closed' && state !== 'empty' && state.path !== undefined) {
		if (state.mode === 'output' && state.written === false) {
			states.set(state.path, 'empty');
		} else {
			states.delete(state.path);
		}
	}
	states.set(key, 'closed');
}

/** `Print #f, "";` writes nothing (measured: a later read raises 62). */
function writesNothing(toks: readonly VbaToken[], numberStart: number): boolean {
	const rest = toks.slice(numberStart + 2);
	return tokenText(toks[0]) === 'print' && rest.length === 2 && rest[0].rawText === '""' && rest[1].rawText === ';';
}

function forgetPathsNamedIn(states: FileStates, toks: readonly VbaToken[]): void {
	for (const tok of toks) {
		const lower = tokenName(tok)?.toLowerCase();
		if (lower) {
			forgetPath(states, `path:${lower}`);
		}
	}
}

/** A path key whose name may now hold another path: neither an empty file nor an open one is known by it. */
function forgetPath(states: FileStates, path: string): void {
	states.delete(path);
	for (const [key, state] of states) {
		if (typeof state === 'object' && state.openPath === path) {
			const rest = { ...state };
			delete rest.openPath;
			states.set(key, rest);
		}
	}
}

/**
 * `Kill p`, `FileCopy p, q` or `Name p As q` while p is open (FileCopy: open
 * in any mode but Input): Run-time error
 * 55, File already open (issue #419, measured in Excel 16.0). True when the
 * statement is one of these, which changes no file number.
 */
function checkOpenPathUse(base: Span, toks: readonly VbaToken[], states: FileStates, push: PushFn): boolean {
	const head = tokenText(toks[0]);
	if ((head !== 'kill' && head !== 'filecopy' && head !== 'name') || (head === 'name' && !toks.some((tok) => tokenText(tok) === 'as')) || ['=', '.', '('].includes(toks[1]?.rawText ?? '')) {
		return false;
	}
	const end = toks.findIndex((tok, k) => k > 0 && (tok.rawText === ',' || tokenText(tok) === 'as'));
	const pathToks = toks.slice(1, end < 0 ? toks.length : end);
	const path = pathToks.length !== 1 ? undefined
		: pathToks[0].kind === 'stringLiteral' ? `path:"${stringLiteralValue(pathToks[0].rawText)}"`
			: tokenName(pathToks[0]) !== undefined ? `path:${tokenName(pathToks[0])!.toLowerCase()}` : undefined;
	// FileCopy reads a file open For Input, and is refused one open in any other mode.
	const open = path ? [...states].find(([, state]) => typeof state === 'object' && state.openPath === path && (head !== 'filecopy' || state.mode !== 'input')) : undefined;
	if (open) {
		const word = head === 'kill' ? 'Kill' : head === 'filecopy' ? 'FileCopy' : 'Name';
		push('fileAlreadyOpen', `'${pathToks[0].rawText}' is the path of file ${describeKey(open[0])}, still open from the Open statement above; ${word} on an open file raises Run-time error '55': File already open. Close it first.`, { start: base.start + toks[0].start, end: base.start + pathToks[0].end });
	}
	if (path && head !== 'filecopy') {
		forgetPath(states, path);
	}
	return true;
}

/**
 * The literal file number at toks[index], a minus sign included, when no file
 * can have it: 0, below 0, or past 512.
 */
function impossibleFileNumber(toks: readonly VbaToken[], index: number): { value: number; first: VbaToken; last: VbaToken } | undefined {
	const literal = signedIntegerAt(toks, index);
	return literal && (literal.value < 1 || literal.value > MAX_FILE_NUMBER) ? literal : undefined;
}

/** A record or position literal below 1 at toks[index], standing alone in its slot. */
function recordBelowOne(toks: readonly VbaToken[], index: number): { value: number; first: VbaToken; last: VbaToken } | undefined {
	const literal = signedIntegerAt(toks, index);
	if (!literal || literal.value >= 1) {
		return undefined;
	}
	const after = toks[toks.indexOf(literal.last) + 1];
	return after === undefined || after.rawText === ',' ? literal : undefined;
}

function signedIntegerAt(toks: readonly VbaToken[], index: number): { value: number; first: VbaToken; last: VbaToken } | undefined {
	const negative = toks[index]?.rawText === '-';
	const tok = toks[negative ? index + 1 : index];
	if (tok?.kind !== 'integerLiteral') {
		return undefined;
	}
	const value = parseVbaIntegerLiteral(tok.rawText);
	return value === undefined ? undefined : { value: negative ? -value : value, first: toks[index], last: tok };
}

/** Where each number a Close names starts: after `#`, or alone in its slot. */
function closeNumberIndexes(toks: readonly VbaToken[]): number[] {
	if (tokenText(toks[0]) !== 'close') {
		return [];
	}
	const out: number[] = [];
	for (let i = 1; i < toks.length; i++) {
		if (toks[i].rawText === '#') {
			out.push(i + 1);
		} else if (i === 1 || toks[i - 1].rawText === ',') {
			out.push(i);
		}
	}
	return out;
}

interface ParsedOpen {
	mode: FileMode;
	key: string | undefined;
	numberIndex: number;
	/** The `path:` key of a path that is one name or one string literal. */
	path?: string;
	len?: { value: number; token: VbaToken };
}

function parseOpen(toks: readonly VbaToken[]): ParsedOpen | undefined {
	let mode: FileMode = 'random';
	let numberIndex = -1;
	let pathEnd = -1;
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
			pathEnd = pathEnd < 0 ? i : pathEnd;
			const next = tokenText(toks[i + 1]);
			if (next === 'input' || next === 'output' || next === 'append' || next === 'random' || next === 'binary') {
				mode = next;
			}
		} else if (word === 'access' || word === 'shared' || word === 'lock') {
			pathEnd = pathEnd < 0 ? i : pathEnd;
		} else if (word === 'as') {
			pathEnd = pathEnd < 0 ? i : pathEnd;
			numberIndex = toks[i + 1]?.rawText === '#' ? i + 2 : i + 1;
			break;
		}
	}
	if (numberIndex < 0 || !toks[numberIndex]) {
		return undefined;
	}
	const pathToks = toks.slice(1, pathEnd);
	const path = pathToks.length !== 1 ? undefined
		: pathToks[0].kind === 'stringLiteral' ? `path:"${stringLiteralValue(pathToks[0].rawText)}"`
			: tokenName(pathToks[0]) !== undefined ? `path:${tokenName(pathToks[0])!.toLowerCase()}` : undefined;
	// `Len = 0` after the number.
	const lenAt = toks.findIndex((tok, k) => k > numberIndex && tokenText(tok) === 'len' && toks[k + 1]?.rawText === '=');
	const lenTok = lenAt > 0 ? toks[lenAt + 2] : undefined;
	const lenValue = lenTok?.kind === 'integerLiteral' && toks[lenAt + 3] === undefined ? parseVbaIntegerLiteral(lenTok.rawText) : undefined;
	return {
		mode,
		key: fileNumberKey(toks[numberIndex]),
		numberIndex,
		...(path ? { path } : {}),
		...(lenValue !== undefined && lenTok ? { len: { value: lenValue, token: lenTok } } : {}),
	};
}

/** The key a file number token identifies: its literal value, or the variable's name. */
function fileNumberKey(tok: VbaToken | undefined): string | undefined {
	if (!tok) {
		return undefined;
	}
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
		if (opened?.path) {
			out.add(opened.path.slice('path:'.length));
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
