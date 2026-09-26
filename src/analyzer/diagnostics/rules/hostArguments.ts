// Rule family: host object model arguments the code proves wrong (issue #122).
//
// Office collections are 1-based, a cell has a row and a column of at least
// 1, and a sheet name has a spelling Excel enforces. Every case here was
// measured through pyVBAharness on 2026-09-26 in Excel, Word and PowerPoint
// 16.0 (build 20326): each compiles and raises every time it runs, whatever
// the document holds.
//
//  - host-argument-out-of-range
//      Index 0 or below into a host collection: Worksheets(0), Sheets(0),
//      Workbooks(0), Names(0), Charts(0), Windows(0), ListObjects(0),
//      Comments(0), Hyperlinks(0), CommandBars(0), AddIns(0), Worksheets.Item(0)
//      -> 9 in Excel; Shapes(0) -> -2147024809; PivotTables(0) -> 1004; a
//      Range-valued collection (Rows, Columns, Areas, Cells) -> 1004.
//      Paragraphs(0), Documents(0), Tables(0), Sections(0), Words(0), ...
//      -> 5941 in Word (Shapes(0) -> -2147024809). Slides(0), Presentations(0),
//      Shapes(0), Designs(0), Slides.Add 0 -> -2147188160 in PowerPoint.
//      Cells(0, 1), Cells(1, 0), Cells(0) -> 1004. Range("A0"), Range("$A$0"),
//      Range("Sheet1!A0"), Range("0:0"), Range("A1048577"), Range("XFE1"),
//      Range("A0", "B2") -> 1004; Range("A:A") and Range("XFD1048576") run.
//      Range("A1").Offset(-1, 0), Range("A1").Offset(0, -1) -> 1004;
//      Range("B2").Offset(-1, -1) runs. Resize(0, 1), Resize(1, 0), Resize(0),
//      Resize(-1, 1) -> 1004. ActiveDocument.Range(-1, 0), Range(0, -1),
//      Range(1, 0) -> 4608 in Word; Range(0, 0) runs.
//  - sheet-name-invalid
//      Worksheets(1).Name = "a:b", "", a 32-character name, or a name holding
//      any of : \ / ? * [ ] -> 1004. A 31-character name runs.
//  - multi-cell-range-as-scalar
//      A multi-cell address literal is an array when read as a value:
//      `s = Range("A1:B2")` with s As String (or Long, Integer, Double,
//      Boolean, Date), `s = Range("A1:B2").Value`, `Range("A1:B2") = 5`,
//      `< 5`, `+ 1`, `& "x"` -> 13. A single-cell address runs.

import { parseVbaIntegerLiteral } from '../../constants/integerConstantExpression';
import type { HostObjectModel } from '../../host/excelObjectModel';
import {
	getHostMembers,
	getHostType,
	resolveHostGlobal,
	resolveHostGlobalMember,
	resolveHostMember,
} from '../../host/hostModel';
import type { MemberCompletionContext } from '../../completion/memberAccess';
import { resolveReceiverTypeAt } from '../../completion/memberAccess';
import type { VbaToken } from '../../lexer/tokenKinds';
import type { ProcedureNode, Span } from '../../parser/nodes';
import { buildModuleSymbols } from '../../symbols/buildModuleSymbols';
import { procedureSymbolFor, type PushFn } from '../analysisContext';
import { stringLiteralValue, typeEnvironmentFor, normalizeType } from '../typeInference';
import {
	bareAssignmentTarget,
	firstExecutableTokenIndex,
	matchParenFrom,
	statementTokens,
	tokenName,
	tokenText,
	type ProcedureStatementVisitor,
} from '../walker';

const EXCEL_MAX_ROW = 1048576;
const EXCEL_MAX_COLUMN = 16384;
const SHEET_NAME_MAX = 31;
const SHEET_NAME_FORBIDDEN = /[:\\/?*[\]]/;

const SCALAR_TYPES: ReadonlySet<string> = new Set([
	'string', 'long', 'integer', 'double', 'single', 'boolean', 'date', 'byte', 'currency', 'longlong',
]);

const SCALAR_OPERATORS: ReadonlySet<string> = new Set(['=', '<', '>', '<=', '>=', '<>', '+', '-', '*', '/', '\\', '&', '^']);

interface HostCallee {
	/** The member's own name as written. */
	name: string;
	/** Qualified host type the callee's value has, when the model says. */
	returns: string | undefined;
	/** Qualified host type of the receiver, or 'global' for a bare name. */
	receiver: string;
	/** Index of the name token. */
	nameIndex: number;
	/** Index of the `(` after the name, or -1 for a paren-less statement call. */
	openIndex: number;
	/** Index of the matching `)`, or the last token for a paren-less call. */
	closeIndex: number;
	/** Top-level argument groups. */
	args: VbaToken[][];
}

export function checkHostArguments(
	source: string,
	symbols: ReturnType<typeof buildModuleSymbols>,
	memberCtx: MemberCompletionContext,
	push: PushFn,
): ProcedureStatementVisitor {
	const model = memberCtx.model;
	const host = model?.hostName ?? 'Excel';
	if (host !== 'Excel' && host !== 'Word' && host !== 'PowerPoint') {
		return () => undefined;
	}
	const moduleNames = new Set<string>();
	for (const child of symbols.root.children ?? []) {
		moduleNames.add(child.name.toLowerCase());
	}
	return (proc: ProcedureNode) => {
		const env = typeEnvironmentFor(symbols, proc);
		const sourceNames = new Set(moduleNames);
		for (const child of procedureSymbolFor(symbols, proc)?.children ?? []) {
			sourceNames.add(child.name.toLowerCase());
		}
		// A single-line If is one span here: its branches would otherwise be
		// walked twice, once inside the whole statement and once on their own.
		return (stmt) => checkSpan(source, stmt.span, host, model, memberCtx, env, sourceNames, push);
	};
}

function checkSpan(
	source: string,
	span: Span,
	host: 'Excel' | 'Word' | 'PowerPoint',
	model: HostObjectModel | undefined,
	memberCtx: MemberCompletionContext,
	env: ReadonlyMap<string, string>,
	sourceNames: ReadonlySet<string>,
	push: PushFn,
): void {
	const toks = statementTokens(source, span);
	const at = (from: number, to: number): Span => ({ start: span.start + toks[from].start, end: span.start + toks[to].end });
	if (host === 'Excel') {
		checkSheetNameAssignment(source, span, toks, memberCtx, push);
	}
	for (let i = 0; i < toks.length; i++) {
		const callee = hostCalleeAt(source, span, toks, i, model, memberCtx, sourceNames);
		if (!callee) {
			continue;
		}
		const calleeSpan = at(callee.nameIndex, callee.closeIndex);
		const lower = callee.name.toLowerCase();
		// Index 0 into a 1-based collection, `Worksheets(0)` or `Worksheets.Item(0)`.
		const collection = lower === 'item' && isCollectionType(callee.receiver, model)
			? callee.receiver
			: callee.returns && isCollectionType(callee.returns, model) && callee.openIndex > 0 ? callee.returns : undefined;
		if (collection && callee.args.length === 1 && lower !== 'cells' && lower !== 'range') {
			const index = integerLiteralValue(callee.args[0]);
			if (index !== undefined && index < 1) {
				const error = collectionIndexError(host, collection, model);
				push(
					'hostArgumentOutOfRange',
					`Index ${index} is never an element: ${host} collections start at 1. This will raise Run-time error '${error.number}': ${error.text}.`,
					at(callee.openIndex + 1, callee.closeIndex - 1),
				);
				continue;
			}
		}
		if (host === 'Excel') {
			checkExcelCallee(source, span, toks, callee, calleeSpan, env, push);
		} else if (host === 'Word') {
			if (lower === 'range' && callee.receiver === 'Word.Document' && callee.openIndex > 0) {
				const start = callee.args[0] ? integerLiteralValue(callee.args[0]) : undefined;
				const end = callee.args[1] ? integerLiteralValue(callee.args[1]) : undefined;
				const bad = (start !== undefined && start < 0) || (end !== undefined && end < 0) || (start !== undefined && end !== undefined && end < start);
				if (bad) {
					push(
						'hostArgumentOutOfRange',
						`Document.Range takes character positions from 0, with End at or after Start. This will raise Run-time error '4608': Value out of range.`,
						at(callee.openIndex + 1, callee.closeIndex - 1),
					);
				}
			}
		} else if (lower === 'add' && callee.receiver === 'PowerPoint.Slides' && callee.args.length >= 1) {
			const index = integerLiteralValue(callee.args[0]);
			if (index !== undefined && index < 1) {
				push(
					'hostArgumentOutOfRange',
					`Slides.Add places the new slide at Index, which starts at 1. This will raise Run-time error '-2147188160': Integer out of range.`,
					argSpan(span, callee.args[0]),
				);
			}
		}
	}
}

/**
 * The host member a name-with-arguments at `toks[i]` calls: a bare host
 * global (`Worksheets(0)`, `Cells(0, 1)`), a member of the hidden Global
 * interface (`Rows(0)`, `Names(0)`), or a member of a resolved receiver
 * (`ActiveDocument.Paragraphs(0)`, `Range("A1").Offset(-1, 0)`). A name the
 * module or procedure declares is the source's, not the host's.
 */
function hostCalleeAt(
	source: string,
	span: Span,
	toks: readonly VbaToken[],
	i: number,
	model: HostObjectModel | undefined,
	memberCtx: MemberCompletionContext,
	sourceNames: ReadonlySet<string>,
): HostCallee | undefined {
	const name = tokenName(toks[i]);
	if (!name) {
		return undefined;
	}
	const qualified = toks[i - 1]?.rawText === '.';
	const parenthesized = toks[i + 1]?.rawText === '(';
	let openIndex = -1;
	let closeIndex: number;
	let args: VbaToken[][];
	if (parenthesized) {
		openIndex = i + 1;
		closeIndex = matchParenFrom(toks, openIndex);
		if (closeIndex < 0) {
			return undefined;
		}
		args = splitTopLevel(toks.slice(openIndex + 1, closeIndex));
	} else if (qualified && i === firstExecutableTokenIndexOfMemberCall(toks, i) && toks[i + 1] && toks[i + 1].rawText !== '.' && toks[i + 1].rawText !== '=') {
		// Statement form: `ActivePresentation.Slides.Add 0, ppLayoutBlank`.
		closeIndex = toks.length - 1;
		args = splitTopLevel(toks.slice(i + 1));
	} else {
		return undefined;
	}
	if (!qualified) {
		if (sourceNames.has(name.toLowerCase()) || !parenthesized) {
			return undefined;
		}
		const globalType = resolveHostGlobal(name, model);
		if (globalType) {
			return { name, returns: globalType, receiver: 'global', nameIndex: i, openIndex, closeIndex, args };
		}
		const member = resolveHostGlobalMember(name, model);
		if (member) {
			return { name, returns: member.returns, receiver: 'global', nameIndex: i, openIndex, closeIndex, args };
		}
		return undefined;
	}
	const receiver = hostReceiverType(resolveReceiverTypeAt(source, span.start + toks[i - 1].end, memberCtx), model);
	if (!receiver) {
		return undefined;
	}
	const member = resolveHostMember(receiver, name, model);
	if (!member) {
		return undefined;
	}
	return { name, returns: member.returns, receiver, nameIndex: i, openIndex, closeIndex, args };
}

/**
 * The host type a resolved receiver names. A one-part union - what a
 * collection's Object-declared Item gives, `Worksheets(1)` (issue #114) - is
 * its part; a union of several types is not judged.
 */
function hostReceiverType(resolved: string | undefined, model: HostObjectModel | undefined): string | undefined {
	if (!resolved) {
		return undefined;
	}
	const parts = resolved.startsWith('union:') ? resolved.slice('union:'.length).split('|') : [resolved];
	if (parts.length !== 1 || !getHostType(parts[0], model)) {
		return undefined;
	}
	return parts[0];
}

/** The member call's name index when the statement is `a.b.Name args`, else -1. */
function firstExecutableTokenIndexOfMemberCall(toks: readonly VbaToken[], nameIndex: number): number {
	let j = nameIndex;
	while (j >= 2 && toks[j - 1].rawText === '.' && tokenName(toks[j - 2])) {
		j -= 2;
		if (toks[j - 1]?.rawText === ')') {
			const open = toks.findIndex((tok, k) => tok.rawText === '(' && matchParenFrom(toks, k) === j - 1);
			if (open < 1) {
				return -1;
			}
			j = open;
		}
	}
	return j === firstExecutableTokenIndex(toks) ? nameIndex : -1;
}

function isCollectionType(type: string, model: HostObjectModel | undefined): boolean {
	const members = getHostMembers(type, model);
	return members.some((m) => m.name === 'Item') && members.some((m) => m.name === 'Count');
}

function collectionIndexError(
	host: 'Excel' | 'Word' | 'PowerPoint',
	collection: string,
	model: HostObjectModel | undefined,
): { number: string; text: string } {
	const bare = collection.slice(collection.indexOf('.') + 1);
	if (host === 'PowerPoint') {
		return { number: '-2147188160', text: 'Integer out of range' };
	}
	if (bare === 'Shapes') {
		return { number: '-2147024809', text: 'The index into the specified collection is out of bounds' };
	}
	if (host === 'Word') {
		return { number: '5941', text: 'The requested member of the collection does not exist' };
	}
	if (bare === 'PivotTables' || bare === 'Range' || resolveHostMember(collection, 'Item', model)?.returns === 'Excel.Range') {
		return { number: '1004', text: 'Application-defined or object-defined error' };
	}
	return { number: '9', text: 'Subscript out of range' };
}

function checkExcelCallee(
	source: string,
	span: Span,
	toks: readonly VbaToken[],
	callee: HostCallee,
	calleeSpan: Span,
	env: ReadonlyMap<string, string>,
	push: PushFn,
): void {
	const lower = callee.name.toLowerCase();
	if (callee.openIndex < 0) {
		return;
	}
	const argsSpan = callee.closeIndex > callee.openIndex + 1
		? { start: span.start + toks[callee.openIndex + 1].start, end: span.start + toks[callee.closeIndex - 1].end }
		: calleeSpan;
	if (lower === 'cells' && callee.returns === 'Excel.Range') {
		for (const arg of callee.args) {
			const value = integerLiteralValue(arg);
			if (value !== undefined && value < 1) {
				push('hostArgumentOutOfRange', `Cells takes a row and a column of at least 1; ${value} names no cell. This will raise Run-time error '1004': Application-defined or object-defined error.`, argSpan(span, arg));
				return;
			}
		}
		return;
	}
	if (lower === 'resize' && callee.receiver === 'Excel.Range') {
		for (const arg of callee.args) {
			const value = integerLiteralValue(arg);
			if (value !== undefined && value < 1) {
				push('hostArgumentOutOfRange', `Resize needs at least one row and one column; ${value} gives none. This will raise Run-time error '1004': Application-defined or object-defined error.`, argSpan(span, arg));
				return;
			}
		}
		return;
	}
	if (lower === 'offset' && callee.receiver === 'Excel.Range') {
		const origin = singleCellReceiver(toks, callee.nameIndex - 1);
		if (!origin) {
			return;
		}
		const rowOffset = callee.args[0] ? integerLiteralValue(callee.args[0]) : 0;
		const columnOffset = callee.args[1] ? integerLiteralValue(callee.args[1]) : 0;
		if (rowOffset === undefined || columnOffset === undefined) {
			return;
		}
		const row = origin.row + rowOffset;
		const column = origin.column + columnOffset;
		if (row < 1 || column < 1) {
			push('hostArgumentOutOfRange', `Offset(${rowOffset}, ${columnOffset}) from ${origin.text} lands at row ${row}, column ${column}, off the sheet. This will raise Run-time error '1004': Application-defined or object-defined error.`, argsSpan);
		}
		return;
	}
	if (lower === 'range' && callee.returns === 'Excel.Range') {
		const areas = callee.args.map((arg) => (arg.length === 1 && arg[0].kind === 'stringLiteral' ? parseA1Address(stringLiteralValue(arg[0].rawText)) : undefined));
		for (let k = 0; k < callee.args.length; k++) {
			const area = areas[k];
			if (area && !area.valid) {
				push('hostArgumentOutOfRange', `"${area.text}" is not a cell address Excel accepts: rows run 1 to ${EXCEL_MAX_ROW} and columns A to XFD. This will raise Run-time error '1004': Method 'Range' of object failed.`, argSpan(span, callee.args[k]));
				return;
			}
		}
		if (callee.args.length === 1 && areas[0]?.valid && areas[0].multiCell) {
			checkMultiCellAsScalar(source, span, toks, callee, areas[0].text, env, push);
		}
	}
}

/**
 * `Range("A1:B2")` read as a value is a two-dimensional array (issue #122):
 * assigned to a scalar variable, or combined with a scalar operator, it is a
 * type mismatch. `.Value` after it changes nothing.
 */
function checkMultiCellAsScalar(
	source: string,
	span: Span,
	toks: readonly VbaToken[],
	callee: HostCallee,
	address: string,
	env: ReadonlyMap<string, string>,
	push: PushFn,
): void {
	let end = callee.closeIndex;
	if (toks[end + 1]?.rawText === '.' && ['value', 'value2'].includes(tokenText(toks[end + 2]))) {
		end += 2;
	}
	if (toks[end + 1]?.rawText === '.' || toks[end + 1]?.rawText === '(') {
		return; // a member or an index: not the range read as a value
	}
	const start = callee.receiver === 'global' ? callee.nameIndex : receiverStart(toks, callee.nameIndex);
	const valueSpan = { start: span.start + toks[start].start, end: span.start + toks[end].end };
	const message = (use: string): string =>
		`Range("${address}") read as a value is a two-dimensional array, ${use}. This will raise Run-time error '13': Type mismatch.`;
	const bare = bareAssignmentTarget(source, span);
	if (bare) {
		const eq = toks.findIndex((tok) => tok.rawText === '=');
		if (eq === start - 1 && end === toks.length - 1) {
			const target = normalizeType(env.get(bare.name.toLowerCase()));
			if (target && SCALAR_TYPES.has(target)) {
				push('multiCellRangeAsScalar', message(`which a ${env.get(bare.name.toLowerCase())} variable cannot hold`), valueSpan);
			}
			return;
		}
		if (eq < 0 || start <= eq) {
			return;
		}
	} else {
		const head = tokenText(toks[firstExecutableTokenIndex(toks)]);
		if (head !== 'if' && head !== 'elseif' && head !== 'while' && head !== 'until' && head !== 'do' && head !== 'loop' && head !== 'select' && head !== 'case') {
			return;
		}
	}
	// The operator on either side, never the assignment's own `=`.
	const eqIndex = bare ? toks.findIndex((tok) => tok.rawText === '=') : -1;
	const before = start - 1 === eqIndex ? undefined : toks[start - 1];
	const after = toks[end + 1];
	const operator = [after, before].find((tok) => tok && ((tok.kind === 'operator' && SCALAR_OPERATORS.has(tok.rawText)) || tokenText(tok) === 'mod'));
	if (operator) {
		push('multiCellRangeAsScalar', message(`which '${operator.rawText}' cannot combine with a scalar`), valueSpan);
	}
}

/** `Worksheets(1).Name = "a:b"`: the receiver's type and the literal decide. */
function checkSheetNameAssignment(
	source: string,
	span: Span,
	toks: readonly VbaToken[],
	memberCtx: MemberCompletionContext,
	push: PushFn,
): void {
	const n = toks.length;
	if (n < 4 || toks[n - 1].kind !== 'stringLiteral' || toks[n - 2].rawText !== '=' || tokenText(toks[n - 3]) !== 'name' || toks[n - 4].rawText !== '.') {
		return;
	}
	const resolved = resolveReceiverTypeAt(source, span.start + toks[n - 4].end, memberCtx);
	// `Worksheets(1)` is a one-part union, `Sheets(1)` a Worksheet-or-Chart union: both are sheets.
	const parts = resolved ? (resolved.startsWith('union:') ? resolved.slice('union:'.length).split('|') : [resolved]) : [];
	if (parts.length === 0 || !parts.every((part) => part === 'Excel.Worksheet' || part === 'Excel.Chart')) {
		return;
	}
	const name = stringLiteralValue(toks[n - 1].rawText);
	let problem: string | undefined;
	if (name.length === 0) {
		problem = 'a sheet name cannot be blank';
	} else if (name.length > SHEET_NAME_MAX) {
		problem = `a sheet name has at most ${SHEET_NAME_MAX} characters, and this one has ${name.length}`;
	} else if (SHEET_NAME_FORBIDDEN.test(name)) {
		problem = 'a sheet name cannot contain any of : \\ / ? * [ ]';
	}
	if (problem) {
		push('sheetNameInvalid', `Excel refuses this name: ${problem}. This will raise Run-time error '1004': You typed an invalid name for a sheet or chart.`, { start: span.start + toks[n - 1].start, end: span.start + toks[n - 1].end });
	}
}

/** The single-cell literal `Range("B2")` ending at `toks[closeIndex]`, when that is the receiver. */
function singleCellReceiver(toks: readonly VbaToken[], dotIndex: number): { row: number; column: number; text: string } | undefined {
	if (toks[dotIndex]?.rawText !== '.' || toks[dotIndex - 1]?.rawText !== ')') {
		return undefined;
	}
	const close = dotIndex - 1;
	const open = toks.findIndex((tok, k) => tok.rawText === '(' && matchParenFrom(toks, k) === close);
	if (open < 1 || tokenText(toks[open - 1]) !== 'range' || close !== open + 2 || toks[open + 1].kind !== 'stringLiteral') {
		return undefined;
	}
	const area = parseA1Address(stringLiteralValue(toks[open + 1].rawText));
	if (!area?.valid || area.multiCell || area.row === undefined || area.column === undefined) {
		return undefined;
	}
	return { row: area.row, column: area.column, text: `Range("${area.text}")` };
}

/** Index of the first token of the receiver chain that ends at the dot before `nameIndex`. */
function receiverStart(toks: readonly VbaToken[], nameIndex: number): number {
	let j = nameIndex;
	while (j >= 2 && toks[j - 1].rawText === '.') {
		j -= 1;
		if (toks[j - 1]?.rawText === ')') {
			const open = toks.findIndex((tok, k) => tok.rawText === '(' && matchParenFrom(toks, k) === j - 1);
			j = open >= 1 ? open : j;
		}
		j -= 1;
	}
	return j;
}

interface A1Area {
	text: string;
	valid: boolean;
	multiCell: boolean;
	row?: number;
	column?: number;
}

/**
 * Parses an A1-style address literal: `A1`, `$A$1`, `A1:B2`, `A:A`, `1:1`,
 * with an optional `Sheet1!` or `'My Sheet'!` prefix. Anything else (a name,
 * an R1C1 address, a union) is not judged.
 */
function parseA1Address(text: string): A1Area | undefined {
	const body = text.replace(/^(?:'[^']*'|[^!'\s]+)!/, '');
	const cell = /^\$?([A-Za-z]{1,3})\$?(\d+)$/;
	const parts = body.split(':');
	if (parts.length > 2) {
		return undefined;
	}
	const cells = parts.map((part) => cell.exec(part));
	if (cells.every((match) => match)) {
		const rows = cells.map((match) => Number(match![2]));
		const columns = cells.map((match) => columnNumber(match![1]));
		const valid = rows.every((row) => row >= 1 && row <= EXCEL_MAX_ROW) && columns.every((column) => column >= 1 && column <= EXCEL_MAX_COLUMN);
		const multiCell = cells.length === 2 && (rows[0] !== rows[1] || columns[0] !== columns[1]);
		return { text, valid, multiCell, row: rows[0], column: columns[0] };
	}
	if (parts.length === 2) {
		const columnsOnly = parts.map((part) => /^\$?([A-Za-z]{1,3})$/.exec(part));
		if (columnsOnly.every((match) => match)) {
			const valid = columnsOnly.every((match) => columnNumber(match![1]) <= EXCEL_MAX_COLUMN);
			return { text, valid, multiCell: true };
		}
		const rowsOnly = parts.map((part) => /^\$?(\d+)$/.exec(part));
		if (rowsOnly.every((match) => match)) {
			const valid = rowsOnly.every((match) => Number(match![1]) >= 1 && Number(match![1]) <= EXCEL_MAX_ROW);
			return { text, valid, multiCell: true };
		}
	}
	return undefined;
}

function columnNumber(letters: string): number {
	let n = 0;
	for (const ch of letters.toUpperCase()) {
		n = n * 26 + (ch.charCodeAt(0) - 64);
	}
	return n;
}

/** The whole-number value of an argument that is a literal, optionally negated. */
function integerLiteralValue(arg: readonly VbaToken[]): number | undefined {
	const toks = arg.filter((tok) => tok.kind !== 'comment');
	if (toks.length === 1 && toks[0].kind === 'integerLiteral') {
		return parseVbaIntegerLiteral(toks[0].rawText);
	}
	if (toks.length === 2 && toks[0].rawText === '-' && toks[1].kind === 'integerLiteral') {
		const value = parseVbaIntegerLiteral(toks[1].rawText);
		return value === undefined ? undefined : -value;
	}
	return undefined;
}

function splitTopLevel(toks: readonly VbaToken[]): VbaToken[][] {
	const out: VbaToken[][] = [];
	let current: VbaToken[] = [];
	let depth = 0;
	for (const tok of toks) {
		if (tok.rawText === '(') {
			depth++;
		} else if (tok.rawText === ')') {
			depth--;
		}
		if (tok.rawText === ',' && depth === 0) {
			out.push(current);
			current = [];
			continue;
		}
		current.push(tok);
	}
	if (current.length > 0 || out.length > 0) {
		out.push(current);
	}
	return out;
}

function argSpan(span: Span, arg: readonly VbaToken[]): Span {
	return { start: span.start + arg[0].start, end: span.start + arg[arg.length - 1].end };
}
