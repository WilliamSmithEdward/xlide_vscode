// Worksheet formula text: the form Excel shows and the form the file stores.
//
// A formula as a user types it is not the text an OOXML cell keeps. The file
// prefixes newer functions (_xlfn.), LET and LAMBDA names (_xlpm., and _xlop.
// for an optional LAMBDA parameter) and built-in functions passed by name
// (_xleta.), and spells the spill (A1#) and implicit intersection (@)
// operators as _xlfn.ANCHORARRAY() and _xlfn.SINGLE(). A shared formula is
// stored once, for its first cell, and shifted for every other one. Each of
// these forms was checked against Excel 16 storing it.

import {
	ETA_FUNCTIONS,
	FUNCTION_SIGNATURES,
	FUTURE_FUNCTIONS,
	REFERENCE_FUNCTIONS,
	RESERVED_FUNCTIONS,
} from './xlsxFunctionNames';

/** The last row and column of a worksheet: XFD1048576. */
export const MAX_ROW = 1048576;
export const MAX_COLUMN = 16384;

/** Excel's own limit on the length of a formula. */
const MAX_FORMULA_LENGTH = 8192;

/** A formula Excel would refuse; writing it would make the file unreadable. */
export class FormulaError extends Error {}

/** What a formula may refer to in its workbook. Names are upper case. */
export interface FormulaContext {
	sheets: ReadonlySet<string>;
	/** Defined names and table names, which shadow a function of the same name. */
	names: ReadonlySet<string>;
	/** Each table's columns, by table name. */
	tables: ReadonlyMap<string, ReadonlySet<string>>;
}

// ------------------------------------------------------------ A1 conversions

export function columnToIndex(letters: string): number {
	let n = 0;
	for (const ch of letters.toUpperCase()) {
		n = n * 26 + (ch.charCodeAt(0) - 64);
	}
	return n;
}

export function indexToColumn(index: number): string {
	let n = index;
	let out = '';
	while (n > 0) {
		const rem = (n - 1) % 26;
		out = String.fromCharCode(65 + rem) + out;
		n = Math.floor((n - 1) / 26);
	}
	return out;
}

// ----------------------------------------------------------------- tokens

type TokenKind =
	| 'string' | 'quoted' | 'bracket' | 'error' | 'number'
	| 'cell' | 'columns' | 'rows' | 'name' | 'space' | 'op';

interface Token {
	kind: TokenKind;
	text: string;
	/** A string, quoted name or bracket still open at the end of the text. */
	open?: boolean;
}

const CELL_RE = /^(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})/;
const COLUMNS_RE = /^(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})/;
const ROWS_RE = /^(\$?)(\d{1,7}):(\$?)(\d{1,7})/;
const NAME_RE = /^[\p{L}_\\][\p{L}\p{N}_.\\?]*/u;
const NAME_START_RE = /[\p{L}_\\]/u;
const NAME_CHAR_RE = /[\p{L}\p{N}_.\\?]/u;
const NUMBER_RE = /^(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?/;
const ERROR_RE = /^#(?:NULL!|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|N\/A|GETTING_DATA|SPILL!|CALC!|FIELD!|BLOCKED!|UNKNOWN!|CONNECT!|BUSY!|PYTHON!|EXTERNAL!)/i;

const validColumn = (letters: string): boolean => columnToIndex(letters) <= MAX_COLUMN;
const validRow = (digits: string): boolean => Number(digits) >= 1 && Number(digits) <= MAX_ROW;

/**
 * Whether the text at `at` carries a reference on into something else: a
 * longer name, a function call (LOG10), a sheet name (Q1!) or a table (T1[).
 */
function continuesPast(text: string, at: number): boolean {
	const next = text[at];
	// A1.:A9 trims a range; the dot belongs to the operator, not to a name.
	if (next === '.' && text[at + 1] === ':') { return false; }
	return next !== undefined && (NAME_CHAR_RE.test(next) || next === '(' || next === '!' || next === '[');
}

function closingQuote(text: string, open: number): number {
	const quote = text[open];
	for (let i = open + 1; i < text.length; i++) {
		if (text[i] !== quote) { continue; }
		if (text[i + 1] === quote) { i++; continue; }
		return i;
	}
	return -1;
}

/** Brackets nest in structured references, where an apostrophe escapes. */
function closingBracket(text: string, open: number): number {
	let depth = 0;
	for (let i = open; i < text.length; i++) {
		const ch = text[i];
		if (ch === "'") { i++; continue; }
		if (ch === '[') { depth++; }
		if (ch === ']' && --depth === 0) { return i; }
	}
	return -1;
}

/** Whether a token can be followed by the spill operator. */
function endsOperand(token: Token | undefined): boolean {
	return token !== undefined
		&& (token.kind === 'cell' || token.kind === 'name' || (token.kind === 'op' && token.text === ')'));
}

/** Split formula text into tokens whose texts join back to the input exactly. */
function tokenize(formula: string): Token[] {
	const tokens: Token[] = [];
	let i = 0;
	const take = (kind: TokenKind, end: number, open = false): void => {
		tokens.push(open ? { kind, text: formula.slice(i, end), open } : { kind, text: formula.slice(i, end) });
		i = end;
	};
	while (i < formula.length) {
		const ch = formula[i];
		const rest = formula.slice(i);
		if (ch === '"' || ch === "'") {
			const end = closingQuote(formula, i);
			take(ch === '"' ? 'string' : 'quoted', end < 0 ? formula.length : end + 1, end < 0);
		} else if (ch === '[') {
			const end = closingBracket(formula, i);
			take('bracket', end < 0 ? formula.length : end + 1, end < 0);
		} else if (/\s/.test(ch)) {
			take('space', i + /^\s+/.exec(rest)![0].length);
		} else if (ch === '#') {
			const error = endsOperand(tokens[tokens.length - 1]) ? null : ERROR_RE.exec(rest);
			take(error ? 'error' : 'op', i + (error ? error[0].length : 1));
		} else if (ch === '$' || /[A-Za-z]/.test(ch)) {
			const columns = COLUMNS_RE.exec(rest);
			const cell = CELL_RE.exec(rest);
			if (columns && validColumn(columns[2]) && validColumn(columns[4]) && !continuesPast(rest, columns[0].length)) {
				take('columns', i + columns[0].length);
			} else if (cell && validColumn(cell[2]) && validRow(cell[4]) && !continuesPast(rest, cell[0].length)) {
				take('cell', i + cell[0].length);
			} else if (ch !== '$') {
				take('name', i + NAME_RE.exec(rest)![0].length);
			} else {
				take('op', i + 1);
			}
		} else if ((ch === ':' && formula[i + 1] === '.') || (ch === '.' && formula[i + 1] === ':')) {
			// The trim-range operators :. .: and .:. (A1:.A9 drops trailing blanks).
			take('op', i + (ch === '.' && formula[i + 2] === '.' ? 3 : 2));
		} else if (/[0-9.]/.test(ch)) {
			const rows = ROWS_RE.exec(rest);
			const number = NUMBER_RE.exec(rest);
			if (rows && validRow(rows[2]) && validRow(rows[4]) && !continuesPast(rest, rows[0].length)) {
				take('rows', i + rows[0].length);
			} else {
				take(number ? 'number' : 'op', i + (number ? number[0].length : 1));
			}
		} else if (NAME_START_RE.test(ch)) {
			take('name', i + NAME_RE.exec(rest)![0].length);
		} else {
			take('op', i + 1);
		}
	}
	return tokens;
}

const isOp = (token: Token | undefined, text: string): boolean => token?.kind === 'op' && token.text === text;

/** Index of the token closing the parenthesis or brace at `open`, or -1. */
function closingParen(tokens: Token[], open: number): number {
	let depth = 0;
	for (let i = open; i < tokens.length; i++) {
		if (isOp(tokens[i], '(') || isOp(tokens[i], '{')) { depth++; }
		if ((isOp(tokens[i], ')') || isOp(tokens[i], '}')) && --depth === 0) { return i; }
	}
	return -1;
}

/** The argument ranges [start, end) of a call whose parentheses sit at `open` and `close`. */
function argumentRanges(tokens: Token[], open: number, close: number): Array<[number, number]> {
	const out: Array<[number, number]> = [];
	let depth = 0;
	let start = open + 1;
	for (let i = open + 1; i < close; i++) {
		if (isOp(tokens[i], '(') || isOp(tokens[i], '{')) { depth++; }
		if (isOp(tokens[i], ')') || isOp(tokens[i], '}')) { depth--; }
		if (depth === 0 && isOp(tokens[i], ',')) {
			out.push([start, i]);
			start = i + 1;
		}
	}
	out.push([start, close]);
	return out;
}

/** The one non-space token in a range, if it holds exactly one. */
function soleToken(tokens: Token[], [start, end]: [number, number]): Token | undefined {
	const solid = tokens.slice(start, end).filter((t) => t.kind !== 'space');
	return solid.length === 1 ? solid[0] : undefined;
}

// ------------------------------------------------------------ shared formulas

function shiftColumn(abs: string, letters: string, delta: number): string | undefined {
	const col = abs ? columnToIndex(letters) : columnToIndex(letters) + delta;
	return col >= 1 && col <= MAX_COLUMN ? `${abs}${indexToColumn(col)}` : undefined;
}

function shiftRow(abs: string, digits: string, delta: number): string | undefined {
	const row = abs ? Number(digits) : Number(digits) + delta;
	return row >= 1 && row <= MAX_ROW ? `${abs}${row}` : undefined;
}

function shiftToken(token: Token, rowDelta: number, colDelta: number): string {
	let parts: Array<string | undefined>;
	if (token.kind === 'cell') {
		const [, colAbs, letters, rowAbs, digits] = CELL_RE.exec(token.text)!;
		parts = [shiftColumn(colAbs, letters, colDelta), shiftRow(rowAbs, digits, rowDelta)];
		return parts.every((p) => p !== undefined) ? parts.join('') : '#REF!';
	}
	if (token.kind === 'columns') {
		const [, abs1, first, abs2, last] = COLUMNS_RE.exec(token.text)!;
		parts = [shiftColumn(abs1, first, colDelta), shiftColumn(abs2, last, colDelta)];
		return parts.every((p) => p !== undefined) ? parts.join(':') : '#REF!';
	}
	if (token.kind === 'rows') {
		const [, abs1, first, abs2, last] = ROWS_RE.exec(token.text)!;
		parts = [shiftRow(abs1, first, rowDelta), shiftRow(abs2, last, rowDelta)];
		return parts.every((p) => p !== undefined) ? parts.join(':') : '#REF!';
	}
	return token.text;
}

/**
 * Re-target a formula by a row and column offset, as a shared formula is for
 * each cell after its first: relative references move, $-anchored parts stay,
 * and function names, sheet names, strings and table columns are left alone.
 */
export function shiftFormula(formula: string, rowDelta: number, colDelta: number): string {
	if (rowDelta === 0 && colDelta === 0) { return formula; }
	return tokenize(formula).map((token) => shiftToken(token, rowDelta, colDelta)).join('');
}

// --------------------------------------------------------------- display form

const STORED_PREFIX_RE = /^(?:_xlfn\.|_xlws\.|_xlpm\.|_xleta\.)+/i;
const OPTIONAL_PARAMETER_RE = /^_xlop\./i;

/** The trim-range operators and the functions a file stores them as. */
const TRIM_OPERATORS: ReadonlyMap<string, string> = new Map([
	[':.', '_TRO_TRAILING'],
	['.:', '_TRO_LEADING'],
	['.:.', '_TRO_ALL'],
]);
const TRIM_FUNCTIONS: ReadonlyMap<string, string> = new Map([...TRIM_OPERATORS].map(([op, fn]) => [fn, op]));

/** Whether a range of tokens is a single reference: `F2`, `Sheet1!F2`, `'A b'!Name`. */
function isReference(tokens: Token[], [start, end]: [number, number]): boolean {
	const solid = tokens.slice(start, end).filter((t) => t.kind !== 'space');
	const last = solid[solid.length - 1];
	if (!last || (last.kind !== 'cell' && last.kind !== 'name')) { return false; }
	return solid.length === 1
		|| (solid.length === 3 && isOp(solid[1], '!') && (solid[0].kind === 'name' || solid[0].kind === 'quoted'));
}

/** The index of the one colon at the top level of a range of tokens, or -1. */
function soleColon(tokens: Token[], [start, end]: [number, number]): number {
	let depth = 0;
	const colons: number[] = [];
	for (let i = start; i < end; i++) {
		if (isOp(tokens[i], '(') || isOp(tokens[i], '{')) { depth++; }
		if (isOp(tokens[i], ')') || isOp(tokens[i], '}')) { depth--; }
		if (depth === 0 && isOp(tokens[i], ':')) { colons.push(i); }
	}
	return colons.length === 1 ? colons[0] : -1;
}

/** Whether a range of tokens is one function call, such as INDEX(A2:A9,3). */
function isCall(tokens: Token[], [start, end]: [number, number]): boolean {
	let first = start;
	while (first < end && tokens[first].kind === 'space') { first++; }
	let last = end - 1;
	while (last > first && tokens[last].kind === 'space') { last--; }
	return tokens[first]?.kind === 'name' && isOp(tokens[first + 1], '(') && closingParen(tokens, first + 1) === last;
}

/**
 * Whether @ can prefix a range of tokens without parentheses: a reference or
 * a range of two, one literal, a call, or an expression already in parentheses.
 */
function isPrimary(tokens: Token[], [start, end]: [number, number]): boolean {
	const solid: number[] = [];
	for (let i = start; i < end; i++) {
		if (tokens[i].kind !== 'space') { solid.push(i); }
	}
	if (solid.length === 0) { return false; }
	const first = solid[0];
	const last = solid[solid.length - 1];
	if (isReference(tokens, [start, end]) || (solid.length === 1 && tokens[first].kind !== 'op')) { return true; }
	const opensAt = tokens[first].kind === 'name' && isOp(tokens[first + 1], '(') ? first + 1 : first;
	if (isOp(tokens[opensAt], '(') && closingParen(tokens, opensAt) === last) { return true; }
	const colon = soleColon(tokens, [start, end]);
	return colon > 0 && isReference(tokens, [start, colon]) && isReference(tokens, [colon + 1, end]);
}

/** A stored structured reference's brackets as Excel shows them: [[#This Row],[Col]] as [@Col]. */
function displayTableBrackets(brackets: string): string {
	const inner = brackets.slice(1, -1).trim();
	if (/^#This Row$/i.test(inner)) { return '[@]'; }
	const items = inner.startsWith('[') ? tableItems(inner) : [];
	if (items.length < 2 || !/^\[#This Row\]$/i.test(items[0])) { return brackets; }
	const rest = items.slice(1);
	return rest.length === 1 && /^\[[\p{L}\p{N}_.]+\]$/u.test(rest[0]) ? `[@${rest[0].slice(1, -1)}]` : `[@${rest.join(',')}]`;
}

function renderDisplay(tokens: Token[], start: number, end: number): string {
	let out = '';
	for (let i = start; i < end; i++) {
		const token = tokens[i];
		if (token.kind === 'bracket' && tokens[i - 1]?.kind === 'name') {
			out += displayTableBrackets(token.text);
			continue;
		}
		if (token.kind !== 'name') {
			out += token.text;
			continue;
		}
		const bare = token.text.replace(STORED_PREFIX_RE, '');
		const upper = bare.toUpperCase();
		const operator = /^_xlfn\./i.test(token.text)
			&& (upper === 'ANCHORARRAY' || upper === 'SINGLE' || TRIM_FUNCTIONS.has(upper));
		const close = operator && isOp(tokens[i + 1], '(') ? closingParen(tokens, i + 1) : -1;
		if (close > 0) {
			const inner: [number, number] = [i + 2, close];
			const argument = renderDisplay(tokens, inner[0], inner[1]);
			if (upper === 'ANCHORARRAY') {
				out += isReference(tokens, inner) || isCall(tokens, inner) ? `${argument.trim()}#` : `ANCHORARRAY(${argument})`;
			} else if (upper === 'SINGLE') {
				out += isPrimary(tokens, inner) ? `@${argument.trim()}` : `@(${argument})`;
			} else {
				// A trimmed range: the operator takes the place of the range's colon.
				const colon = soleColon(tokens, inner);
				out += colon < 0
					? `${bare}(${argument})`
					: `${renderDisplay(tokens, inner[0], colon)}${TRIM_FUNCTIONS.get(upper)}${renderDisplay(tokens, colon + 1, inner[1])}`;
			}
			i = close;
			continue;
		}
		out += OPTIONAL_PARAMETER_RE.test(token.text) ? `[${token.text.replace(OPTIONAL_PARAMETER_RE, '')}]` : bare;
	}
	return out;
}

/**
 * A stored formula as Excel shows it: prefixes gone, ANCHORARRAY(F2) as F2#,
 * SINGLE(x) as @x and a trimmed range as A1:.A9.
 */
export function formulaForDisplay(stored: string): string {
	if (!/_xl|#This Row/i.test(stored)) { return stored; }
	const tokens = tokenize(stored);
	return renderDisplay(tokens, 0, tokens.length);
}

// ----------------------------------------------------------------- stored form
//
// A written formula is parsed as Excel parses one typed into a cell, and
// refused where Excel would refuse it: a file holding a formula Excel cannot
// parse does not open at all. The same parse produces the stored text. The
// rules were found by having Excel accept or reject each construct.

/** A run of argument counts: min, min + step, ... up to max. */
interface CountRange {
	min: number;
	max: number;
	step: number;
}

/** What Excel's parser accepts in a call to one built-in function. */
interface FunctionRule {
	counts: readonly CountRange[];
	/** Argument positions that take only a reference, as SUMIF's range does. */
	fixed: readonly number[];
	/** A repeating tail, as SUMIFS' criteria ranges: positions from `from` on, every `step`. */
	repeat?: { from: number; step: number; offsets: readonly number[] };
}

function acceptsCount(rule: FunctionRule, count: number): boolean {
	return rule.counts.some((r) => count >= r.min && count <= r.max && (count - r.min) % r.step === 0);
}

function describeArity(rule: FunctionRule): string {
	const parts = rule.counts.map((r) => {
		if (r.min === r.max) { return `${r.min}`; }
		return r.step === 1 ? `${r.min} to ${r.max}` : `${r.min} to ${r.max} in steps of ${r.step}`;
	});
	const only = rule.counts.length === 1 && rule.counts[0].min === rule.counts[0].max;
	const noun = only && rule.counts[0].min === 1 ? 'argument' : 'arguments';
	return `${only ? 'exactly ' : ''}${parts.join(', or ')} ${noun}`;
}

function takesReference(rule: FunctionRule, position: number): boolean {
	if (rule.fixed.includes(position)) { return true; }
	const repeat = rule.repeat;
	return repeat !== undefined && position >= repeat.from && repeat.offsets.includes((position - repeat.from) % repeat.step);
}

function parseRule(counts: string, refs = ''): FunctionRule {
	const [fixed, tail] = refs.split('+');
	const [from, step, offsets] = (tail ?? '').split('/');
	return {
		counts: counts.split(',').map((item) => {
			const [range, rangeStep = '1'] = item.split('/');
			const [min, max = min] = range.split('-').map(Number);
			return { min, max, step: Number(rangeStep) };
		}),
		fixed: fixed ? fixed.split(',').map(Number) : [],
		...(tail ? { repeat: { from: Number(from), step: Number(step), offsets: offsets.split(',').map(Number) } } : {}),
	};
}

const FUNCTION_RULES: ReadonlyMap<string, FunctionRule> = new Map(
	FUNCTION_SIGNATURES.split('\n').filter((line) => line.trim()).map((line) => {
		const [name, counts, refs] = line.trim().split(/\s+/);
		return [name, parseRule(counts, refs)];
	}),
);

/**
 * What an operand is, as far as the parser cares: a reference (A1, A1:B2,
 * INDEX(...)), a value (1, "a", SUM(...)), a built-in function named without
 * a call (an eta lambda), or a name or UDF result that could be either.
 */
type Kind = 'reference' | 'value' | 'function' | 'unknown';

const referenceLike = (kind: Kind): boolean => kind === 'reference' || kind === 'unknown';

interface Part {
	text: string;
	kind: Kind;
	/** The index just past the part's last token. */
	end: number;
}

const ARITHMETIC = new Set(['+', '-', '*', '/', '^', '&', '=', '<', '>']);
const LITERAL_KINDS = new Set<TokenKind>(['number', 'string', 'error']);

const isBoolean = (token: Token | undefined): boolean => token?.kind === 'name' && /^(?:TRUE|FALSE)$/i.test(token.text);

/** What Excel takes as a LET or LAMBDA name: no dots, unlike a defined name. */
const PARAMETER_NAME_RE = /^[\p{L}_][\p{L}\p{N}_]*$/u;

const TABLE_SPECIALS = /^#(?:All|Data|Headers|Totals|This Row)$/i;

/** The items in a structured reference's brackets: [#This Row],[Col] or [Col1]:[Col2]. */
function tableItems(text: string): string[] {
	const items: string[] = [];
	let depth = 0;
	let start = 0;
	for (let i = 0; i < text.length; i++) {
		const ch = text[i];
		if (ch === "'") { i++; continue; }
		if (ch === '[') { depth++; }
		if (ch === ']') { depth--; }
		if (ch === ',' && depth === 0) {
			items.push(text.slice(start, i).trim());
			start = i + 1;
		}
	}
	items.push(text.slice(start).trim());
	return items.filter((item) => item !== '');
}

/**
 * A structured reference as the file stores it, checked as Excel checks one:
 * the table and each column must exist, and Table1[@Col] is stored as
 * Table1[[#This Row],[Col]].
 */
function structuredReference(table: string, brackets: string, context: FormulaContext): string {
	const columns = context.tables.get(table.toUpperCase());
	if (!columns) {
		throw new FormulaError(`${table} is not a table in this workbook.`);
	}
	const inner = brackets.slice(1, -1).trim();
	const thisRow = inner.startsWith('@');
	const body = thisRow ? inner.slice(1).trim() : inner;
	const items = body.startsWith('[') ? tableItems(body) : body ? [`[${body}]`] : [];
	for (const item of items) {
		for (const part of item.split(/\]\s*:\s*\[/)) {
			const name = part.replace(/^\[|\]$/g, '').replace(/'(.)/g, '$1').trim();
			if (!TABLE_SPECIALS.test(name) && !columns.has(name.toUpperCase())) {
				throw new FormulaError(`table ${table} has no column named ${name}.`);
			}
		}
	}
	const all = thisRow ? ['[#This Row]', ...items] : items;
	if (all.length === 0) { return `${table}[]`; }
	return all.length === 1 && !all[0].includes(':') ? `${table}${all[0]}` : `${table}[${all.join(',')}]`;
}
const bareName = (name: string): string => name.replace(/^(?:_xlfn\.|_xlws\.)+/i, '').toUpperCase();
const alreadyStored = (name: string): boolean => /^_xl/i.test(name);
const isRefError = (token: Token | undefined): boolean => token?.kind === 'error' && /^#REF!$/i.test(token.text);

function describe(token: Token): string {
	return `'${token.kind === 'op' ? token.text : token.text.trim()}'`;
}

/** Whether a token can begin an operand, so a space before it is the intersection operator. */
function startsOperand(token: Token): boolean {
	return token.kind === 'op' ? isOp(token, '(') || isOp(token, '{') : token.kind !== 'space';
}

/** Structural faults Excel would reject, which would leave the file unreadable. */
function checkStructure(tokens: Token[]): void {
	const open = tokens.find((t) => t.open);
	if (open) {
		const what = open.kind === 'string' ? 'a string' : open.kind === 'quoted' ? 'a quoted sheet name' : 'a [ bracket';
		throw new FormulaError(`${what} is not closed.`);
	}
	const stack: string[] = [];
	for (const token of tokens) {
		if (isOp(token, '(') || isOp(token, '{')) { stack.push(token.text); }
		if (isOp(token, ')') || isOp(token, '}')) {
			const expected = token.text === ')' ? '(' : '{';
			if (stack.pop() !== expected) {
				throw new FormulaError(`the formula has a ${token.text} without its opening ${expected}.`);
			}
		}
	}
	if (stack.length > 0) {
		throw new FormulaError(`the formula has a ${stack[stack.length - 1]} that is never closed.`);
	}
	if (tokens.every((t) => t.kind === 'space')) {
		throw new FormulaError('the formula is empty.');
	}
}

/**
 * Whether the formula is one call left open at its end, such as SUM(A2,B2,
 * which Excel closes itself when it is typed. Excel closes nothing else:
 * SUM(ABS(A2 and A2+SUM(A2 are refused.
 */
function closesItself(tokens: Token[]): boolean {
	if (tokens.some((t) => t.open)) { return false; }
	const first = tokens.findIndex((t) => t.kind !== 'space');
	const opens = tokens.filter((t) => isOp(t, '(')).length;
	const closes = tokens.filter((t) => isOp(t, ')')).length;
	return opens === 1 && closes === 0 && first >= 0
		&& (tokens[first].kind === 'name' || tokens[first].kind === 'cell') && isOp(tokens[first + 1], '(');
}

function unquote(text: string): string {
	return text.slice(1, -1).replace(/''/g, "'");
}

/** Every sheet a formula names must be in this workbook; links to other workbooks are refused. */
function checkSheetReferences(tokens: Token[], context: FormulaContext): void {
	for (let i = 0; i < tokens.length; i++) {
		if (!isOp(tokens[i], '!')) { continue; }
		const before = tokens[i - 1];
		const sheets: string[] = [];
		if (before?.kind === 'quoted') {
			sheets.push(...unquote(before.text).split(':'));
		} else if (before?.kind === 'name') {
			sheets.push(before.text);
			// A 3-D reference names its first sheet too: Sheet1:Sheet3!A1.
			if (isOp(tokens[i - 2], ':') && tokens[i - 3]?.kind === 'name') {
				sheets.push(tokens[i - 3].text);
			}
		}
		const external = before?.kind === 'bracket' || tokens[i - 2]?.kind === 'bracket'
			|| sheets.some((sheet) => sheet.startsWith('['));
		if (external) {
			throw new FormulaError('it refers to another workbook; XLIDE writes formulas within this workbook only.');
		}
		const missing = sheets.find((sheet) => !context.sheets.has(sheet.toUpperCase()));
		if (missing !== undefined) {
			throw new FormulaError(`'${missing}' is not a sheet in this workbook.`);
		}
	}
}

class FormulaCompiler {
	constructor(private readonly tokens: Token[], private readonly context: FormulaContext) {}

	compile(): string {
		return this.expression(0, this.tokens.length, true, new Set()).text;
	}

	/** Tokens [start, end) as they are. */
	private text(start: number, end: number): string {
		return this.tokens.slice(start, end).map((t) => t.text).join('');
	}

	private skipSpace(i: number, end: number): number {
		while (i < end && this.tokens[i].kind === 'space') { i++; }
		return i;
	}

	private filled([start, end]: [number, number]): boolean {
		return this.skipSpace(start, end) < end;
	}

	/** Terms joined by arithmetic, concatenation and comparison. */
	private expression(start: number, end: number, unions: boolean, scope: ReadonlySet<string>): Part {
		let part = this.term(start, end, unions, scope);
		for (;;) {
			const at = this.skipSpace(part.end, end);
			if (at >= end) {
				return { text: part.text + this.text(part.end, at), kind: part.kind, end: at };
			}
			const token = this.tokens[at];
			if (token.kind === 'op' && ARITHMETIC.has(token.text)) {
				const next = this.tokens[at + 1];
				const pair = (isOp(token, '<') && (isOp(next, '=') || isOp(next, '>'))) || (isOp(token, '>') && isOp(next, '='));
				const after = at + (pair ? 2 : 1);
				const right = this.term(after, end, unions, scope);
				part = { text: part.text + this.text(part.end, after) + right.text, kind: 'value', end: right.end };
				continue;
			}
			if (isOp(token, ';')) {
				throw new FormulaError('arguments are separated by commas; a ; separates rows only in an {array} constant.');
			}
			throw new FormulaError(`${describe(token)} cannot follow a value here.`);
		}
	}

	/** An operand and the % signs after it. */
	private term(start: number, end: number, unions: boolean, scope: ReadonlySet<string>): Part {
		const operand = this.union(start, end, unions, scope);
		let text = operand.text;
		let kind = operand.kind;
		let i = operand.end;
		while (i < end && isOp(this.tokens[i], '%')) {
			text += '%';
			kind = 'value';
			i++;
		}
		return { text, kind, end: i };
	}

	/** Intersections joined by the union operator (a comma, outside an argument list). */
	private union(start: number, end: number, unions: boolean, scope: ReadonlySet<string>): Part {
		let part = this.intersection(start, end, scope);
		for (;;) {
			const at = this.skipSpace(part.end, end);
			if (!unions || at >= end || !isOp(this.tokens[at], ',')) { return part; }
			const next = this.skipSpace(at + 1, end);
			if (next >= end) {
				throw new FormulaError('the formula ends after a , where a reference is expected.');
			}
			const right = this.intersection(next, end, scope);
			part = this.join(part, right, 'a union (,)', this.text(part.end, next));
		}
	}

	/** Operands joined by the intersection operator: a space between two references. */
	private intersection(start: number, end: number, scope: ReadonlySet<string>): Part {
		let part = this.prefixed(start, end, scope);
		for (;;) {
			const at = this.skipSpace(part.end, end);
			if (at === part.end || at >= end || !startsOperand(this.tokens[at])) { return part; }
			const right = this.prefixed(at, end, scope);
			part = this.join(part, right, 'an intersection (a space between two references)', this.text(part.end, at));
		}
	}

	/**
	 * The prefix operators + - and @ on a range. @ binds tighter than a union:
	 * @A2,A4 is SINGLE(A2),A4, as Excel stores it.
	 */
	private prefixed(start: number, end: number, scope: ReadonlySet<string>): Part {
		const at = this.skipSpace(start, end);
		const lead = this.text(start, at);
		if (at >= end) {
			throw new FormulaError('the formula ends where a value is expected.');
		}
		const token = this.tokens[at];
		if (isOp(token, '+') || isOp(token, '-')) {
			const inner = this.prefixed(at + 1, end, scope);
			return { text: `${lead}${token.text}${inner.text}`, kind: 'value', end: inner.end };
		}
		if (isOp(token, '@')) {
			const inner = this.prefixed(at + 1, end, scope);
			return { text: `${lead}_xlfn.SINGLE(${inner.text})`, kind: referenceLike(inner.kind) ? 'reference' : 'value', end: inner.end };
		}
		const part = this.range(at, end, scope);
		return { ...part, text: lead + part.text };
	}

	/**
	 * Operands joined by the range operator :, or a trimmed range (:. .: .:.).
	 * Excel takes spaces around the operator and drops them when it stores it.
	 */
	private range(start: number, end: number, scope: ReadonlySet<string>): Part {
		let part = this.primary(start, end, scope);
		for (;;) {
			const at = this.skipSpace(part.end, end);
			const token = this.tokens[at];
			if (at >= end || token.kind !== 'op' || !(token.text === ':' || TRIM_OPERATORS.has(token.text))) {
				return part;
			}
			const next = this.skipSpace(at + 1, end);
			if (next >= end) {
				throw new FormulaError(`the formula ends after ${token.text}, where a reference is expected.`);
			}
			const right = this.primary(next, end, scope);
			const trim = TRIM_OPERATORS.get(token.text);
			part = trim
				? { ...this.join(part, right, 'a range', ':'), text: `_xlfn.${trim}(${part.text}:${right.text})` }
				: this.join(part, right, 'a range', ':');
		}
	}

	private join(left: Part, right: Part, what: string, between: string): Part {
		if (!referenceLike(left.kind) || !referenceLike(right.kind)) {
			throw new FormulaError(`${what} joins references, not values.`);
		}
		return { text: left.text + between + right.text, kind: 'reference', end: right.end };
	}

	/** One operand, with any spill # and argument list (calling a LAMBDA it holds) after it. */
	private primary(start: number, end: number, scope: ReadonlySet<string>): Part {
		let part = this.atom(start, end, scope);
		for (;;) {
			if (part.end >= end) { return part; }
			const token = this.tokens[part.end];
			if (isOp(token, '#')) {
				if (!referenceLike(part.kind) || this.tokens[part.end - 1].kind === 'error') {
					throw new FormulaError('the spill operator # follows a reference, such as A1#.');
				}
				part = { text: `_xlfn.ANCHORARRAY(${part.text})`, kind: 'reference', end: part.end + 1 };
			} else if (isOp(token, '(')) {
				if (part.kind === 'value') {
					throw new FormulaError('a value cannot be called like a function.');
				}
				const close = closingParen(this.tokens, part.end);
				part = { text: `${part.text}(${this.argumentsText(part.end, close, scope)})`, kind: 'unknown', end: close + 1 };
			} else {
				return part;
			}
		}
	}

	private atom(start: number, end: number, scope: ReadonlySet<string>): Part {
		const token = this.tokens[start];
		if (start >= end || !token) {
			throw new FormulaError('the formula ends where a value is expected.');
		}
		const next = this.tokens[start + 1];
		const threeD = token.kind === 'name' && isOp(next, ':') && this.tokens[start + 2]?.kind === 'name'
			&& isOp(this.tokens[start + 3], '!');
		if (threeD || ((token.kind === 'name' || token.kind === 'quoted') && isOp(next, '!'))) {
			const at = threeD ? start + 4 : start + 2;
			const target = this.tokens[at];
			const valid = at < end && target !== undefined
				&& (['cell', 'columns', 'rows', 'name'].includes(target.kind) || isRefError(target));
			if (!valid) {
				throw new FormulaError('a sheet name must be followed by a cell reference or a name, such as Sheet1!A1.');
			}
			return { text: this.text(start, at + 1), kind: 'reference', end: at + 1 };
		}
		if (token.kind === 'name' && isOp(next, '(')) {
			return this.call(start, scope);
		}
		if (token.kind === 'name' && next?.kind === 'bracket') {
			return { text: structuredReference(token.text, next.text, this.context), kind: 'reference', end: start + 2 };
		}
		switch (token.kind) {
			case 'cell':
			case 'columns':
			case 'rows':
				return { text: token.text, kind: 'reference', end: start + 1 };
			case 'error':
				return { text: token.text, kind: isRefError(token) ? 'reference' : 'value', end: start + 1 };
			case 'number':
			case 'string':
				return { text: token.text, kind: 'value', end: start + 1 };
			case 'name':
				return this.name(start, scope);
			case 'bracket':
				throw new FormulaError(`a structured reference needs its table name, such as Table1${token.text}.`);
			case 'quoted':
				throw new FormulaError(`${describe(token)} must be followed by ! and a reference.`);
			default:
				break;
		}
		if (isOp(token, '(')) {
			const close = closingParen(this.tokens, start);
			if (!this.filled([start + 1, close])) {
				throw new FormulaError('the formula has empty parentheses.');
			}
			const inner = this.expression(start + 1, close, true, scope);
			return { text: `(${inner.text})`, kind: inner.kind, end: close + 1 };
		}
		if (isOp(token, '{')) {
			const close = closingParen(this.tokens, start);
			this.arrayConstant(start + 1, close);
			return { text: this.text(start, close + 1), kind: 'value', end: close + 1 };
		}
		throw new FormulaError(`${describe(token)} is not a value.`);
	}

	private name(i: number, scope: ReadonlySet<string>): Part {
		const text = this.tokens[i].text;
		const upper = text.toUpperCase();
		if (/^_xleta\./i.test(text)) { return { text, kind: 'function', end: i + 1 }; }
		if (alreadyStored(text)) { return { text, kind: 'unknown', end: i + 1 }; }
		if (scope.has(upper)) { return { text: `_xlpm.${text}`, kind: 'unknown', end: i + 1 }; }
		if (upper === 'TRUE' || upper === 'FALSE') { return { text: upper, kind: 'value', end: i + 1 }; }
		// A built-in named without a call is a function passed as a value
		// (an eta lambda), unless the workbook has a name or table so called.
		if (ETA_FUNCTIONS.has(upper) && !this.context.names.has(upper)) {
			return { text: `_xleta.${upper}`, kind: 'function', end: i + 1 };
		}
		return { text, kind: 'unknown', end: i + 1 };
	}

	private call(start: number, scope: ReadonlySet<string>): Part {
		const token = this.tokens[start];
		const close = closingParen(this.tokens, start + 1);
		const upper = token.text.toUpperCase();
		const name = bareName(token.text);
		const args = argumentRanges(this.tokens, start + 1, close);
		const count = args.length === 1 && !this.filled(args[0]) ? 0 : args.length;
		const end = close + 1;
		if (!alreadyStored(token.text) && scope.has(upper)) {
			// A LET name holding a LAMBDA, called.
			return { text: `_xlpm.${token.text}(${this.argumentsText(start + 1, close, scope)})`, kind: 'unknown', end };
		}
		if (name === 'LET') {
			return { text: `_xlfn.LET(${this.letArguments(args, scope)})`, kind: 'unknown', end };
		}
		if (name === 'LAMBDA') {
			return { text: `_xlfn.LAMBDA(${this.lambdaArguments(args, count, scope)})`, kind: 'function', end };
		}
		// Macro-sheet commands, SPLIT and names that read as R1C1 references
		// (C1A, R2D2) are refused, where any other unknown name is a UDF.
		if (!alreadyStored(token.text) && (RESERVED_FUNCTIONS.has(name) || /^[RC](?:\d|$)/.test(name))) {
			throw new FormulaError(`${name} is not a function Excel accepts in a cell formula.`);
		}
		const rule = FUNCTION_RULES.get(name);
		if (rule && !acceptsCount(rule, count)) {
			throw new FormulaError(`${name} takes ${describeArity(rule)}, not ${count}.`);
		}
		const rendered = args.map((range, position) => {
			if (!this.filled(range)) { return this.text(range[0], range[1]); }
			const part = this.expression(range[0], range[1], false, scope);
			if (rule && takesReference(rule, position) && !referenceLike(part.kind)) {
				throw new FormulaError(`argument ${position + 1} of ${name} must be a reference, such as A1:A10, not a value.`);
			}
			return part.text;
		});
		// Excel stores a built-in's name in upper case; a UDF keeps its own.
		const stored = alreadyStored(token.text)
			? token.text
			: FUTURE_FUNCTIONS.get(name) ?? (rule || ETA_FUNCTIONS.has(name) ? name : token.text);
		const kind: Kind = REFERENCE_FUNCTIONS.has(name) ? 'reference' : rule ? 'value' : 'unknown';
		return { text: `${stored}(${rendered.join(',')})`, kind, end };
	}

	/** The arguments of a call on a LAMBDA value, such as the (2) of A1(2). */
	private argumentsText(open: number, close: number, scope: ReadonlySet<string>): string {
		return argumentRanges(this.tokens, open, close)
			.map((range) => (this.filled(range) ? this.expression(range[0], range[1], false, scope).text : this.text(range[0], range[1])))
			.join(',');
	}

	/** LET(name1, value1, ..., calculation): each name is visible after its own value. */
	private letArguments(args: Array<[number, number]>, outer: ReadonlySet<string>): string {
		if (args.length < 3 || args.length % 2 === 0) {
			throw new FormulaError('LET takes name and value pairs followed by a calculation.');
		}
		const scope = new Set(outer);
		const declared = new Set<string>();
		const out: string[] = [];
		args.forEach((range, index) => {
			if (index % 2 === 0 && index < args.length - 1) {
				const name = this.parameterName(range, 'LET');
				if (declared.has(name.toUpperCase())) {
					throw new FormulaError(`LET gives ${name} a value twice.`);
				}
				declared.add(name.toUpperCase());
				out.push(this.keepSpacing(range, `_xlpm.${name}`));
				return;
			}
			out.push(this.filled(range) ? this.expression(range[0], range[1], false, scope).text : this.text(range[0], range[1]));
			if (index % 2 === 1) {
				scope.add(this.parameterName(args[index - 1], 'LET').toUpperCase());
			}
		});
		return out.join(',');
	}

	/** LAMBDA(parameter, [optional], ..., calculation): the parameters are visible in the calculation. */
	private lambdaArguments(args: Array<[number, number]>, count: number, outer: ReadonlySet<string>): string {
		const body = args[args.length - 1];
		if (count === 0 || !this.filled(body)) {
			throw new FormulaError('LAMBDA needs a calculation as its last argument.');
		}
		const scope = new Set(outer);
		const declared = new Set<string>();
		const parameters = args.slice(0, -1).map((range) => {
			const token = soleToken(this.tokens, range);
			const optional = token?.kind === 'bracket';
			const name = optional ? token.text.slice(1, -1).trim() : this.parameterName(range, 'LAMBDA');
			if (optional && !PARAMETER_NAME_RE.test(name)) {
				throw new FormulaError('a LAMBDA parameter is letters, digits and underscores, such as rate or _total.');
			}
			if (declared.has(name.toUpperCase())) {
				throw new FormulaError(`LAMBDA names the parameter ${name} twice.`);
			}
			declared.add(name.toUpperCase());
			scope.add(name.toUpperCase());
			return this.keepSpacing(range, `${optional ? '_xlop.' : '_xlpm.'}${name}`);
		});
		return [...parameters, this.expression(body[0], body[1], false, scope).text].join(',');
	}

	private parameterName(range: [number, number], fn: string): string {
		const token = soleToken(this.tokens, range);
		const name = token?.kind === 'name' ? token.text.replace(/^_xlpm\./i, '') : '';
		if (!PARAMETER_NAME_RE.test(name) || isBoolean(token)) {
			throw new FormulaError(
				`a ${fn} name is letters, digits and underscores, such as rate or _total, and not a cell reference.`,
			);
		}
		return name;
	}

	/** A replacement for the one solid token in a range, with the range's spacing kept. */
	private keepSpacing([start, end]: [number, number], replacement: string): string {
		return this.tokens.slice(start, end).map((t) => (t.kind === 'space' ? t.text : replacement)).join('');
	}

	/** {1,2;3,4}: rows of equal length, each element a constant. */
	private arrayConstant(start: number, end: number): void {
		const widths: number[] = [];
		let width = 0;
		let expectElement = true;
		for (let i = start; i < end; i++) {
			const token = this.tokens[i];
			if (token.kind === 'space') { continue; }
			if (expectElement) {
				const signed = (isOp(token, '-') || isOp(token, '+')) && this.tokens[i + 1]?.kind === 'number';
				if (!signed && !LITERAL_KINDS.has(token.kind) && !isBoolean(token)) {
					throw new FormulaError('an {array} constant holds only numbers, text, TRUE, FALSE and error values.');
				}
				i += signed ? 1 : 0;
				width++;
				expectElement = false;
				continue;
			}
			if (isOp(token, ',') || isOp(token, ';')) {
				if (isOp(token, ';')) { widths.push(width); width = 0; }
				expectElement = true;
				continue;
			}
			throw new FormulaError(`${describe(token)} cannot appear in an {array} constant.`);
		}
		if (expectElement) {
			throw new FormulaError('an {array} constant has an empty element.');
		}
		widths.push(width);
		if (widths.some((w) => w !== widths[0])) {
			throw new FormulaError('the rows of an {array} constant must all be the same length.');
		}
	}
}

/**
 * A formula as a user types it (without the leading =), in the form the file
 * stores. Throws FormulaError for text Excel would refuse, since a file
 * holding it would not open.
 */
export function formulaForFile(formula: string, context: FormulaContext): string {
	let tokens = tokenize(formula);
	if (closesItself(tokens)) {
		tokens = tokenize(`${formula})`);
	}
	checkStructure(tokens);
	checkSheetReferences(tokens, context);
	const stored = new FormulaCompiler(tokens, context).compile();
	if (stored.length > MAX_FORMULA_LENGTH) {
		throw new FormulaError(`the formula is longer than Excel's limit of ${MAX_FORMULA_LENGTH} characters.`);
	}
	return stored;
}
