// Rule: an error value read where a number or text is needed (issue #310).
// Measured in Excel 16.0 (build 20430, 2026-10-02): a Variant holding an
// error value, from CVErr or from Excel, raises 13 as an operand of an
// arithmetic, `&` or comparison operator, after Not, as an If condition or
// a Select Case subject, in Val, Abs or Len, and Let into a typed local.
// IsError, TypeName, CStr, CLng and CInt, and a Let into another Variant,
// run.
//
// Issue #607 adds, each measured: unary minus, And, Or and Like; IIf's and
// Choose's first argument; a Do or For condition or bound; the arguments of
// Int, Fix, Sgn, Round, Sqr, CDate, Str, Format, Trim, Left, UCase, InStr,
// Mid, Hex, Chr and Space, all 13, and of CByte, 6 (2042 does not fit); an
// array's index; a ByVal typed parameter of the module's procedure; an
// element of Array(...) given to Join; and WorksheetFunction.Sum, 1004.
// CDbl, CLng, CVar, IsError and Application.Sum run.
//
// Where the value is known: `CVErr(...)` itself; a Variant local whose
// straight-line assignment is one; `Evaluate("1/0")` and `Evaluate("NA()")`;
// an element of `Array(...)` that is one; and a cell right after the
// procedure wrote `=1/0` or `=NA()` into it, through Formula or Value, as
// `Range("A1")`, `ActiveSheet.Range("A1")` or `Cells(1, 1)`. A Variant
// local given any of these holds the error from then on, though the cell
// changes after (issue #607).

import type { ConditionalActivityTracker } from '../../conditional/conditionalCompilation';
import { statementLabelDeclaration } from '../../flow/procedureLabels';
import type { VbaToken } from '../../lexer/tokenKinds';
import { splitTopLevelTokenGroups } from '../../lexer/tokenHelpers';
import type { BodyNode, ModuleNode, Span , ProcedureNode } from '../../parser/nodes';
import { isLeafStatement } from '../../parser/nodes';
import type { buildModuleSymbols } from '../../symbols/buildModuleSymbols';
import { procedureSymbolFor, type PushFn } from '../analysisContext';
import { straightLineAssignments, type ReachingAssignments } from '../straightLineValues';
import { isKnownScalarType, normalizeType, stringLiteralValue } from '../typeInference';
import {
	activeModuleMembers,
	blockHeaderLineSpan,
	matchParenFrom,
	statementAndBranchSpans,
	statementTokens,
	tokenName,
	tokenText,
} from '../walker';
import { moduleOptionBase } from './arrays';

const BINARY_OPERATORS: ReadonlySet<string> = new Set(['+', '-', '*', '/', '\\', '^', 'mod', '&', '=', '<>', '<', '>', '<=', '>=', 'and', 'or', 'xor', 'eqv', 'imp', 'like']);

const COMPARISONS: ReadonlySet<string> = new Set(['=', '<>', '<', '>', '<=', '>=']);

/** `Case 1`, `Case "a", x`, `Case 1 To 5`, `Case Is > 2`: a Case with an item of literals only. */
function literalCaseItem(all: readonly VbaToken[]): boolean {
	const toks = all.filter((tok) => tok.kind !== 'comment');
	if (tokenText(toks[0]) !== 'case' || tokenText(toks[1]) === 'else') {
		return false;
	}
	const items = toks.length > 1 ? splitTopLevelTokenGroups(toks, 1, ',', toks.length) : [];
	return items.some((item) => item.length > 0 && item.every((tok) => isLiteral(tok) || ['to', 'is'].includes(tokenText(tok)) || COMPARISONS.has(tok.rawText)));
}

/** A number or text literal. */
function isLiteral(tok: VbaToken | undefined): boolean {
	return tok?.kind === 'integerLiteral' || tok?.kind === 'floatLiteral' || tok?.kind === 'stringLiteral';
}

/** VBA functions that need a number or text of their argument, any of them (issue #607). */
const VALUE_FUNCTIONS: ReadonlySet<string> = new Set([
	'val', 'abs', 'len', 'int', 'fix', 'sgn', 'round', 'sqr', 'cdate', 'cbyte', 'str', 'format', 'trim', 'left', 'ucase', 'instr', 'mid', 'hex', 'chr', 'space',
]);

/** Functions whose first argument is read as a number or a Boolean (issue #607). */
const FIRST_ARGUMENT_FUNCTIONS: ReadonlySet<string> = new Set(['iif', 'choose']);

/** A formula whose value is an error whatever the sheet holds. */
const ERROR_FORMULA = /^\s*(?:\d+(?:\.\d+)?\s*\/\s*0|na\(\s*\))\s*$/i;

const MESSAGE_TAIL = "This will raise Run-time error '13': Type mismatch.";

export function checkErrorValues(
	source: string,
	mod: ModuleNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
	procedureFilter?: (member: ProcedureNode) => boolean,
): void {
	const optionBase = moduleOptionBase(mod, activity);
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind === 'Procedure' && procedureFilter && !procedureFilter(member)) { continue; }
		if (member.kind !== 'Procedure') {
			continue;
		}
		const children = procedureSymbolFor(symbols, member)?.children ?? [];
		// By name, so a procedure of many locals asks in constant time (issue #322).
		const locals = new Map(children.filter((child) => child.kind === 'localVariable' || child.kind === 'parameter').map((child) => [child.name.toLowerCase(), child]));
		const local = (lower: string) => locals.get(lower);
		const variantLocal = (lower: string): boolean => {
			const child = local(lower);
			const type = normalizeType(child?.asType);
			return child?.kind === 'localVariable' && child.visibility !== 'Static' && !child.isArray && (type === undefined || type === 'variant');
		};
		const scalarLocal = (lower: string): string | undefined => {
			const child = local(lower);
			const type = normalizeType(child?.asType);
			if (child && !child.isArray && type && type !== 'variant' && isKnownScalarType(type)) {
				return child.asType;
			}
			return lower === member.name.toLowerCase() && member.procKind === 'Function' && member.returnType && normalizeType(member.returnType) !== 'variant' && isKnownScalarType(normalizeType(member.returnType)!) ? member.returnType : undefined;
		};
		const arrayLocal = (lower: string): boolean => local(lower)?.kind === 'localVariable' && local(lower)!.isArray === true;
		// `TakeL(v)` with `ByVal p As Long` (issue #607).
		const byValScalarParam = (lower: string, slot: number): string | undefined => {
			const proc = (symbols.root.children ?? []).find((sym) => (sym.kind === 'function' || sym.kind === 'sub') && sym.name.toLowerCase() === lower);
			const param = proc?.children?.filter((child) => child.kind === 'parameter')[slot];
			const type = normalizeType(param?.asType);
			return param?.byVal && type && type !== 'variant' && isKnownScalarType(type) ? `${param.name} As ${param.asType}` : undefined;
		};
		const reaching = straightLineAssignments(source, member.body, activity);
		// Variant locals given an error value, by name: the text of the value
		// the walk saw reach, which must still reach where it is read
		// (issue #607). A cell's error is taken when it is read.
		const errorLocals = new Map<string, { text: string; what: string }>();
		// Cells a formula of the procedure made an error, by address, in one
		// straight run of statements.
		let errorCells = new Map<string, string>();
		const check = (span: Span, toks: readonly VbaToken[], held: ReachingAssignments | undefined, header: boolean, selectHasLiteralCase = false): void => {
			const operand = (i: number): { end: number; what: string } | undefined => errorOperand(toks, i, held, variantLocal, errorCells, optionBase)
				?? givenError(toks, i, held);
			const head = tokenText(toks[0]);
			// The `=` of an assignment is no comparison.
			const assignAt = header ? -1 : assignmentEquals(toks);
			// An operator reported through its left operand is not reported again
			// through its right one.
			let reportedOperator = -1;
			// The target of `v = ...` is written, not read.
			for (let i = assignAt + 1; i < toks.length; i++) {
				const found = operand(i);
				if (!found) {
					continue;
				}
				if (i - 1 === reportedOperator) {
					i = found.end - 1;
					continue;
				}
				const { end, what } = found;
				const at = { start: span.start + toks[i].start, end: span.start + toks[end - 1].end };
				const before = toks[i - 1];
				const after = toks[end];
				const beforeText = tokenText(before);
				const binaryBefore = BINARY_OPERATORS.has(beforeText) && i - 1 !== assignAt && i - 1 > 0 && !['(', ',', '='].includes(toks[i - 2]?.rawText ?? '') && !BINARY_OPERATORS.has(tokenText(toks[i - 2]));
				const binaryAfter = BINARY_OPERATORS.has(tokenText(after));
				let use: string | undefined;
				let error = '13';
				// Two error values compare: `v = CVErr(2042)` runs (issue #310,
				// measured), so a comparison is judged against a literal only.
				const operator = binaryAfter ? after : before;
				const other = binaryAfter ? toks[end + 1] : toks[i - 2];
				const otherWhole = binaryAfter
					? toks[end + 2] === undefined || !['(', '.', '!'].includes(toks[end + 2].rawText) && !BINARY_OPERATORS.has(tokenText(toks[end + 2]))
					: !BINARY_OPERATORS.has(tokenText(toks[i - 3])) && toks[i - 3]?.rawText !== '.';
				const otherHeld = other && tokenName(other) && variantLocal(tokenText(other)) ? held?.get(tokenText(other))?.filter((tok) => tok.kind !== 'comment') : undefined;
				const otherValue = isLiteral(other) || tokenText(other) === 'empty' || (otherHeld?.length === 1 && (isLiteral(otherHeld[0]) || tokenText(otherHeld[0]) === 'empty'));
				const comparedWithValue = !COMPARISONS.has(operator?.rawText ?? '') || (otherValue && otherWhole);
				if ((binaryBefore || binaryAfter) && comparedWithValue) {
					use = `an operand of ${operator.rawText}`;
				} else if (beforeText === 'not') {
					use = 'the operand of Not';
				} else if (before?.rawText === '-' && (i - 1 === 0 || i - 1 === assignAt + 1 || ['(', ',', '='].includes(toks[i - 2]?.rawText ?? '') || BINARY_OPERATORS.has(tokenText(toks[i - 2])))) {
					use = 'the operand of -';
				} else if (header && head === 'do' && ['while', 'until'].includes(tokenText(toks[1])) && i === 2 && end === toks.length) {
					use = 'the Do condition';
				} else if (header && head === 'loop' && ['while', 'until'].includes(tokenText(toks[1])) && i === 2 && end === toks.length) {
					use = 'the Loop condition';
				} else if (header && head === 'for' && ['to', 'step'].includes(beforeText) || (header && head === 'for' && before?.rawText === '=' && tokenText(after) === 'to')) {
					use = 'a bound of the For';
				} else if ((before?.rawText === '(' || before?.rawText === ',') && (after?.rawText === ')' || after?.rawText === ',')) {
					const call = callAround(toks, i);
					if (call) {
						const name = tokenText(toks[call.name]);
						const qualified = toks[call.name - 1]?.rawText === '.';
						if (!qualified && VALUE_FUNCTIONS.has(name)) {
							use = `an argument of ${toks[call.name].rawText}`;
							error = name === 'cbyte' ? '6' : '13';
						} else if (!qualified && FIRST_ARGUMENT_FUNCTIONS.has(name) && call.slot === 0) {
							use = `the first argument of ${toks[call.name].rawText}`;
						} else if (qualified && tokenText(toks[call.name - 2]) === 'worksheetfunction' && name === 'sum' && toks[call.name - 3]?.rawText !== '.') {
							use = 'an argument of WorksheetFunction.Sum';
							error = '1004';
						} else if (!qualified && name === 'array' && arrayGivenToJoin(toks, call.name)) {
							use = 'an element of the array Join is given';
						} else if (!qualified && arrayLocal(name)) {
							use = `an index of '${toks[call.name].rawText}'`;
						} else if (!qualified) {
							const param = byValScalarParam(name, call.slot);
							use = param ? `the argument of ${toks[call.name].rawText}'s ByVal ${param}` : undefined;
						}
					}
				} else if (header && (head === 'if' || head === 'elseif') && i === 1 && tokenText(after) === 'then') {
					use = 'the If condition';
				} else if (header && head === 'select' && tokenText(toks[1]) === 'case' && i === 2 && end === toks.length && selectHasLiteralCase) {
					use = 'the Select Case subject, compared with each Case';
				} else if (i === assignAt + 1 && end === toks.length && (assignAt === 1 || (assignAt === 2 && head === 'let'))) {
					const target = toks[assignAt - 1];
					const type = scalarLocal(tokenText(target));
					use = type ? `Let into '${target.rawText}', a ${type}` : undefined;
				}
				if (use) {
					const tail = error === '6' ? "This will raise Run-time error '6': Overflow." : error === '1004' ? "This will raise Run-time error '1004': Unable to get the Sum property of the WorksheetFunction class." : MESSAGE_TAIL;
					push('variantValueMisuse', `${what}, which is no number or text, and here it is ${use}${error === '6' ? ', whose 2042 or so does not fit a Byte' : ''}. ${tail}`, at);
					if (binaryAfter) {
						reportedOperator = end;
					}
				}
				i = end - 1;
			}
		};
		// A local the walk still sees holding what was given it in error.
		const givenError = (toks: readonly VbaToken[], i: number, held: ReachingAssignments | undefined): { end: number; what: string } | undefined => {
			const lower = toks[i - 1]?.rawText === '.' || toks[i + 1]?.rawText === '.' || toks[i + 1]?.rawText === '(' ? undefined : tokenName(toks[i])?.toLowerCase();
			const given = lower ? errorLocals.get(lower) : undefined;
			const now = lower ? held?.get(lower)?.filter((tok) => tok.kind !== 'comment').map((tok) => tok.rawText).join(' ') : undefined;
			return given && now === given.text ? { end: i + 1, what: `'${toks[i].rawText}' ${given.what}` } : undefined;
		};
		// `v = Range("A1").Value` after the error was written there, `v = a(1)`.
		const noteGiven = (toks: readonly VbaToken[], held: ReachingAssignments | undefined): void => {
			const lower = tokenName(toks[0])?.toLowerCase();
			if (!lower || toks[1]?.rawText !== '=' || !variantLocal(lower)) {
				return;
			}
			const value = toks.slice(2);
			const found = errorOperand(value, 0, held, variantLocal, errorCells, optionBase) ?? arrayLiteralElement(value, optionBase);
			if (found && found.end === value.length) {
				errorLocals.set(lower, { text: value.map((tok) => tok.rawText).join(' '), what: `holds what ${found.what.replace(/^'?([^' ]+)'? /, '$1 ')}, an error value,` });
			} else {
				errorLocals.delete(lower);
			}
		};
		const visit = (body: readonly BodyNode[]): void => {
			for (const node of body) {
				if (activity?.isInactive(node.span)) {
					continue;
				}
				if (!isLeafStatement(node)) {
					// A Dim runs nothing; a block may run anything.
					if ('body' in node && Array.isArray(node.body)) {
						errorCells = new Map();
						const header = blockHeaderLineSpan(source, node.span);
						// A Case of a number or text compares the subject with a value.
						const literalCase = node.kind === 'SelectBlock' && (node.body as BodyNode[]).some((item) => isLeafStatement(item) && literalCaseItem(statementTokens(source, item.span)));
						check(header, statementTokens(source, header), reaching.get(node), true, literalCase);
						visit(node.body as BodyNode[]);
						errorCells = new Map();
					}
					continue;
				}
				if (statementLabelDeclaration(source, node.span)) {
					errorCells = new Map();
				}
				// A single-line If is judged as its condition, up to Then, and
				// then each branch as a statement of its own.
				const branches = node.kind === 'Statement' && node.singleLineIfBranches !== undefined;
				for (const span of statementAndBranchSpans(node)) {
					const toks = statementTokens(source, span);
					if (branches && span === node.span) {
						check(span, toks.slice(0, toks.findIndex((tok) => tokenText(tok) === 'then') + 1), reaching.get(node), true);
					} else {
						check(span, toks, reaching.get(node), false);
					}
				}
				const own = statementTokens(source, node.span);
				noteGiven(own, reaching.get(node));
				errorCells = nextErrorCells(own, errorCells);
			}
		};
		visit(member.body);
	}
}

/**
 * The error value starting at `toks[i]`, with the index after it and how
 * it is shown, or undefined.
 */
function errorOperand(
	toks: readonly VbaToken[],
	i: number,
	held: ReachingAssignments | undefined,
	variantLocal: (lower: string) => boolean,
	errorCells: ReadonlyMap<string, string>,
	optionBase: number,
): { end: number; what: string } | undefined {
	if (toks[i - 1]?.rawText === '.') {
		return undefined;
	}
	const word = tokenText(toks[i]);
	const close = toks[i + 1]?.rawText === '(' ? matchParenFrom(toks, i + 1) : -1;
	if (word === 'cverr' && close > 0) {
		return { end: close + 1, what: `${toks.slice(i, close + 1).map((tok) => tok.rawText).join('')} is an error value` };
	}
	const formula = word === 'evaluate' && close === i + 3 && toks[i + 2].kind === 'stringLiteral' ? stringLiteralValue(toks[i + 2].rawText).replace(/^=/, '') : undefined;
	if (formula !== undefined && ERROR_FORMULA.test(formula)) {
		return { end: close + 1, what: `${toks.slice(i, close + 1).map((tok) => tok.rawText).join('')} gives an error value` };
	}
	// `Range("A1").Value` after the procedure wrote `=1/0` there, or
	// `ActiveSheet.Range("A1")`, `Cells(1, 1)` (issue #607).
	const cell = cellAt(toks, i);
	if (cell && errorCells.size > 0) {
		const valueRead = toks[cell.end]?.rawText === '.' && ['value', 'value2'].includes(tokenText(toks[cell.end + 1])) ? cell.end + 2 : -1;
		const read = valueRead > 0 ? valueRead : toks[cell.end]?.rawText !== '.' ? cell.end : -1;
		if (errorCells.has(cell.address) && read > 0 && toks[read]?.rawText !== '(') {
			return { end: read, what: `${toks.slice(i, read).map((tok) => tok.rawText).join('')} holds the error value of ${errorCells.get(cell.address)}` };
		}
		return undefined;
	}
	if (cell) {
		return undefined;
	}
	const lower = tokenName(toks[i])?.toLowerCase();
	if (!lower || !held || !variantLocal(lower) || toks[i + 1]?.rawText === '.') {
		return undefined;
	}
	const value = held.get(lower)?.filter((tok) => tok.kind !== 'comment');
	if (!value) {
		return undefined;
	}
	// `a(0)` with `a = Array(CVErr(2007))`.
	if (close > 0) {
		const index = close === i + 3 && toks[i + 2].kind === 'integerLiteral' ? Number(toks[i + 2].rawText) : undefined;
		if (index === undefined || tokenText(value[0]) !== 'array' || value[1]?.rawText !== '(' || matchParenFrom(value, 1) !== value.length - 1) {
			return undefined;
		}
		const elements = value.length > 3 ? splitTopLevelTokenGroups(value, 2, ',', value.length - 1) : [];
		const element = elements[index - optionBase];
		return element && isErrorSource(element) ? { end: close + 1, what: `'${toks[i].rawText}(${index})' holds an error value from ${element.map((tok) => tok.rawText).join('')} here` } : undefined;
	}
	return isErrorSource(value) ? { end: i + 1, what: `'${toks[i].rawText}' holds an error value from ${value.map((tok) => tok.rawText).join('')} here` } : undefined;
}

/**
 * The index of the `=` that makes the statement an assignment: after an
 * optional Let, a chain of names, each maybe indexed, as in `v = `,
 * `v(0) = ` or `t.x = `. -1 for any other statement.
 */
function assignmentEquals(toks: readonly VbaToken[]): number {
	let j = tokenText(toks[0]) === 'let' ? 1 : 0;
	for (;;) {
		if (!tokenName(toks[j])) {
			return -1;
		}
		j++;
		if (toks[j]?.rawText === '(') {
			const close = matchParenFrom(toks, j);
			if (close < 0) {
				return -1;
			}
			j = close + 1;
		}
		if (toks[j]?.rawText === '=') {
			return j;
		}
		if (toks[j]?.rawText !== '.') {
			return -1;
		}
		j++;
	}
}

/**
 * A cell named by literals at `toks[i]`: `Range("A1")`, `Cells(1, 1)`, each
 * maybe after `ActiveSheet.`, with its A1 address in lower case and the
 * index after it.
 */
function cellAt(toks: readonly VbaToken[], i: number): { address: string; end: number } | undefined {
	let at = i;
	if (tokenText(toks[at]) === 'activesheet' && toks[at + 1]?.rawText === '.') {
		at += 2;
	} else if (toks[at - 1]?.rawText === '.') {
		return undefined;
	}
	const word = tokenText(toks[at]);
	if (word === 'range' && toks[at + 1]?.rawText === '(' && toks[at + 2]?.kind === 'stringLiteral' && toks[at + 3]?.rawText === ')') {
		return { address: stringLiteralValue(toks[at + 2].rawText).replace(/\$/g, '').toLowerCase(), end: at + 4 };
	}
	if (word === 'cells' && toks[at + 1]?.rawText === '(' && toks[at + 2]?.kind === 'integerLiteral' && toks[at + 3]?.rawText === ','
		&& toks[at + 4]?.kind === 'integerLiteral' && toks[at + 5]?.rawText === ')') {
		const row = Number(toks[at + 2].rawText);
		let column = Number(toks[at + 4].rawText);
		let letters = '';
		while (column > 0) {
			letters = String.fromCharCode(97 + ((column - 1) % 26)) + letters;
			column = Math.floor((column - 1) / 26);
		}
		return letters && row > 0 ? { address: `${letters}${row}`, end: at + 6 } : undefined;
	}
	return undefined;
}

/** `Array(1, CVErr(2007))(1)`: an element of an array literal that is an error value. */
function arrayLiteralElement(value: readonly VbaToken[], optionBase: number): { end: number; what: string } | undefined {
	if (tokenText(value[0]) !== 'array' || value[1]?.rawText !== '(') {
		return undefined;
	}
	const close = matchParenFrom(value, 1);
	if (close < 0 || value[close + 1]?.rawText !== '(' || value[close + 2]?.kind !== 'integerLiteral' || value[close + 3]?.rawText !== ')') {
		return undefined;
	}
	const elements = close > 2 ? splitTopLevelTokenGroups(value, 2, ',', close) : [];
	const element = elements[Number(value[close + 2].rawText) - optionBase];
	return element && isErrorSource(element) ? { end: close + 4, what: `an element ${element.map((tok) => tok.rawText).join('')} of an array` } : undefined;
}

/** The call whose argument list holds `toks[i]`: its name's index and the argument's slot. */
function callAround(toks: readonly VbaToken[], i: number): { name: number; slot: number } | undefined {
	let depth = 0;
	let slot = 0;
	for (let k = i - 1; k >= 0; k--) {
		const raw = toks[k].rawText;
		if (raw === ')') {
			depth++;
		} else if (raw === '(') {
			if (depth === 0) {
				return tokenName(toks[k - 1]) ? { name: k - 1, slot } : undefined;
			}
			depth--;
		} else if (raw === ',' && depth === 0) {
			slot++;
		}
	}
	return undefined;
}

/** Whether the Array call at `at` is the first argument of Join. */
function arrayGivenToJoin(toks: readonly VbaToken[], at: number): boolean {
	return toks[at - 1]?.rawText === '(' && tokenText(toks[at - 2]) === 'join' && toks[at - 3]?.rawText !== '.';
}

/** `CVErr(...)`, or `Evaluate` of a literal formula that is an error, whole. */
function isErrorSource(value: readonly VbaToken[]): boolean {
	const word = tokenText(value[0]);
	if (value[1]?.rawText !== '(' || matchParenFrom(value, 1) !== value.length - 1) {
		return false;
	}
	if (word === 'cverr') {
		return true;
	}
	return word === 'evaluate' && value.length === 4 && value[2].kind === 'stringLiteral' && ERROR_FORMULA.test(stringLiteralValue(value[2].rawText).replace(/^=/, ''));
}

/**
 * The cells known to hold an error after a statement: `Range("A1").Formula
 * = "=1/0"` adds A1; anything else that names Range, Cells or a sheet, or
 * may run other code, ends what is known.
 */
function nextErrorCells(toks: readonly VbaToken[], cells: ReadonlyMap<string, string>): Map<string, string> {
	// `Range("A1").Formula = "=1/0"`; Value takes a formula too (issue #607).
	const cell = cellAt(toks, 0);
	if (cell && toks[cell.end]?.rawText === '.' && ['formula', 'value'].includes(tokenText(toks[cell.end + 1])) && toks[cell.end + 2]?.rawText === '='
		&& toks.length === cell.end + 4 && toks[cell.end + 3].kind === 'stringLiteral') {
		const formula = stringLiteralValue(toks[cell.end + 3].rawText);
		const next = new Map(cells);
		if (formula.startsWith('=') && ERROR_FORMULA.test(formula.slice(1)) && /^[a-z]{1,3}\d+$/.test(cell.address)) {
			next.set(cell.address, formula);
		} else {
			next.clear();
		}
		return next;
	}
	// A read leaves the cells as they are; a plain Let to a local, or into
	// the Function's result, does too.
	const plainLet = tokenName(toks[0]) !== undefined && toks[1]?.rawText === '=' && !toks.slice(2).some((tok) => ['activate', 'select', 'calculate', 'clear', 'delete', 'insert'].includes(tokenText(tok)));
	return plainLet && cells.size > 0 && !toks.slice(2).some((tok, k) => tok.kind === 'identifier' && toks[k + 3]?.rawText === '(' && !['range', 'cverr', 'val', 'abs', 'len', 'iserror', 'evaluate'].includes(tokenText(tok)))
		? new Map(cells)
		: new Map();
}
