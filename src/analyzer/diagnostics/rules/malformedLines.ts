// Rule family: lines the VBE cannot parse, as typing leaves them (issue #234).
// Measured in Excel 16.0 (build 20326, 2026-09-30) with a full compile.
//
//  - malformed-statement:
//      `#asdf`, a directive that is not #If, #ElseIf, #Else, #End If or
//      #Const -> "Syntax error" in a procedure and after one, "Expected: If
//      or Else or ElseIf or End or EndIf or Const" above the first.
//      `[asdf`, a bracket never closed -> "Syntax error", "Missing end
//      bracket" in an Enum.
//      `Sub` or `Function` with no name -> "Expected: identifier" above the
//      first procedure, "Syntax error" after one.
//      `Private Sub S(a b c)`, `S(x Property)`: a parameter followed by
//      another word -> "Expected: list separator or )".
//      An Enum line that is not `name [= value]`: `asdf qwer`, `.asdf` ->
//      "Invalid inside Enum"; `asdf & qwer` -> "Expected: expression".
//  - reserved-keyword-in-expression: a statement keyword where a value
//      goes: `Array(1, Const, 3)`, `Debug.Print RaiseEvent`, `Case Open`,
//      `Dim a(Implements) As Long` -> "Syntax error"; `Const K = Dim` ->
//      "Expected: expression".
//  - statement-outside-procedure: a lone `:` after a procedure -> "Only
//      comments may appear after End Sub, End Function, or End Property". It
//      compiles above the first procedure.
//
// None of it is judged in an inactive #If branch, which the VBE does not
// parse.

import type { ConditionalActivityTracker } from '../../conditional/conditionalCompilation';
import { tokenizeCached } from '../../lexer/tokenize';
import { firstTokenAtOrAfter } from '../../lexer/tokenHelpers';
import type { VbaToken } from '../../lexer/tokenKinds';
import type { EnumNode, ModuleNode, ProcedureNode, Span, VariableGroupNode } from '../../parser/nodes';
import type { PushFn } from '../analysisContext';
import {
	activeModuleMembers,
	bareAssignmentTarget,
	forEachStatement,
	forEachVariableGroup,
	matchParenFrom,
	statementTokens,
	tokenText,
} from '../walker';

const DIRECTIVE_WORDS: ReadonlySet<string> = new Set(['if', 'elseif', 'else', 'end', 'endif', 'const']);

/**
 * Statement keywords that can never stand where a value goes. `To`, `Is`,
 * `Else`, `As` and `New` are left out: each has a place inside a value's
 * statement (`Case 1 To 3`, `Case Is > 1`, `Case Else`, `Set x = New C`).
 */
const STATEMENT_KEYWORDS: ReadonlySet<string> = new Set([
	'const', 'raiseevent', 'open', 'close', 'implements', 'dim', 'redim', 'static', 'sub', 'function',
	'declare', 'enum', 'event', 'option', 'call', 'goto', 'gosub', 'exit', 'resume', 'do',
	'loop', 'wend', 'while', 'until', 'with', 'select', 'case', 'next', 'for', 'each', 'then', 'elseif',
	'if', 'end', 'public', 'private', 'friend', 'global', 'let', 'set', 'stop', 'return', 'lock',
	'unlock', 'preserve', 'withevents', 'type', 'put', 'get', 'kill', 'step',
]);

const PARAMETER_MODIFIERS: ReadonlySet<string> = new Set(['optional', 'byval', 'byref', 'paramarray']);

type Place = 'procedure' | 'top' | 'after';

export function checkMalformedLines(
	source: string,
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	const procedures = mod.members.filter((m): m is ProcedureNode => m.kind === 'Procedure');
	const namedProcedures = procedures.filter((p) => p.name !== '');
	const firstNamed = namedProcedures[0];
	const place = (offset: number): Place => {
		if (insideMember(namedProcedures, offset)) {
			return 'procedure';
		}
		return !firstNamed || offset < firstNamed.span.start ? 'top' : 'after';
	};
	const enums = mod.members.filter((m): m is EnumNode => m.kind === 'Enum');
	const inEnum = (offset: number): boolean => insideMember(enums, offset);
	const active = (span: Span): boolean => !activity?.isInactive(span);
	const toks = tokenizeCached(source);

	for (let i = 0; i < toks.length; i++) {
		const tok = toks[i];
		const span = { start: tok.start, end: tok.end };
		if (tok.kind === 'directive' && active(span)) {
			const word = toks[i + 1] && toks[i + 1].kind !== 'newline' ? toks[i + 1] : undefined;
			if (!word || !DIRECTIVE_WORDS.has(tokenText(word))) {
				const error = place(tok.start) === 'top' ? 'Expected: If or Else or ElseIf or End or EndIf or Const' : 'Syntax error';
				const shown = word ? `#${word.rawText}` : '#';
				push('malformedStatement', `'${shown}' is no compiler directive: only #If, #ElseIf, #Else, #End If and #Const are. This is a VBE compile error: ${error}.`, word ? { start: tok.start, end: word.end } : span);
			}
		} else if (tok.kind === 'bracketedIdentifier' && !tok.rawText.endsWith(']') && active(span)) {
			const error = inEnum(tok.start) ? 'Missing end bracket' : 'Syntax error';
			push('malformedStatement', `'${tok.rawText}' opens a bracketed name that never closes. This is a VBE compile error: ${error}.`, span);
		} else if (tok.kind === 'colon' && lineHoldsOnlyColons(toks, i) && place(tok.start) === 'after' && active(span)) {
			push('statementOutsideProcedure', "A ':' after a procedure separates nothing. This is a VBE compile error: Only comments may appear after End Sub, End Function, or End Property.", span);
		}
	}

	checkStatementForms(toks, (span) => active(span) && place(span.start) === 'procedure', push);

	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind === 'Procedure') {
			checkProcedureHeader(toks, member, place, push);
			forEachStatement(member.body, (stmt) => {
				checkValueKeywords(source, stmt.span, 'statement', push);
				checkKeywordQualifiers(source, stmt.span, push);
				checkValueWords(source, stmt.span, push);
			}, activity);
			forEachVariableGroup(member.body, (group) => {
				checkDeclarationKeywords(source, group, push);
			}, activity);
		} else if (member.kind === 'VariableGroup') {
			checkDeclarationKeywords(source, member, push);
		} else if (member.kind === 'Enum') {
			checkEnumLines(source, member, activity, push);
		}
	}
}

/** Members are source-ordered and non-overlapping; their boundary tokens are outside. */
function insideMember(members: readonly { span: Span }[], offset: number): boolean {
	let lo = 0;
	let hi = members.length;
	while (lo < hi) {
		const mid = lo + Math.floor((hi - lo) / 2);
		if (members[mid].span.start < offset) {
			lo = mid + 1;
		} else {
			hi = mid;
		}
	}
	return lo > 0 && offset < members[lo - 1].span.end;
}

/**
 * The operators a VB.NET-style compound assignment puts before `=`: `x += 1`.
 * The other operators are an operator run, invalid-expression-syntax's.
 */
const COMPOUND_ASSIGNMENT_OPERATORS: ReadonlySet<string> = new Set(['+', '-']);

/**
 * Statement forms the VBE refuses (issue #369, measured in Excel 16.0):
 * `Else If b Then` on its own line, `x += 1`, `Exit While`,
 * `Call Debug.Print(...)` and any Call of Debug, a Case after Case Else, a
 * Loop with a condition after a Do with one, a Loop with both While and
 * Until, and `Case Is` with To or Like. Each statement of each line is read
 * in order; `inProcedure` says whether a span is active code inside a
 * procedure.
 */
function checkStatementForms(
	toks: readonly VbaToken[],
	inProcedure: (span: Span) => boolean,
	push: PushFn,
): void {
	const selects: Array<{ sawElse: boolean }> = [];
	const dos: boolean[] = [];
	const at = (first: VbaToken, last: VbaToken): Span => ({ start: first.start, end: last.end });
	const report = (first: VbaToken, last: VbaToken, what: string, error: string): void => {
		push('malformedStatement', `${what} This is a VBE compile error: ${error}.`, at(first, last));
	};
	let line: VbaToken[] = [];
	const flushLine = (): void => {
		const words = line.filter((tok) => tok.kind !== 'comment');
		line = [];
		if (words.length === 0 || !inProcedure(at(words[0], words[words.length - 1]))) {
			return;
		}
		// `Else If b Then` with nothing after Then: Else and a block If header.
		if (tokenText(words[0]) === 'else' && tokenText(words[1]) === 'if' && tokenText(words[words.length - 1]) === 'then') {
			report(words[0], words[words.length - 1], "'Else If ... Then' starts a block If on the Else line; write ElseIf, one word.", 'Syntax error');
		}
		for (const statement of splitStatements(words)) {
			checkStatement(statement);
		}
	};
	const checkStatement = (words: readonly VbaToken[]): void => {
		const head = tokenText(words[0]);
		for (let k = 0; k + 1 < words.length; k++) {
			if (tokenText(words[k]) === 'exit' && tokenText(words[k + 1]) === 'while') {
				report(words[k], words[k + 1], "'Exit While' is no statement: a While loop has no exit of its own, so use Do While ... Loop and Exit Do.", 'Syntax error');
			}
			if (tokenText(words[k]) === 'call' && tokenText(words[k + 1]) === 'debug' && words[k + 2]?.rawText === '.' && words[k + 3] !== undefined) {
				report(words[k], words[k + 3], `Call cannot run Debug.${words[k + 3].rawText}: write it without Call.`, 'Syntax error');
			}
		}
		// `x += 1`: an operator glued to the assignment's `=`.
		if (words[0].kind === 'identifier') {
			const eq = words.findIndex((tok) => tok.rawText === '=');
			const op = words[eq - 1];
			if (eq > 1 && op.end === words[eq].start && COMPOUND_ASSIGNMENT_OPERATORS.has(op.rawText) && words.slice(1, eq - 1).every((tok) => tok.rawText === '.' || tok.kind === 'identifier')) {
				const target = words.slice(0, eq - 1).map((tok) => tok.rawText).join('');
				report(op, words[eq], `'${op.rawText}=' is no VBA operator: write ${target} = ${target} ${op.rawText} ...`, 'Syntax error');
			}
		}
		if (head === 'select' && tokenText(words[1]) === 'case') {
			selects.push({ sawElse: false });
		} else if (head === 'end' && tokenText(words[1]) === 'select') {
			selects.pop();
		} else if (head === 'case') {
			const select = selects[selects.length - 1];
			if (tokenText(words[1]) === 'else') {
				if (select) {
					select.sawElse = true;
				}
			} else {
				if (select?.sawElse) {
					report(words[0], words[words.length - 1], 'A Case after Case Else in the same Select Case can never run.', 'Case without Select Case');
				}
				checkCaseItems(words.slice(1));
			}
		} else if (head === 'do') {
			dos.push(words.length > 1 && (tokenText(words[1]) === 'while' || tokenText(words[1]) === 'until'));
		} else if (head === 'loop') {
			const conditioned = dos.pop();
			const kinds = words.slice(1).filter((tok) => tokenText(tok) === 'while' || tokenText(tok) === 'until');
			if (kinds.length > 1) {
				report(kinds[0], kinds[1], 'A Loop takes one condition, While or Until, not both.', 'Syntax error');
			} else if (kinds.length === 1 && conditioned) {
				report(words[0], kinds[0], 'A Loop with a condition closes a Do with none; this Do has its own.', 'Loop without Do');
			}
		}
	};
	const checkCaseItems = (items: readonly VbaToken[]): void => {
		let depth = 0;
		let start = 0;
		for (let k = 0; k <= items.length; k++) {
			const tok = items[k];
			if (tok?.rawText === '(') {
				depth++;
			} else if (tok?.rawText === ')') {
				depth--;
			}
			if (tok !== undefined && (depth > 0 || tok.rawText !== ',')) {
				continue;
			}
			const item = items.slice(start, k);
			start = k + 1;
			if (tokenText(item[0]) !== 'is') {
				continue;
			}
			if (tokenText(item[1]) === 'like') {
				report(item[0], item[1], "Case Is takes a comparison operator, and Like is none.", 'Syntax error');
			} else if (item.some((t) => tokenText(t) === 'to')) {
				const to = item.find((t) => tokenText(t) === 'to')!;
				report(item[0], to, 'Case Is takes one value: a range with To is a Case of its own.', 'Syntax error');
			}
		}
	};
	for (const tok of toks) {
		if (tok.kind === 'newline') {
			flushLine();
		} else {
			line.push(tok);
		}
	}
	flushLine();
}

/** A line's statements, split at top-level colons; a label's colon ends the label. */
function splitStatements(words: readonly VbaToken[]): VbaToken[][] {
	const out: VbaToken[][] = [[]];
	for (const tok of words) {
		if (tok.kind === 'colon') {
			out.push([]);
		} else {
			out[out.length - 1].push(tok);
		}
	}
	return out.filter((statement) => statement.length > 0);
}

/** True when the colon at `index` stands on a line with nothing else but colons and a comment. */
function lineHoldsOnlyColons(toks: readonly VbaToken[], index: number): boolean {
	for (let j = index - 1; j >= 0 && toks[j].kind !== 'newline'; j--) {
		if (toks[j].kind !== 'colon') {
			return false;
		}
	}
	for (let j = index + 1; j < toks.length && toks[j].kind !== 'newline'; j++) {
		if (toks[j].kind !== 'colon' && toks[j].kind !== 'comment') {
			return false;
		}
	}
	return true;
}

/** `Sub` with no name, and a parameter followed by another word. */
function checkProcedureHeader(
	toks: readonly VbaToken[],
	proc: ProcedureNode,
	place: (offset: number) => Place,
	push: PushFn,
): void {
	let i = firstTokenAtOrAfter(toks, proc.span.start);
	if (i === toks.length) {
		return;
	}
	const header: VbaToken[] = [];
	for (; i < toks.length && toks[i].kind !== 'newline' && toks[i].kind !== 'comment'; i++) {
		header.push(toks[i]);
	}
	if (proc.name === '') {
		const keyword = header.find((tok) => ['sub', 'function', 'property'].includes(tokenText(tok)));
		if (keyword) {
			const error = place(proc.span.start) === 'top' ? 'Expected: identifier' : 'Syntax error';
			push('malformedStatement', `'${keyword.rawText}' needs a name after it. This is a VBE compile error: ${error}.`, { start: keyword.start, end: keyword.end });
		}
		return;
	}
	const nameAt = header.findIndex((tok) => proc.nameSpan !== undefined && tok.start === proc.nameSpan.start);
	const open = nameAt >= 0 && header[nameAt + 1]?.rawText === '(' ? nameAt + 1 : -1;
	if (open < 0) {
		return;
	}
	const close = matchParenFrom(header, open);
	if (close < 0) {
		return;
	}
	let from = open + 1;
	for (let k = open + 1; k <= close; k++) {
		if (k < close && header[k].rawText !== ',') {
			continue;
		}
		const junk = parameterJunk(header.slice(from, k));
		if (junk) {
			push('malformedStatement', `'${junk.rawText}' follows a parameter's name where only As, = or a comma can. This is a VBE compile error: Expected: list separator or ).`, { start: junk.start, end: junk.end });
			return;
		}
		from = k + 1;
	}
}

/** The token after a parameter's name that no parameter form allows there. */
function parameterJunk(param: readonly VbaToken[]): VbaToken | undefined {
	let i = 0;
	while (i < param.length && PARAMETER_MODIFIERS.has(tokenText(param[i]))) {
		i++;
	}
	const name = param[i];
	if (!name) {
		return undefined;
	}
	i++;
	// A type character glued to the name, `s$`, and an array's `()`.
	if (param[i] && param[i].start === name.end && /^[$%&!#@]$/.test(param[i].rawText)) {
		i++;
	}
	if (param[i]?.rawText === '(' && param[i + 1]?.rawText === ')') {
		i += 2;
	}
	const next = param[i];
	if (!next || next.kind === 'comment' || tokenText(next) === 'as' || next.rawText === '=') {
		return undefined;
	}
	return next;
}

/**
 * A statement keyword where a value goes: inside parentheses, after an
 * assignment's `=`, after Print, and after Case.
 */
function checkValueKeywords(source: string, span: Span, context: 'statement' | 'const', push: PushFn): void {
	const toks = statementTokens(source, span);
	const print = toks.findIndex((tok) => tokenText(tok) === 'print');
	let valueFrom: number;
	if (context === 'const') {
		valueFrom = 0;
	} else if (tokenText(toks[0]) === 'case') {
		valueFrom = 1;
	} else if (print >= 0) {
		valueFrom = print + 1;
	} else {
		valueFrom = valueStart(source, span, toks);
	}
	let depth = 0;
	for (let i = 0; i < toks.length; i++) {
		const raw = toks[i].rawText;
		if (raw === '(') {
			depth++;
			continue;
		}
		if (raw === ')') {
			depth--;
			continue;
		}
		if (toks[i].kind !== 'keyword' || !STATEMENT_KEYWORDS.has(tokenText(toks[i]))) {
			continue;
		}
		if (depth === 0 && (valueFrom < 0 || i < valueFrom)) {
			continue;
		}
		// `.Open`, `!Close` and `Type:=1` name a member or an argument.
		const prev = toks[i - 1]?.rawText;
		if (prev === '.' || prev === '!' || toks[i + 1]?.rawText === ':=') {
			continue;
		}
		const error = context === 'const' ? 'Expected: expression' : 'Syntax error';
		push('reservedKeywordInExpression', `'${toks[i].rawText}' is a statement keyword and cannot stand where a value goes. This is a VBE compile error: ${error}.`, { start: span.start + toks[i].start, end: span.start + toks[i].end });
		return;
	}
}

/**
 * Words with a statement's or a print list's meaning only, named where a value
 * goes (issue #318, measured in Excel 16.0): `TypeName(Tab)`, `Main = Print`,
 * `1 + Shared` are a Syntax error, and so are `Tab(5)` and `Spc(5)` outside a
 * Print list. Input, Len and Array take an argument list, and Seek
 * needs its argument: bare, the first three are a Syntax error and
 * Seek is "Argument not optional".
 */
export const VALUE_WORD_ERRORS: ReadonlyMap<string, string> = new Map(Object.entries({
	tab: 'Syntax error',
	spc: 'Syntax error',
	print: 'Syntax error',
	write: 'Syntax error',
	shared: 'Syntax error',
	input: 'Syntax error',
	len: 'Syntax error',
	array: 'Syntax error',
	seek: 'Argument not optional',
}));

function checkValueWords(source: string, span: Span, push: PushFn): void {
	const toks = statementTokens(source, span);
	const valueFrom = valueStart(source, span, toks);
	// A Print list, `Debug.Print "a"; Tab(5)` or `Print #1, Spc(2)`, takes Tab and Spc.
	const print = toks.findIndex((tok, i) => tokenText(tok) === 'print' && (i === 0 || toks[i - 1].rawText === '.'));
	let depth = 0;
	for (let i = 0; i < toks.length; i++) {
		depth += toks[i].rawText === '(' ? 1 : toks[i].rawText === ')' ? -1 : 0;
		const word = tokenText(toks[i]);
		const error = VALUE_WORD_ERRORS.get(word);
		const next = toks[i + 1]?.rawText;
		if (!error || (depth === 0 && (valueFrom < 0 || i < valueFrom)) || ['.', '!'].includes(toks[i - 1]?.rawText ?? '') || ['.', '!', ':='].includes(next ?? '')) {
			continue;
		}
		if ((word === 'tab' || word === 'spc') && print >= 0 && i > print) {
			continue;
		}
		if ((word === 'input' || word === 'len' || word === 'seek' || word === 'array') && (next === '(' || next === '$')) {
			continue;
		}
		const what = word === 'seek' ? 'is a function that needs its file number' : word === 'tab' || word === 'spc' ? 'belongs to a Print list' : 'is a reserved word';
		push(word === 'seek' ? 'argumentCount' : 'reservedKeywordInExpression', `'${toks[i].rawText}' ${what} and cannot stand here as a value. This is a VBE compile error: ${error}.`, { start: span.start + toks[i].start, end: span.start + toks[i].end });
		return;
	}
}

/**
 * Words that cannot qualify a member: `Main = Print.Hi()` is a Syntax error
 * whether or not a module of that name exists (issue #247, measured in Excel
 * 16.0). Get, Put, Open and Stop are statement keywords, reported above.
 */
const NO_QUALIFIER_WORDS: ReadonlySet<string> = new Set(['circle', 'pset', 'scale', 'print', 'input', 'tab', 'spc', 'array', 'lbound', 'date']);

function checkKeywordQualifiers(source: string, span: Span, push: PushFn): void {
	const toks = statementTokens(source, span);
	for (let i = 0; i + 1 < toks.length; i++) {
		const word = tokenText(toks[i]);
		const prev = toks[i - 1]?.rawText;
		// Opening a statement, `Date.Hi` compiles and raises 424 at run time.
		if (!NO_QUALIFIER_WORDS.has(word) || toks[i + 1].rawText !== '.' || prev === '.' || prev === '!' || (i === 0 && word === 'date')) {
			continue;
		}
		const error = i === 0 && word === 'print' ? 'Method not valid without suitable object' : 'Syntax error';
		push('reservedKeywordInExpression', `'${toks[i].rawText}' is a reserved word and cannot qualify a member, even when a module bears the name. This is a VBE compile error: ${error}.`, { start: span.start + toks[i].start, end: span.start + toks[i].end });
		return;
	}
}

/** Where an assignment's value starts, or -1 for any other statement. */
function valueStart(source: string, span: Span, toks: readonly VbaToken[]): number {
	const bare = bareAssignmentTarget(source, span);
	if (!bare || bare.valueTokens.length === 0) {
		return -1;
	}
	// Both are relative to the statement's start.
	const first = bare.valueTokens[0];
	return toks.findIndex((tok) => tok.start === first.start);
}

/** Array bounds and a Const's value: `Dim a(Implements) As Long`, `Const K = Dim`. */
function checkDeclarationKeywords(source: string, group: VariableGroupNode, push: PushFn): void {
	for (const decl of group.declarations) {
		const toks = statementTokens(source, decl.span);
		const eq = group.isConst ? toks.findIndex((tok) => tok.rawText === '=') : -1;
		if (eq >= 0) {
			const value = toks.slice(eq + 1);
			if (value.length > 0) {
				checkValueKeywords(source, { start: decl.span.start + value[0].start, end: decl.span.end }, 'const', push);
			}
			continue;
		}
		checkValueKeywords(source, decl.span, 'statement', push);
	}
}

/** An Enum line that is not `name [= value]`. */
function checkEnumLines(source: string, member: EnumNode, activity: ConditionalActivityTracker | undefined, push: PushFn): void {
	for (const line of member.members) {
		if (activity?.isInactive(line.span)) {
			continue;
		}
		const toks = statementTokens(source, line.span).filter((tok) => tok.kind !== 'comment');
		const name = toks[0];
		if (!name) {
			continue;
		}
		// `.asdf`. A number, `123 = 2`, is invalid-identifier-start's.
		if (name.kind === 'punctuation' || name.kind === 'operator' || name.kind === 'unknown') {
			push('malformedStatement', `'${source.slice(line.span.start, line.span.end).trim()}' is no Enum member: a member is a name, with = and a value if it has one. This is a VBE compile error: Invalid inside Enum.`, line.span);
			continue;
		}
		const next = toks[1];
		if (!next || next.rawText === '=' || (next.start === name.end && /^[$%&!#@]$/.test(next.rawText))) {
			continue;
		}
		const operator = next.kind === 'operator';
		push(
			'malformedStatement',
			operator
				? `'${next.rawText}' follows Enum member '${name.rawText}' where only = can. This is a VBE compile error: Expected: expression.`
				: `'${next.rawText}' follows Enum member '${name.rawText}': a member is a name, with = and a value if it has one. This is a VBE compile error: Invalid inside Enum.`,
			{ start: line.span.start + next.start, end: line.span.start + next.end },
		);
	}
}
