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
	'property', 'declare', 'enum', 'event', 'option', 'call', 'goto', 'gosub', 'exit', 'resume', 'do',
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
	const firstNamed = procedures.find((p) => p.name !== '');
	const place = (offset: number): Place => {
		if (procedures.some((p) => p.name !== '' && offset > p.span.start && offset < p.span.end)) {
			return 'procedure';
		}
		return !firstNamed || offset < firstNamed.span.start ? 'top' : 'after';
	};
	const enums = mod.members.filter((m): m is EnumNode => m.kind === 'Enum');
	const inEnum = (offset: number): boolean => enums.some((e) => offset > e.span.start && offset < e.span.end);
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

	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind === 'Procedure') {
			checkProcedureHeader(toks, member, place, push);
			forEachStatement(member.body, (stmt) => {
				checkValueKeywords(source, stmt.span, 'statement', push);
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
	let i = toks.findIndex((tok) => tok.start >= proc.span.start);
	if (i < 0) {
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
