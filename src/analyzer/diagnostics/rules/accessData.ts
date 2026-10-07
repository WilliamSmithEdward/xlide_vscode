// Rule family: Access SQL literals and DAO recordsets whose failure the
// code itself shows (issue #312). Measured in Access 16.0 (2026-10-02),
// each case creating its own table; each compiles and raises every time.
//
// runtime-argument-value, on SQL in a string literal:
//  - Execute of a SELECT raises 3065, RunSQL of one 2342; a SELECT ... INTO
//    makes a table and runs, and a TRANSFORM raises 3065 too (issue #611).
//  - Execute of "", or of text whose first word is no SQL verb but which
//    reads as SQL (`DELET FROM T1`), raises 3078: Execute takes it for the
//    name of a query, and none has that name. OpenRecordset does the same.
//  - A quote left open raises 3075 through Execute or OpenRecordset, and
//    2342 through RunSQL. A double-quoted string is a string too, and a '
//    inside it is text. An INSERT with a parenthesis left open, or with
//    neither VALUES nor SELECT, raises 3134; any other SQL with one left
//    open raises 3075.
//  - CreateQueryDef of SQL whose first word is no SQL verb raises 3129.
//  - DLookup whose criteria end in a comparison with nothing after raises
//    2342. The other domain functions raise 3075 there, and every one does
//    for criteria ending in AND or OR, or leaving a quote open (#611).
//
// host-argument-out-of-range, on a DAO.Recordset local followed in a straight line
// from `Set rs = ....OpenRecordset(...)`:
//  - writing a field, or Update, with no Edit or AddNew since raises 3020; a
//    Move ends an Edit, and `!Nm = x` inside `With rs` is rs's;
//  - Edit, AddNew or Delete on a snapshot (dbOpenSnapshot), and Edit or
//    AddNew on a forward-only one, raise 3251; on one opened dbReadOnly,
//    3027;
//  - any use after rs.Close raises 3420, through another name for it too.

import type { ConditionalActivityTracker } from '../../conditional/conditionalCompilation';
import { statementLabelDeclaration } from '../../flow/procedureLabels';
import type { VbaToken } from '../../lexer/tokenKinds';
import { splitTopLevelTokenGroups } from '../../lexer/tokenHelpers';
import type { BodyNode, ModuleNode, Span } from '../../parser/nodes';
import { isLeafStatement } from '../../parser/nodes';
import type { PushFn } from '../analysisContext';
import { walkEnteringBlocks } from '../dataflow';
import { normalizeType, stringLiteralValue, typeEnvironmentFor } from '../typeInference';
import type { buildModuleSymbols } from '../../symbols/buildModuleSymbols';
import {
	activeModuleMembers,
	blockHeaderLineSpan,
	forEachStatement,
	forEachVariableGroup,
	matchParenFrom,
	setAssignmentTarget,
	statementAndBranchSpans,
	statementTokens,
	statementTokensAfterLeadingLabel,
	tokenName,
	tokenText,
} from '../walker';
import { namesIn } from './shared';

/** The words an Access SQL statement starts with. */
const SQL_VERBS: ReadonlySet<string> = new Set(['select', 'insert', 'update', 'delete', 'create', 'alter', 'drop', 'transform', 'parameters', 'procedure']);

const COMPARISON_AT_END = /(?:=|<>|<|>|\blike|\bin)\s*$/i;

/** Access's domain aggregate functions, which read their criteria as SQL. */
const DOMAIN_FUNCTIONS: ReadonlySet<string> = new Set(['dlookup', 'dcount', 'dsum', 'davg', 'dmin', 'dmax', 'dfirst', 'dlast', 'dstdev', 'dstdevp', 'dvar', 'dvarp']);

export function checkAccessData(
	source: string,
	mod: ModuleNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
	host: string | undefined,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	if (host !== 'Access') {
		return;
	}
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		const env = typeEnvironmentFor(symbols, member);
		const isDatabase = (lower: string): boolean => ['dao.database', 'database'].includes(normalizeType(env.get(lower)) ?? '');
		const isRecordset = (lower: string): boolean => ['dao.recordset', 'recordset'].includes(normalizeType(env.get(lower)) ?? '');
		forEachStatement(member.body, (stmt) => {
			for (const span of statementAndBranchSpans(stmt)) {
				checkSqlLiterals(span, statementTokens(source, span), isDatabase, push);
			}
		}, activity);
		// An error handler may resume after a failed Open/Edit/Close.
		if (/\bon\s+error\b/i.test(source.slice(member.span.start, member.span.end))) { continue; }
		const locals = new Set<string>();
		forEachVariableGroup(member.body, group => {
			if (!group.isConst && group.modifier.toLowerCase() !== 'static') {
				for (const decl of group.declarations) { locals.add(decl.name.toLowerCase()); }
			}
		}, activity);
		checkRecordsets(source, member.body, isDatabase, lower => locals.has(lower) && isRecordset(lower), activity, push);
	}
}

/** `CurrentDb`, `Application.CurrentDb`, or a local As DAO.Database, ending at `toks[end]`. */
function databaseAt(toks: readonly VbaToken[], end: number, isDatabase: (lower: string) => boolean): boolean {
	const word = tokenText(toks[end]);
	const qualified = toks[end - 1]?.rawText === '.';
	if (word === 'currentdb' || word === 'codedb') {
		return !qualified || tokenText(toks[end - 2]) === 'application';
	}
	if (word === ')' && tokenText(toks[end - 1]) === '(' && ['currentdb', 'codedb'].includes(tokenText(toks[end - 2]))) {
		return toks[end - 3]?.rawText !== '.' || tokenText(toks[end - 4]) === 'application';
	}
	return !qualified && tokenName(toks[end]) !== undefined && isDatabase(word);
}

/** The arguments of the call whose name is at `i`, written with parentheses or as a statement. */
function callArguments(toks: readonly VbaToken[], i: number): VbaToken[][] | undefined {
	if (toks[i + 1]?.rawText === '(') {
		const close = matchParenFrom(toks, i + 1);
		return close > i + 2 ? splitTopLevelTokenGroups([...toks], i + 2, ',', close) : close === i + 2 ? [] : undefined;
	}
	return i + 1 < toks.length ? splitTopLevelTokenGroups([...toks], i + 1, ',', toks.length) : [];
}

function checkSqlLiterals(span: Span, toks: readonly VbaToken[], isDatabase: (lower: string) => boolean, push: PushFn): void {
	for (let i = 2; i < toks.length; i++) {
		const word = tokenText(toks[i]);
		const onDatabase = toks[i - 1].rawText === '.' && databaseAt(toks, i - 2, isDatabase);
		const doCmd = toks[i - 1].rawText === '.' && tokenText(toks[i - 2]) === 'docmd' && (toks[i - 3]?.rawText !== '.' || tokenText(toks[i - 4]) === 'application');
		const kind = onDatabase && (word === 'execute' || word === 'openrecordset' || word === 'createquerydef') ? word : doCmd && word === 'runsql' ? 'runsql' : undefined;
		if (!kind) {
			continue;
		}
		const args = callArguments(toks, i);
		const arg = args?.[kind === 'createquerydef' ? 1 : 0];
		if (arg?.length !== 1 || arg[0].kind !== 'stringLiteral') {
			continue;
		}
		const problem = sqlProblem(kind, stringLiteralValue(arg[0].rawText));
		if (problem) {
			push('runtimeArgumentValue', `${problem}.`, { start: span.start + arg[0].start, end: span.start + arg[0].end });
		}
	}
	// `DLookup("Nm", "T1", "ID = ")`: criteria that end in a comparison, in
	// AND or OR, or leave a quote open.
	for (let i = 0; i < toks.length; i++) {
		const name = tokenText(toks[i]);
		if (!DOMAIN_FUNCTIONS.has(name) || toks[i - 1]?.rawText === '.' || toks[i + 1]?.rawText !== '(') {
			continue;
		}
		const criteria = callArguments(toks, i)?.[2];
		if (criteria?.length !== 1 || criteria[0].kind !== 'stringLiteral') {
			continue;
		}
		const text = stringLiteralValue(criteria[0].rawText);
		const at = { start: span.start + criteria[0].start, end: span.start + criteria[0].end };
		const shown = toks[i].rawText;
		if (COMPARISON_AT_END.test(text)) {
			push('runtimeArgumentValue', name === 'dlookup'
				? `The criteria of DLookup end in a comparison with nothing to compare with. This will raise Run-time error '2342': A RunSQL action requires an argument consisting of an SQL statement.`
				: `The criteria of ${shown} end in a comparison with nothing to compare with. This will raise Run-time error '3075': Syntax error (missing operator) in query expression.`, at);
		} else if (/\b(?:and|or|not)\s*$/i.test(text)) {
			push('runtimeArgumentValue', `The criteria of ${shown} end in '${/(\w+)\s*$/.exec(text)![1]}' with nothing after it. This will raise Run-time error '3075': Syntax error (missing operator) in query expression.`, at);
		} else if (openQuote(text)) {
			push('runtimeArgumentValue', `The criteria of ${shown} leave a quote open. This will raise Run-time error '3075': Syntax error in string in query expression.`, at);
		}
	}
}

/** Why Access refuses this SQL text, with the error it raises, or undefined. */
function sqlProblem(kind: string, sql: string): string | undefined {
	const first = /^[\s(]*([A-Za-z]+)/.exec(sql)?.[1]?.toLowerCase();
	const quotes = openQuote(sql);
	if (kind === 'runsql') {
		if (first === 'select') {
			return `RunSQL runs an action query, and "${sql}" is a SELECT. This will raise Run-time error '2342': A RunSQL action requires an argument consisting of an SQL statement`;
		}
		return quotes ? `The SQL leaves a quote open. This will raise Run-time error '2342': A RunSQL action requires an argument consisting of an SQL statement` : undefined;
	}
	if (kind === 'createquerydef') {
		return first && !SQL_VERBS.has(first)
			? `"${first}" starts no SQL statement. This will raise Run-time error '3129': Invalid SQL statement; expected 'DELETE', 'INSERT', 'PROCEDURE', 'SELECT', or 'UPDATE'`
			: undefined;
	}
	if (sql.trim() === '' && kind === 'execute') {
		return `Execute has no SQL to run, and no query is named "". This will raise Run-time error '3078': The Microsoft Access database engine cannot find the input table or query`;
	}
	// Execute/OpenRecordset may receive the name of a runtime QueryDef,
	// even when that name resembles misspelled SQL. Absence is not proven.
	if (quotes) {
		return `The SQL leaves a quote open. This will raise Run-time error '3075': Syntax error in string in query expression`;
	}
	if (kind === 'execute' && first === 'insert' && (openParenthesis(sql) || !/^\s*insert\s+into\s+(?:\[[^\]]*\]|[^\s(]+)\s*(?:\([^)]*\)\s*)?(?:values|select)\b/i.test(sql))) {
		return `The INSERT ${openParenthesis(sql) ? 'leaves a parenthesis open' : 'has neither VALUES nor a SELECT'}. This will raise Run-time error '3134': Syntax error in INSERT INTO statement`;
	}
	if (openParenthesis(sql)) {
		return `The SQL leaves a parenthesis open. This will raise Run-time error '3075': Missing ), ], or Item in query expression`;
	}
	// SELECT ... INTO makes a table, an action query (issue #611).
	if (kind === 'execute' && ((first === 'select' && !/\binto\b/i.test(outsideStrings(sql))) || first === 'transform')) {
		return `Execute runs an action query, and this is a ${first === 'select' ? 'SELECT' : 'TRANSFORM'}. This will raise Run-time error '3065': Cannot execute a select query`;
	}
	return undefined;
}

/**
 * The SQL with its strings blanked, and whether one is left open. A string
 * is single- or double-quoted, its quote doubled inside it (issue #611:
 * `"it's"` is one string).
 */
function scanStrings(sql: string): { outside: string; open: boolean } {
	let outside = '';
	let quote: string | undefined;
	for (let i = 0; i < sql.length; i++) {
		const ch = sql[i];
		if (quote) {
			if (ch === quote && sql[i + 1] === quote) {
				i++;
			} else if (ch === quote) {
				quote = undefined;
			}
			outside += ' ';
			continue;
		}
		if (ch === "'" || ch === '"') {
			quote = ch;
			outside += ' ';
			continue;
		}
		outside += ch;
	}
	return { outside, open: quote !== undefined };
}

/** Whether a string of the SQL is left open. */
function openQuote(sql: string): boolean {
	return scanStrings(sql).open;
}

/** The SQL outside its strings. */
function outsideStrings(sql: string): string {
	return scanStrings(sql).outside;
}

/** Whether a parenthesis outside strings is left open. */
function openParenthesis(sql: string): boolean {
	let depth = 0;
	for (const ch of outsideStrings(sql)) {
		depth += ch === '(' ? 1 : ch === ')' ? -1 : 0;
	}
	return depth > 0;
}

/** A copy of the states, names that share one recordset sharing one copy. */
function copyStates(states: ReadonlyMap<string, RecordsetState>): Map<string, RecordsetState> {
	const copies = new Map<RecordsetState, RecordsetState>();
	return new Map([...states].map(([lower, held]) => {
		let copy = copies.get(held);
		if (!copy) {
			copy = { ...held };
			copies.set(held, copy);
		}
		return [lower, copy];
	}));
}

/** What a straight line of statements knows of a recordset local. */
interface RecordsetState {
	snapshot: boolean;
	/** dbOpenForwardOnly: Edit and AddNew raise 3251. */
	forwardOnly: boolean;
	/** Opened dbReadOnly: Edit and AddNew raise 3027. */
	readOnly: boolean;
	editing: boolean;
	closed: boolean;
}

/** The methods that move a recordset's current record, ending an Edit (issue #611). */
const MOVES: ReadonlySet<string> = new Set(['movefirst', 'movelast', 'movenext', 'moveprevious', 'move', 'findfirst', 'findlast', 'findnext', 'findprevious', 'seek', 'requery']);

function checkRecordsets(
	source: string,
	body: readonly BodyNode[],
	isDatabase: (lower: string) => boolean,
	isRecordset: (lower: string) => boolean,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	const state = new Map<string, RecordsetState>();
	const withSubjects: Array<string | undefined> = [];
	const forget = (names: Iterable<string>): void => {
		// Escaping one alias makes every reference to that recordset uncertain.
		const escaped = new Set([...names].map(lower => state.get(lower)).filter(Boolean));
		for (const [lower, held] of state) { if (escaped.has(held)) { state.delete(lower); } }
	};
	const visit = (node: BodyNode): void => {
		if (!isLeafStatement(node)) {
			return;
		}
		if (statementLabelDeclaration(source, node.span)) {
			state.clear();
		}
		if (node.kind === 'Statement' && node.singleLineIfBranches) {
			forget(namesIn(source, node.span));
			return;
		}
		const own = statementTokensAfterLeadingLabel(source, node.span);
		// Inside `With rs`, `!Nm = x` and `.Edit` are rs's (issue #611).
		const subject = withSubjects[withSubjects.length - 1];
		const toks = subject && (own[0]?.rawText === '!' || own[0]?.rawText === '.') ? [{ ...own[0], kind: 'identifier' as const, rawText: subject, end: own[0].start }, ...own] : own;
		const at = (from: number, to: number) => ({ start: node.span.start + toks[from].start, end: node.span.start + toks[to].end });
		const set = setAssignmentTarget(source, node.span);
		if (set) {
			const lower = set.name.toLowerCase();
			const value = set.valueTokens.filter((tok) => tok.kind !== 'comment');
			// `Set r2 = rs` is another name for the same recordset.
			const alias = value.length === 1 ? tokenName(value[0])?.toLowerCase() : undefined;
			const shared = alias ? state.get(alias) : undefined;
			forget(namesIn(source, node.span));
			const open = value.findIndex((tok, k) => tokenText(tok) === 'openrecordset' && value[k - 1]?.rawText === '.' && databaseAt(value, k - 2, isDatabase));
			if (isRecordset(lower) && open > 0 && value[open + 1]?.rawText === '(' && matchParenFrom(value, open + 1) === value.length - 1) {
				const args = callArguments(value, open);
				const type = args?.[1]?.map((tok) => tokenText(tok)).join('');
				const options = args?.[2]?.map((tok) => tokenText(tok)).join('');
				state.set(lower, {
					snapshot: type === 'dbopensnapshot' || type === '4',
					forwardOnly: type === 'dbopenforwardonly' || type === '8',
					readOnly: options === 'dbreadonly' || options === '4',
					editing: false,
					closed: false,
				});
			}
			if (isRecordset(lower) && shared) {
				state.set(lower, shared);
				state.set(alias!, shared);
			}
			return;
		}
		const lower = tokenName(toks[0])?.toLowerCase();
		const member = toks[1]?.rawText === '.' ? tokenText(toks[2]) : undefined;
		const eq = toks.findIndex((tok) => tok.rawText === '=');
		const fieldWrite = eq > 1 && (toks[1]?.rawText === '!' || toks[1]?.rawText === '(' || member === 'fields');
		if (fieldWrite && toks.slice(eq + 1).some(tok => tokenName(tok))) {
			// Evaluate helpers/getters before checking the receiver's edit or
			// closed state: they can start Edit or replace a closed recordset.
			state.clear();
			return;
		}
		// A recordset read past the statement's head: `Main = rs!Nm`, `x = rs.EOF`.
		for (const name of namesIn(source, node.span)) {
			const other = name === lower ? undefined : state.get(name);
			if (!other) {
				continue;
			}
			const uses = toks.map((tok, k) => (tokenText(tok) === name && toks[k - 1]?.rawText !== '.' ? k : -1)).filter((k) => k >= 0);
			if (uses.some((k) => !['.', '!', '('].includes(toks[k + 1]?.rawText ?? ''))) {
				state.delete(name); // passed whole, which may change it
			} else if (other.closed && uses.length > 0) {
				push('hostArgumentOutOfRange', `'${toks[uses[0]].rawText}' was closed above, and a closed recordset has no members. This will raise Run-time error '3420': Object invalid or no longer set.`, at(uses[0], uses[0]));
				state.delete(name);
			}
		}
		const held = lower ? state.get(lower) : undefined;
		if (!held) {
			// Unmodeled statements may run helpers, including in their RHS.
			state.clear();
			return;
		}
		const shown = toks[0].rawText;
		// Any use after Close.
		if (held.closed) {
			push('hostArgumentOutOfRange', `'${shown}' was closed above, and a closed recordset has no members. This will raise Run-time error '3420': Object invalid or no longer set.`, at(0, 0));
			state.delete(lower!);
			return;
		}
		// `rs!Nm = x`, `rs("Nm") = x`, `rs.Fields("Nm") = x`, `rs.Fields("Nm").Value = x`.
		if (fieldWrite || (member === 'update' && toks.length === 3)) {
			if (!held.editing) {
				push('hostArgumentOutOfRange', `'${shown}' is not being edited: no Edit or AddNew came since it was opened or last updated. This will raise Run-time error '3020': Update or CancelUpdate without AddNew or Edit.`, at(0, eq > 1 ? eq - 1 : 2));
				state.delete(lower!);
				return;
			}
			if (member === 'update') {
				held.editing = false;
			}
			return;
		}
		if ((member === 'edit' || member === 'addnew' || member === 'delete') && toks.length === 3) {
			if (held.snapshot || (held.forwardOnly && member !== 'delete')) {
				push('hostArgumentOutOfRange', `'${shown}' is ${held.snapshot ? 'a snapshot' : 'forward-only'}, which cannot be changed. This will raise Run-time error '3251': Operation is not supported for this type of object.`, at(2, 2));
				state.delete(lower!);
				return;
			}
			if (held.readOnly && member !== 'delete') {
				push('hostArgumentOutOfRange', `'${shown}' was opened dbReadOnly, which cannot be changed. This will raise Run-time error '3027': Cannot update. Database or object is read-only.`, at(2, 2));
				state.delete(lower!);
				return;
			}
			if (member !== 'delete') {
				held.editing = true;
			} else {
				// Delete moves nothing, and what it leaves is not followed.
				state.delete(lower!);
			}
			return;
		}
		if (member !== undefined && MOVES.has(member)) {
			held.editing = false;
			return;
		}
		if (member === 'close' && toks.length === 3) {
			held.closed = true;
			return;
		}
		if (member === 'cancelupdate' && toks.length === 3) {
			held.editing = false;
			return;
		}
		// A read leaves it as it is; anything else may change it.
		if (!(eq > 1 && eq < toks.length) || toks.slice(0, eq).some((tok) => tokenText(tok) === lower)) {
			state.delete(lower!);
		}
	};
	walkEnteringBlocks(source, body, (node) => activity?.isInactive(node.span) === true, visit, {
		// One copy per recordset, so two names for it stay one.
		snapshot: () => copyStates(state),
		restore: (saved) => {
			state.clear();
			for (const [lower, held] of copyStates(saved)) {
				state.set(lower, held);
			}
		},
		forget,
		// `With rs` reads rs, and a body line reaching it by `!` or `.` names it.
		touches: (stmt) => {
			const toks = statementTokensAfterLeadingLabel(source, stmt.span);
			if (tokenText(toks[0]) === 'with' && toks.length === 2) {
				return new Set<string>();
			}
			const names = namesIn(source, stmt.span);
			const subject = withSubjects[withSubjects.length - 1];
			return subject && (toks[0]?.rawText === '!' || toks[0]?.rawText === '.') ? new Set([...names, subject.toLowerCase()]) : names;
		},
		enter: (node) => {
			if (node.kind !== 'WithBlock') {
				return;
			}
			const header = statementTokensAfterLeadingLabel(source, blockHeaderLineSpan(source, node.span));
			const name = header.length === 2 ? tokenName(header[1]) : undefined;
			withSubjects.push(name && state.has(name.toLowerCase()) ? name : undefined);
		},
		exit: (node) => {
			if (node.kind === 'WithBlock') {
				withSubjects.pop();
			}
		},
		withBodyRunsThrough: true,
	});
}
