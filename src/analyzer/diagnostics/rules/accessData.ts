// Rule family: Access SQL literals and DAO recordsets whose failure the
// code itself shows (issue #312). Measured in Access 16.0 (2026-10-02),
// each case creating its own table; each compiles and raises every time.
//
// runtime-argument-value, on SQL in a string literal:
//  - Execute of a SELECT raises 3065, RunSQL of one 2342.
//  - Execute of "", or of text whose first word is no SQL verb but which
//    reads as SQL (`DELET FROM T1`), raises 3078: Execute takes it for the
//    name of a query, and none has that name. OpenRecordset does the same.
//  - A quote left open raises 3075 through Execute or OpenRecordset, and
//    2342 through RunSQL. An INSERT with a parenthesis left open raises
//    3134.
//  - CreateQueryDef of SQL whose first word is no SQL verb raises 3129.
//  - DLookup whose criteria end in a comparison with nothing after raises
//    2342.
//
// host-argument-out-of-range, on a DAO.Recordset local followed in a straight line
// from `Set rs = ....OpenRecordset(...)`:
//  - writing a field, or Update, with no Edit or AddNew since raises 3020;
//  - Edit or AddNew on a snapshot (dbOpenSnapshot) raises 3251;
//  - any use after rs.Close raises 3420.

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
	forEachStatement,
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
				checkSqlLiterals(span, statementTokens(source, span).filter((tok) => tok.kind !== 'comment'), isDatabase, push);
			}
		}, activity);
		checkRecordsets(source, member.body, isDatabase, isRecordset, activity, push);
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
	// `DLookup("Nm", "T1", "ID = ")`: criteria that end in a comparison.
	for (let i = 0; i < toks.length; i++) {
		if (tokenText(toks[i]) !== 'dlookup' || toks[i - 1]?.rawText === '.' || toks[i + 1]?.rawText !== '(') {
			continue;
		}
		const criteria = callArguments(toks, i)?.[2];
		if (criteria?.length === 1 && criteria[0].kind === 'stringLiteral' && COMPARISON_AT_END.test(stringLiteralValue(criteria[0].rawText))) {
			push('runtimeArgumentValue', `The criteria of DLookup end in a comparison with nothing to compare with. This will raise Run-time error '2342': A RunSQL action requires an argument consisting of an SQL statement.`, { start: span.start + criteria[0].start, end: span.start + criteria[0].end });
		}
	}
}

/** Why Access refuses this SQL text, with the error it raises, or undefined. */
function sqlProblem(kind: string, sql: string): string | undefined {
	const first = /^[\s(]*([A-Za-z]+)/.exec(sql)?.[1]?.toLowerCase();
	const readsAsSql = /\b(?:from|into|set|values|where)\b/i.test(sql);
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
	if (first && !SQL_VERBS.has(first) && readsAsSql) {
		return `"${first}" starts no SQL statement, so the text is taken for the name of a table or query, and none has that name. This will raise Run-time error '3078': The Microsoft Access database engine cannot find the input table or query`;
	}
	if (kind === 'execute' && first === 'select') {
		return `Execute runs an action query, and this is a SELECT. This will raise Run-time error '3065': Cannot execute a select query`;
	}
	if (quotes) {
		return `The SQL leaves a quote open. This will raise Run-time error '3075': Syntax error in string in query expression`;
	}
	if (kind === 'execute' && first === 'insert' && openParenthesis(sql)) {
		return `The INSERT leaves a parenthesis open. This will raise Run-time error '3134': Syntax error in INSERT INTO statement`;
	}
	return undefined;
}

/** Whether a single-quoted string of the SQL is left open: `''` inside one is a quote. */
function openQuote(sql: string): boolean {
	return (sql.replace(/''/g, '').match(/'/g)?.length ?? 0) % 2 === 1;
}

/** Whether a parenthesis outside quotes is left open. */
function openParenthesis(sql: string): boolean {
	let depth = 0;
	let quoted = false;
	for (const ch of sql) {
		if (ch === "'") {
			quoted = !quoted;
		} else if (!quoted && ch === '(') {
			depth++;
		} else if (!quoted && ch === ')') {
			depth--;
		}
	}
	return depth > 0;
}

/** What a straight line of statements knows of a recordset local. */
interface RecordsetState {
	snapshot: boolean;
	editing: boolean;
	closed: boolean;
}

function checkRecordsets(
	source: string,
	body: readonly BodyNode[],
	isDatabase: (lower: string) => boolean,
	isRecordset: (lower: string) => boolean,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	const state = new Map<string, RecordsetState>();
	const forget = (names: Iterable<string>): void => {
		for (const lower of names) {
			state.delete(lower);
		}
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
		const toks = statementTokensAfterLeadingLabel(source, node.span).filter((tok) => tok.kind !== 'comment');
		const at = (from: number, to: number) => ({ start: node.span.start + toks[from].start, end: node.span.start + toks[to].end });
		const set = setAssignmentTarget(source, node.span);
		if (set) {
			const lower = set.name.toLowerCase();
			forget(namesIn(source, node.span));
			const value = set.valueTokens.filter((tok) => tok.kind !== 'comment');
			const open = value.findIndex((tok, k) => tokenText(tok) === 'openrecordset' && value[k - 1]?.rawText === '.' && databaseAt(value, k - 2, isDatabase));
			if (isRecordset(lower) && open > 0 && value[open + 1]?.rawText === '(' && matchParenFrom(value, open + 1) === value.length - 1) {
				const type = callArguments(value, open)?.[1]?.map((tok) => tokenText(tok)).join('');
				state.set(lower, { snapshot: type === 'dbopensnapshot' || type === '4', editing: false, closed: false });
			}
			return;
		}
		const lower = tokenName(toks[0])?.toLowerCase();
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
			forget([...namesIn(source, node.span)].filter((name) => !state.has(name)));
			return;
		}
		const shown = toks[0].rawText;
		// Any use after Close.
		if (held.closed) {
			push('hostArgumentOutOfRange', `'${shown}' was closed above, and a closed recordset has no members. This will raise Run-time error '3420': Object invalid or no longer set.`, at(0, 0));
			state.delete(lower!);
			return;
		}
		const member = toks[1]?.rawText === '.' ? tokenText(toks[2]) : undefined;
		const eq = toks.findIndex((tok) => tok.rawText === '=');
		// `rs!Nm = x`, `rs("Nm") = x`, `rs.Fields("Nm") = x`, `rs.Fields("Nm").Value = x`.
		const fieldWrite = eq > 1 && (toks[1]?.rawText === '!' || toks[1]?.rawText === '(' || member === 'fields');
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
		if ((member === 'edit' || member === 'addnew') && toks.length === 3) {
			if (held.snapshot) {
				push('hostArgumentOutOfRange', `'${shown}' is a snapshot, which cannot be changed. This will raise Run-time error '3251': Operation is not supported for this type of object.`, at(2, 2));
				state.delete(lower!);
				return;
			}
			held.editing = true;
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
		snapshot: () => new Map([...state].map(([lower, held]) => [lower, { ...held }])),
		restore: (saved) => {
			state.clear();
			for (const [lower, held] of saved) {
				state.set(lower, { ...held });
			}
		},
		forget,
		touches: (stmt) => namesIn(source, stmt.span),
	});
}
