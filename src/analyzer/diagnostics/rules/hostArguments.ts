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
//      Cells(0, 1), Cells(1, 0), Cells(0) -> 1004. Range("$A$0"),
//      Range("A$0"), Range("Sheet1!$A$0"), Range("0:0"), Range("$XFE:$XFE")
//      -> 1004; Range("A:A") and Range("XFD1048576") run. Range("A0"),
//      Range("A1048577") and Range("XFE1") raise 1004 too, but only while
//      no workbook name is spelled that way, so they are not judged (measured
//      2026-10-01): A0, XFE1, A1048577 and XFE are names Names.Add accepts,
//      and Range then finds them, alone, in "A0:D4", in Range("A0", "D4")
//      and after "Sheet1!". A name cannot hold `$` or start with a digit.
//      Range("A1").Offset(-1, 0), Range("A1").Offset(0, -1) -> 1004;
//      Range("B2").Offset(-1, -1) runs. Resize(0, 1), Resize(1, 0), Resize(0),
//      Resize(-1, 1) -> 1004. Past the bottom and right edges (issue #182,
//      measured 2026-09-29): Cells(1048577, 1), Cells(1, 16385), Rows(1048577),
//      Columns(16385), Range("A1048576").Offset(1, 0),
//      Range("A2").Resize(1048576), Range("B2").Cells(1048576, 1) and
//      Range("A5").Rows(1048573) -> 1004; Cells(1048576, 16384) and
//      Range("A2").Offset(0) or Offset(-1) run. On a range, Cells, Item, Rows
//      and Columns count from its top-left cell (issue #275, measured
//      2026-10-01): Range("C3").Cells(-1, -1) is A1, Range("B2:C3").Rows(0)
//      is B1:C1, and one index w columns wide is Cells((i - 1) \ w + 1,
//      (i - 1) Mod w + 1), so Range("B2:C3").Cells(0) is A2. Each raises
//      1004 only where it lands above row 1 or left of column A:
//      Range("A1").Cells(0), Range("B2").Cells(-1). A range the code does not
//      spell out, a variable or ActiveCell, is not judged.
//      ActiveDocument.Range(-1, 0), Range(0, -1),
//      Range(1, 0) -> 4608 in Word; Range(0, 0) runs.
//      Range("") and Range(" ") -> 1004 (issue #276). A name that looks
//      like an address past the sheet, ZZZZ1, may be a workbook name, and
//      Range("A1:ZZZZ1") then runs, so it is not judged.
//      ActiveSheet.Shapes(0) and Sheets(1).Shapes(0) -> -2147024809: a
//      Worksheet and a Chart both have Shapes (issue #276).
//  - sheet-name-invalid
//      Worksheets(1).Name = "a:b", "", a 32-character name, or a name holding
//      any of : \ / ? * [ ] -> 1004. A 31-character name runs. "History" in
//      any case, and an apostrophe first or last, -> 1004, and so does a
//      name String$, Space$ or & spell out (issue #276, measured
//      2026-10-01); "History ", "History1" and "a'b" run.
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
import type { ConditionalActivityTracker } from '../../conditional/conditionalCompilation';
import { resolveReceiverTypeAt } from '../../completion/memberAccess';
import type { VbaToken } from '../../lexer/tokenKinds';
import type { BodyNode, ProcedureNode, Span } from '../../parser/nodes';
import { isLeafStatement } from '../../parser/nodes';
import { jumpTargetLabelDeclaration } from '../../flow/procedureLabels';
import { splitTopLevelTokenGroups } from '../../lexer/tokenHelpers';
import { buildModuleSymbols } from '../../symbols/buildModuleSymbols';
import { procedureSymbolFor, type AnalyzeModuleOptions, type PushFn } from '../analysisContext';
import type { SheetChanges, WorkbookSheetInfo } from '../../symbols/sheetChanges';
import { checkEachCounterPass, loopCountersAt } from '../loopCounters';
import { knownLocalLiteralValuesAt, normalizeType, stringLiteralValue, typeEnvironmentFor, type KnownLocalValue } from '../typeInference';
import {
	bareAssignmentTarget,
	blockHeaderLineSpan,
	firstExecutableTokenIndex,
	matchParenFrom,
	statementTokens,
	statementTokensAfterLeadingLabel,
	tokenName,
	tokenText,
	type ProcedureStatementVisitor,
} from '../walker';
import { worksheetFunctionRefusal } from './worksheetFunctionArguments';
import { WORD_BUILTIN_STYLES } from '../../host/wordBuiltinStyles';

const EXCEL_MAX_ROW = 1048576;
const EXCEL_MAX_COLUMN = 16384;
/**
 * Range members whose arguments are a cell address, a row and column, an
 * offset or a size, not an index into the range they return. Each has its
 * own check. `Range("A2").Offset(0)` and `Offset(-1)` run (issue #182).
 */
const RANGE_COORDINATE_MEMBERS: ReadonlySet<string> = new Set(['cells', 'range', 'offset', 'resize']);
/** The members that count from a range's own top-left cell (issue #275). */
const RANGE_RELATIVE_MEMBERS: ReadonlySet<string> = new Set(['cells', 'item', 'rows', 'columns']);
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
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
	sheets?: WorkbookSheetsCheck,
): ProcedureStatementVisitor {
	const model = memberCtx.model;
	const host = model?.hostName ?? 'Excel';
	if (host !== 'Excel' && host !== 'Word' && host !== 'PowerPoint') {
		return () => undefined;
	}
	const workbook = host === 'Excel' ? sheets : undefined;
	const moduleNames = new Set<string>();
	const moduleArrays = new Set<string>();
	for (const child of symbols.root.children ?? []) {
		if (child.isArray) {
			moduleArrays.add(child.name.toLowerCase());
		}
	}
	for (const child of symbols.root.children ?? []) {
		moduleNames.add(child.name.toLowerCase());
	}
	// A Public procedure of another module named Cells, Worksheets or Range
	// takes the call from Excel's (issue #280, measured in Excel 16.0).
	for (const type of memberCtx.projectClassMembers ?? []) {
		if (type.kind === 'standardModule') {
			for (const member of type.members) {
				moduleNames.add(member.name.toLowerCase());
			}
		}
	}
	return (proc: ProcedureNode) => {
		const env = typeEnvironmentFor(symbols, proc);
		const sourceNames = new Set(moduleNames);
		for (const child of procedureSymbolFor(symbols, proc)?.children ?? []) {
			sourceNames.add(child.name.toLowerCase());
		}
		// A single-line If is one span here: its branches would otherwise be
		// walked twice, once inside the whole statement and once on their own.
		// Array variables, so a range read into one is named as an array. A
		// local shadows a module-level name.
		const arrays = new Set(moduleArrays);
		for (const child of procedureSymbolFor(symbols, proc)?.children ?? []) {
			if (child.isArray) {
				arrays.add(child.name.toLowerCase());
			} else {
				arrays.delete(child.name.toLowerCase());
			}
		}
		// `Cells(r, 1)` inside `For r = 0 To 3`: each pass's value (issue #200).
		const counters = loopCountersAt(source, proc.body, activity);
		// A local known to hold one value here: `r = 0` then `Cells(r, 1)`
		// (issue #345, measured in Excel 16.0).
		let valuesAt: ReturnType<typeof knownLocalLiteralValuesAt> | undefined;
		let documentsAt: ReturnType<typeof newDocumentsAt> | undefined;
		let sheetsAt: ReturnType<typeof activeSheetsAt> | undefined;
		let factsAt: ReturnType<typeof sheetFactsAt> | undefined;
		// A sheet the code just protected (issue #471).
		if (host === 'Excel' && /\.\s*protect\b/i.test(source.slice(proc.span.start, proc.span.end))) {
			checkProtectedSheets(source, proc, activity, push);
		}
		return (stmt) => {
			const stmtCounters = counters.get(stmt);
			const known = (valuesAt ??= knownLocalLiteralValuesAt(source, proc, symbols, activity))(stmt);
			const heldBy = (arg: readonly VbaToken[]): KnownLocalValue | undefined => {
				const toks = arg.filter((tok) => tok.kind !== 'comment');
				const held = toks.length === 1 ? known.get(tokenName(toks[0])?.toLowerCase() ?? '') : undefined;
				return held && !held.contentMutated ? held : undefined;
			};
			const knownNumber = (arg: readonly VbaToken[]): number | undefined => {
				const held = heldBy(arg);
				return held?.kind === 'number' && Number.isInteger(held.value) ? held.value as number : undefined;
			};
			const stringOf = (arg: readonly VbaToken[]): string | undefined => {
				const held = heldBy(arg);
				return literalString(arg) ?? (held?.kind === 'string' ? held.value as string : undefined);
			};
			const literalOrKnown = (arg: readonly VbaToken[]): number | undefined => integerLiteralValue(arg) ?? knownNumber(arg);
			// A document the procedure just added (issue #497).
			const documents = host === 'Word' ? (documentsAt ??= newDocumentsAt(source, proc, activity)).get(stmt) : undefined;
			if (documents) {
				checkNewDocumentUses(stmt.span, statementTokens(source, stmt.span), documents, literalOrKnown, push);
			}
			// A sheet the code knows is not active (issue #470).
			const sheetState = host === 'Excel' ? (sheetsAt ??= activeSheetsAt(source, proc, activity)).get(stmt) : undefined;
			if (sheetState) {
				checkUnqualifiedCorners(stmt.span, statementTokens(source, stmt.span), sheetState, push);
			}
			// Intersect and Union of literal ranges, and a sheet the code just added (issue #472).
			if (host === 'Excel') {
				checkSheetFacts(stmt.span, statementTokens(source, stmt.span), (factsAt ??= sheetFactsAt(source, proc, activity)).get(stmt), push);
			}
			if (!stmtCounters) {
				checkSpan(source, stmt.span, host, model, memberCtx, env, arrays, sourceNames, literalOrKnown, push, stringOf);
				if (workbook) {
					checkWorkbookSheetAccess(source, stmt.span, workbook, sourceNames, literalOrKnown, push);
				}
				return;
			}
			checkEachCounterPass(source, stmt.span, stmtCounters, () => undefined, (values, report) => {
				const valueOf = (arg: readonly VbaToken[]): number | undefined => {
					const literal = integerLiteralValue(arg);
					if (literal !== undefined || values.size === 0) {
						return literal ?? knownNumber(arg);
					}
					const toks = arg.filter((tok) => tok.kind !== 'comment');
					return toks.length === 1 ? values.get(tokenName(toks[0])?.toLowerCase() ?? '') ?? knownNumber(arg) : undefined;
				};
				checkSpan(source, stmt.span, host, model, memberCtx, env, arrays, sourceNames, valueOf, report, stringOf);
				if (workbook) {
					checkWorkbookSheetAccess(source, stmt.span, workbook, sourceNames, valueOf, report);
				}
			}, push);
		};
	};
}

/** What a statement knows of the active sheet: the sheet variables known not to be it, and a With's subject. */
interface ActiveSheetState {
	notActive: ReadonlySet<string>;
	subject?: string;
}

/**
 * The sheet variables each statement knows are not the active sheet (issue
 * #470, measured in Excel 16.0): `Set w2 = Worksheets.Add` activates the new
 * sheet, so every sheet the code held before is not active, and `w2` is.
 * `Set w1 = ActiveSheet` and `w1.Activate` make w1 the active one. Any
 * other statement that may activate a sheet, a call, a label, or a block
 * other than a With ends what is known.
 */
function activeSheetsAt(source: string, proc: ProcedureNode, activity: ConditionalActivityTracker | undefined): Map<BodyNode, ActiveSheetState> {
	const out = new Map<BodyNode, ActiveSheetState>();
	const sheets = new Set<string>();
	let notActive = new Set<string>();
	const visit = (list: readonly BodyNode[], subject: string | undefined): void => {
		for (const node of list) {
			if (activity?.isInactive(node.span)) {
				continue;
			}
			if (!isLeafStatement(node)) {
				const header = statementTokensAfterLeadingLabel(source, blockHeaderLineSpan(source, node.span)).filter((tok) => tok.kind !== 'comment');
				// The body runs from the state the block is entered with; after
				// it, what it may have activated is not known.
				if ('body' in node && Array.isArray(node.body)) {
					const entry = notActive;
					notActive = new Set(entry);
					visit(node.body as BodyNode[], node.kind === 'WithBlock' && header.length === 2 ? tokenName(header[1])?.toLowerCase() : subject);
				}
				notActive = new Set();
				continue;
			}
			if (jumpTargetLabelDeclaration(source, node.span)) {
				notActive = new Set();
			}
			if (notActive.size > 0) {
				out.set(node, { notActive: new Set(notActive), subject });
			}
			const toks = statementTokensAfterLeadingLabel(source, node.span).filter((tok) => tok.kind !== 'comment');
			const words = toks.map((tok) => tok.rawText.toLowerCase());
			if (words[0] === 'set' && words[2] === '=') {
				const target = words[1];
				const value = words.slice(3).join('');
				if (value === 'activesheet') {
					sheets.add(target);
					notActive.delete(target);
					continue;
				}
				// `Worksheets.Add`, `Sheets.Add(...)`, qualified by a workbook or not.
				if (/^(?:\w+\.)*(?:worksheets|sheets)\.add(?:\(.*\))?$/.test(value)) {
					notActive = new Set([...sheets].filter((sheet) => sheet !== target));
					sheets.add(target);
					continue;
				}
				if (!/[(]/.test(value) || /^(?:\w+\.)*(?:worksheets|sheets)\(/.test(value)) {
					notActive.delete(target);
					continue;
				}
			}
			// A plain assignment with no call keeps what is known.
			const bare = bareAssignmentTarget(source, node.span);
			const calls = toks.some((tok, i) => toks[i + 1]?.rawText === '(' && tok.kind === 'identifier' && toks[i - 1]?.rawText === '.' && ['activate', 'select', 'add', 'copy', 'move'].includes(tokenText(tok)));
			if ((bare || (words[0] === 'set' && words[2] === '=')) && !calls && !words.some((word) => word === 'activate' || word === 'select')) {
				continue;
			}
			// `w1.Activate`: w1 is active, and nothing else is known.
			notActive = new Set();
		}
	};
	visit(proc.body, undefined);
	return out;
}

/**
 * `w1.Range(Cells(1, 1), Cells(2, 2))` with w1 known not to be the active
 * sheet: the unqualified Cells, Range, Rows or Columns are the active
 * sheet's, so the Range raises 1004 (issue #470, measured in Excel 16.0).
 */
function checkUnqualifiedCorners(span: Span, toks: readonly VbaToken[], state: ActiveSheetState, push: PushFn): void {
	for (let i = 0; i + 3 < toks.length; i++) {
		const dotted = toks[i].rawText === '.' && tokenText(toks[i + 1]) === 'range' && toks[i + 2]?.rawText === '(';
		if (!dotted) {
			continue;
		}
		const owner = toks[i - 1] && toks[i - 1].rawText !== ')' ? tokenName(toks[i - 1])?.toLowerCase() : undefined;
		const sheet = owner ?? (i === 0 || ['=', '(', ','].includes(toks[i - 1]?.rawText ?? '') ? state.subject : undefined);
		if (!sheet || !state.notActive.has(sheet)) {
			continue;
		}
		const close = matchParenFrom(toks, i + 2);
		const args = close > i + 3 ? splitTopLevelTokenGroups(toks, i + 3, ',', close) : [];
		if (args.length !== 2) {
			continue;
		}
		for (const arg of args) {
			const part = arg.filter((tok) => tok.kind !== 'comment');
			const word = tokenText(part[0]);
			if (['cells', 'range', 'rows', 'columns'].includes(word) && part[1]?.rawText === '(' && matchParenFrom(part, 1) === part.length - 1) {
				const shown = part.map((tok) => tok.rawText).join('');
				push('hostArgumentOutOfRange', `${shown} here is the active sheet's, and '${owner ? toks[i - 1].rawText : `.${toks[i + 1].rawText}`}' is on a sheet that is not active, so Range cannot span them. This will raise Run-time error '1004': Method 'Range' of object '_Worksheet' failed.`, { start: span.start + part[0].start, end: span.start + part[part.length - 1].end });
				break;
			}
		}
	}
}

/** The sheets each statement knows: new and still empty, and those a new one differs from. */
interface SheetFacts {
	/** Sheets from `Worksheets.Add` nothing has written to yet. */
	empty: ReadonlySet<string>;
	/** Pairs of sheet variables known to be different sheets, `a|b`. */
	distinct: ReadonlySet<string>;
}

/** Members that write to a sheet or its cells, read on the way to a value. */
const SHEET_EDITS: ReadonlySet<string> = new Set(['add', 'insert', 'paste', 'pastespecial', 'copy', 'autofill', 'fill', 'filldown', 'fillright', 'formula', 'formular1c1', 'value', 'value2', 'text', 'clear', 'clearcontents', 'delete', 'sort', 'autofilter', 'texttocolumns', 'removeduplicates']);

/**
 * What each statement knows of the sheets the code just added (issue #472,
 * measured in Excel 16.0): `Set w2 = Worksheets.Add` gives an empty sheet,
 * different from every sheet the code held before. A statement that may
 * write to it, a call, a label, or a block ends what is known.
 */
function sheetFactsAt(source: string, proc: ProcedureNode, activity: ConditionalActivityTracker | undefined): Map<BodyNode, SheetFacts> {
	const out = new Map<BodyNode, SheetFacts>();
	const held = new Set<string>();
	let empty = new Set<string>();
	let distinct = new Set<string>();
	const visit = (list: readonly BodyNode[]): void => {
		for (const node of list) {
			if (activity?.isInactive(node.span)) {
				continue;
			}
			if (!isLeafStatement(node)) {
				if ('body' in node && Array.isArray(node.body)) {
					const [savedEmpty, savedDistinct] = [empty, distinct];
					[empty, distinct] = [new Set(savedEmpty), new Set(savedDistinct)];
					visit(node.body as BodyNode[]);
				}
				[empty, distinct] = [new Set(), new Set()];
				continue;
			}
			if (jumpTargetLabelDeclaration(source, node.span)) {
				[empty, distinct] = [new Set(), new Set()];
			}
			if (empty.size > 0 || distinct.size > 0) {
				out.set(node, { empty: new Set(empty), distinct: new Set(distinct) });
			}
			const toks = statementTokensAfterLeadingLabel(source, node.span).filter((tok) => tok.kind !== 'comment');
			const words = toks.map((tok) => tok.rawText.toLowerCase());
			if (words[0] === 'set' && words[2] === '=') {
				const target = words[1];
				const value = words.slice(3).join('');
				empty.delete(target);
				for (const pair of [...distinct]) {
					if (pair.split('|').includes(target)) {
						distinct.delete(pair);
					}
				}
				if (/^(?:\w+\.)*(?:worksheets|sheets)\.add(?:\(.*\))?$/.test(value)) {
					for (const sheet of held) {
						if (sheet !== target) {
							distinct.add([sheet, target].sort().join('|'));
						}
					}
					empty.add(target);
				}
				if (value === 'activesheet' || /^(?:\w+\.)*(?:worksheets|sheets)/.test(value)) {
					held.add(target);
				}
				continue;
			}
			// A read through the sheet into a variable keeps it; anything else may write.
			const bare = bareAssignmentTarget(source, node.span);
			const edits = toks.some((tok, i) => toks[i - 1]?.rawText === '.' && SHEET_EDITS.has(tokenText(tok)));
			const readsOnly = bare !== undefined && !empty.has(bare.name.toLowerCase()) && !edits;
			if (!readsOnly) {
				for (const sheet of [...empty]) {
					if (words.includes(sheet)) {
						empty.delete(sheet);
					}
				}
			}
		}
	};
	visit(proc.body);
	return out;
}

/** The SpecialCells types an empty sheet has no cells of. */
const EMPTY_SPECIAL_CELLS: ReadonlySet<string> = new Set(['xlcelltypeconstants', 'xlcelltypeformulas', 'xlcelltypeblanks', 'xlcelltypecomments', '2', '-4123', '4', '-4144']);

/**
 * Errors the literals and a new sheet prove (issue #472, measured in Excel
 * 16.0): Intersect of literal ranges that do not meet is Nothing (91 at its
 * member), Union across two sheets raises 1004, and on a sheet just added
 * ShowAllData, SpecialCells of constants, formulas, blanks or comments,
 * AutoFilter and TextToColumns raise 1004 and Find is Nothing (91).
 */
function checkSheetFacts(span: Span, toks: readonly VbaToken[], facts: SheetFacts | undefined, push: PushFn): void {
	const at = (from: number, to: number): Span => ({ start: span.start + toks[from].start, end: span.start + toks[to].end });
	for (let i = 0; i + 1 < toks.length; i++) {
		const word = tokenText(toks[i]);
		if ((word === 'intersect' || word === 'union') && toks[i + 1].rawText === '(' && toks[i - 1]?.rawText !== '.') {
			const close = matchParenFrom(toks, i + 1);
			const args = close > i + 2 ? splitTopLevelTokenGroups(toks, i + 2, ',', close) : [];
			const areas = args.map((arg) => sheetLiteralRange(arg));
			if (args.length !== 2 || areas.some((area) => !area)) {
				continue;
			}
			const [a, b] = areas as Array<{ sheet: string; area: A1Area }>;
			if (word === 'union' && a.sheet !== b.sheet && facts?.distinct.has([a.sheet, b.sheet].sort().join('|'))) {
				push('hostArgumentOutOfRange', `Union takes ranges of one sheet, and '${a.sheet}' and '${b.sheet}' are different sheets. This will raise Run-time error '1004': Method 'Union' of object '_Global' failed.`, at(i, close));
			} else if (word === 'intersect' && a.sheet === b.sheet && toks[close + 1]?.rawText === '.' && !overlaps(a.area, b.area)) {
				push('hostArgumentOutOfRange', `${toks.slice(i + 2, close).map((tok) => tok.rawText).join('')} do not meet, so Intersect is Nothing and has no '.${toks[close + 2]?.rawText ?? ''}'. This will raise Run-time error '91': Object variable or With block variable not set.`, at(i, close));
			}
			continue;
		}
		const sheet = tokenName(toks[i])?.toLowerCase();
		if (!sheet || !facts?.empty.has(sheet) || toks[i - 1]?.rawText === '.' || toks[i + 1].rawText !== '.') {
			continue;
		}
		// The chain from the sheet: `w2.Cells.SpecialCells(...)`, `w2.ShowAllData`.
		for (let k = i + 2; k < toks.length; k++) {
			const member = tokenText(toks[k]);
			const open = toks[k + 1]?.rawText === '(' ? k + 1 : -1;
			const close = open > 0 ? matchParenFrom(toks, open) : k;
			// `w2.Range("A1:B5").AutoFilter Field:=1`: a call statement's arguments.
			const bareCall = open < 0 && i === 0 && k + 1 < toks.length && toks[k + 1].rawText !== '.' && toks[k + 1].rawText !== '=';
			const args = open > 0 && close > open + 1 ? splitTopLevelTokenGroups(toks, open + 1, ',', close)
				: bareCall ? splitTopLevelTokenGroups(toks, k + 1, ',', toks.length) : [];
			const first = args[0]?.filter((tok) => tok.kind !== 'comment').map((tok) => tok.rawText.toLowerCase()).join('');
			let problem: string | undefined;
			if (member === 'showalldata') {
				problem = `'${toks[i].rawText}' is a sheet the code just added, with no filter to show. This will raise Run-time error '1004': Method 'ShowAllData' of object '_Worksheet' failed`;
			} else if (member === 'specialcells' && first !== undefined && EMPTY_SPECIAL_CELLS.has(first)) {
				problem = `'${toks[i].rawText}' is a sheet the code just added, which has no such cells. This will raise Run-time error '1004': No cells were found.`;
			} else if ((member === 'autofilter' || member === 'texttocolumns') && args.length > 0) {
				const why = member === 'autofilter' ? "This can't be applied to the selected range" : 'No data was selected to parse';
				problem = `'${toks[i].rawText}' is a sheet the code just added, and ${toks[k].rawText} has no data to act on. This will raise Run-time error '1004': ${why}`;
			} else if (member === 'find' && toks[close + 1]?.rawText === '.' && args[0]?.length === 1 && args[0][0].kind === 'stringLiteral' && args[0][0].rawText !== '""') {
				problem = `'${toks[i].rawText}' is a sheet the code just added, so Find finds nothing and returns Nothing, which has no '.${toks[close + 2]?.rawText ?? ''}'. This will raise Run-time error '91': Object variable or With block variable not set`;
			}
			if (problem) {
				push('hostArgumentOutOfRange', `${problem}.`, at(k, close));
				break;
			}
			if (toks[close + 1]?.rawText !== '.') {
				break;
			}
			k = close + 1;
		}
	}
}

/** `w2.Range("A1:B2")` or `Range("A1")`: the sheet variable (or "" for none) and the area. */
function sheetLiteralRange(arg: readonly VbaToken[]): { sheet: string; area: A1Area } | undefined {
	const toks = arg.filter((tok) => tok.kind !== 'comment');
	const at = toks.length === 6 && toks[1].rawText === '.' ? 2 : toks.length === 4 ? 0 : -1;
	if (at < 0 || tokenText(toks[at]) !== 'range' || toks[at + 1]?.rawText !== '(' || toks[at + 2]?.kind !== 'stringLiteral' || toks[at + 3]?.rawText !== ')') {
		return undefined;
	}
	const area = parseA1Address(stringLiteralValue(toks[at + 2].rawText));
	return area?.valid && area.row !== undefined && area.column !== undefined ? { sheet: at === 2 ? tokenName(toks[0])?.toLowerCase() ?? '' : '', area } : undefined;
}

/** Whether two A1 areas share a cell. */
function overlaps(a: A1Area, b: A1Area): boolean {
	const box = (area: A1Area): [number, number, number, number] => [
		Math.min(area.row!, area.endRow ?? area.row!), Math.max(area.row!, area.endRow ?? area.row!),
		Math.min(area.column!, area.endColumn ?? area.column!), Math.max(area.column!, area.endColumn ?? area.column!),
	];
	const [ar1, ar2, ac1, ac2] = box(a);
	const [br1, br2, bc1, bc2] = box(b);
	return ar1 <= br2 && br1 <= ar2 && ac1 <= bc2 && bc1 <= ac2;
}

/** Range members a call statement edits cells with. */
const CELL_EDITS: ReadonlySet<string> = new Set(['clearcontents', 'clear', 'clearformats', 'insert', 'delete', 'paste', 'pastespecial', 'autofill', 'filldown', 'fillright', 'merge', 'unmerge', 'sort']);

/** The members of a sheet that reach its cells. */
const CELL_PATHS: ReadonlySet<string> = new Set(['range', 'cells', 'rows', 'columns', 'usedrange']);

/**
 * The sheets the code just protected, with their password, and the faults
 * that follow (issue #471, measured in Excel 16.0): after `w2.Protect`, a
 * write to its cells or a cell edit raises 1004, and `w2.Unprotect` with
 * another password raises 1004. `Protect UserInterfaceOnly:=True` lets the
 * code write, a right Unprotect ends it, and a cell's Locked set by the
 * code, a call, a label or the end of a block ends what is known.
 */
function checkProtectedSheets(source: string, proc: ProcedureNode, activity: ConditionalActivityTracker | undefined, push: PushFn): void {
	let protectedSheets = new Map<string, string>();
	// Sheets the code unlocked a cell on: which cells stay writable is not followed.
	const unlocked = new Set<string>();
	const visit = (list: readonly BodyNode[]): void => {
		for (const node of list) {
			if (activity?.isInactive(node.span)) {
				continue;
			}
			if (!isLeafStatement(node)) {
				if ('body' in node && Array.isArray(node.body)) {
					const entry = protectedSheets;
					protectedSheets = new Map(entry);
					visit(node.body as BodyNode[]);
				}
				protectedSheets = new Map();
				continue;
			}
			if (jumpTargetLabelDeclaration(source, node.span)) {
				protectedSheets = new Map();
			}
			const toks = statementTokensAfterLeadingLabel(source, node.span).filter((tok) => tok.kind !== 'comment');
			const words = toks.map((tok) => tok.rawText.toLowerCase());
			const sheet = words[0];
			const at = (from: number, to: number): Span => ({ start: node.span.start + toks[from].start, end: node.span.start + toks[to].end });
			if (words.includes('locked')) {
				unlocked.add(sheet);
			}
			if (words[1] === '.' && (words[2] === 'protect' || words[2] === 'unprotect')) {
				const args = toks.length > 3 ? splitTopLevelTokenGroups(toks, toks[3].rawText === '(' ? 4 : 3, ',', toks[3].rawText === '(' ? matchParenFrom(toks, 3) : toks.length) : [];
				const named = (name: string): VbaToken[] | undefined => args.find((arg) => tokenText(arg[0]) === name && arg[1]?.rawText === ':=')?.slice(2);
				const passwordArg = named('password') ?? (args[0] && args[0][1]?.rawText !== ':=' ? args[0] : undefined);
				const password = passwordArg === undefined ? '' : passwordArg.length === 1 && passwordArg[0].kind === 'stringLiteral' ? stringLiteralValue(passwordArg[0].rawText) : undefined;
				if (words[2] === 'protect') {
					const uiOnly = named('userinterfaceonly');
					if (password === undefined || unlocked.has(sheet) || (uiOnly && tokenText(uiOnly[0]) !== 'false')) {
						protectedSheets.delete(sheet);
					} else {
						protectedSheets.set(sheet, password);
					}
				} else {
					const held = protectedSheets.get(sheet);
					if (held !== undefined && password !== undefined && password !== held && passwordArg) {
						push('hostArgumentOutOfRange', `'${toks[0].rawText}' was protected with another password, which Unprotect must match. This will raise Run-time error '1004': The password you supplied is not correct.`, at(2, toks.length - 1));
						continue;
					}
					protectedSheets.delete(sheet);
				}
				continue;
			}
			if (protectedSheets.has(sheet) && words[1] === '.' && CELL_PATHS.has(words[2])) {
				const eq = toks.findIndex((tok, i) => tok.rawText === '=' && i > 2);
				const edit = toks.findIndex((tok, i) => i > 2 && toks[i - 1]?.rawText === '.' && CELL_EDITS.has(tokenText(tok)));
				const locked = words.includes('locked');
				if (locked) {
					protectedSheets.delete(sheet);
					continue;
				}
				if (eq > 0 || edit > 0) {
					push('hostArgumentOutOfRange', `'${toks[0].rawText}' is protected here, so its cells cannot be changed. This will raise Run-time error '1004': The cell or chart you're trying to change is on a protected sheet.`, at(0, (eq > 0 ? eq : edit + 1) - 1));
				}
				continue;
			}
			// A read keeps what is known; a call or another use of a sheet may unprotect it.
			const bare = bareAssignmentTarget(source, node.span);
			if (!bare || toks.some((tok) => ['unprotect', 'locked'].includes(tokenText(tok)))) {
				if (!(bare || words[0] === 'set') || toks.some((tok) => tokenText(tok) === 'unprotect')) {
					protectedSheets = new Map();
				}
			}
		}
	};
	visit(proc.body);
}

/** Members that add to or edit a document, read on the way to a value. */
const DOCUMENT_EDITS: ReadonlySet<string> = new Set([
	'add', 'addfield', 'addpicture', 'addtable', 'insertafter', 'insertbefore', 'insertparagraph', 'insertparagraphafter',
	'insertparagraphbefore', 'insertbreak', 'insertfile', 'paste', 'delete', 'cut', 'converttotable', 'typetext',
]);

/** A Word document the procedure just added, and the text it wrote into it, if any. */
interface NewDocument {
	/** The whole text the code set through Content.Text, or "" for an untouched document. */
	text: string;
}

/**
 * The documents each statement sees as new (issue #497, measured in Word
 * 16.0): `Set d = Documents.Add` gives an empty document, and
 * `d.Content.Text = "..."` sets its whole text. A statement that names d
 * other than to read through it, a label, or a block that names it ends
 * what is known.
 */
function newDocumentsAt(source: string, proc: ProcedureNode, activity: ConditionalActivityTracker | undefined): Map<BodyNode, ReadonlyMap<string, NewDocument>> {
	const out = new Map<BodyNode, ReadonlyMap<string, NewDocument>>();
	const visit = (list: readonly BodyNode[], state: Map<string, NewDocument>): void => {
		for (const node of list) {
			if (activity?.isInactive(node.span)) {
				continue;
			}
			if (!isLeafStatement(node)) {
				if ('body' in node && Array.isArray(node.body)) {
					visit(node.body as BodyNode[], new Map(state));
					const named = new Set(statementTokens(source, node.span).map((tok) => tokenName(tok)?.toLowerCase()));
					for (const lower of [...state.keys()]) {
						if (named.has(lower)) {
							state.delete(lower);
						}
					}
				}
				continue;
			}
			if (jumpTargetLabelDeclaration(source, node.span)) {
				state.clear();
			}
			if (state.size > 0) {
				out.set(node, new Map(state));
			}
			const toks = statementTokensAfterLeadingLabel(source, node.span).filter((tok) => tok.kind !== 'comment');
			const words = toks.map((tok) => tok.rawText.toLowerCase());
			// `Set d = Documents.Add` or `Set d = Documents.Add()`, plain.
			const from = words[0] === 'set' && words[2] === '=' ? (words[3] === 'application' && words[4] === '.' ? 5 : 3) : -1;
			if (from > 0 && words[from] === 'documents' && words[from + 1] === '.' && words[from + 2] === 'add'
				&& (toks.length === from + 3 || (toks.length === from + 5 && words[from + 3] === '(' && words[from + 4] === ')'))) {
				state.set(words[1], { text: '' });
				continue;
			}
			// `d.Content.Text = "..."` writes the whole text.
			const target = words[0];
			if (state.has(target) && toks.length === 7 && words[1] === '.' && words[2] === 'content' && words[3] === '.' && words[4] === 'text' && words[5] === '=' && toks[6].kind === 'stringLiteral') {
				state.set(target, { text: stringLiteralValue(toks[6].rawText) });
				continue;
			}
			// Anything but a read through the document may change it, and so
			// may a method that adds or edits on the way: `x = d.Tables.Add(...)`.
			const eq = toks.findIndex((tok) => tok.rawText === '=');
			const edits = toks.some((tok, i) => toks[i - 1]?.rawText === '.' && DOCUMENT_EDITS.has(tokenText(tok)));
			const reads = !edits && bareAssignmentTarget(source, node.span) !== undefined && !state.has(target);
			for (const lower of [...state.keys()]) {
				if (words.includes(lower) && !(reads && toks.every((tok, i) => i <= eq || tokenName(tok)?.toLowerCase() !== lower || toks[i + 1]?.rawText === '.'))) {
					state.delete(lower);
				}
			}
		}
	};
	visit(proc.body, new Map());
	return out;
}

/** What a new document holds of each collection, counted from its text. */
function newDocumentCount(document: NewDocument, member: string): number | undefined {
	const characters = document.text.length + 1; // the final paragraph mark
	switch (member) {
		case 'tables':
		case 'fields':
		case 'inlineshapes':
		case 'bookmarks':
			return 0;
		case 'sections':
		case 'paragraphs':
			return 1;
		case 'sentences':
			// Each sentence ends at a stop or at the paragraph's end.
			return (document.text.match(/[.!?]/g) ?? []).length + 1;
		case 'words':
		case 'characters':
			return characters;
	}
	return undefined;
}

/** Members of a new document past what it holds: `d.Tables(1)`, `d.Words(50)`, `d.Range(0, 99999)`. */
function checkNewDocumentUses(
	span: Span,
	toks: readonly VbaToken[],
	documents: ReadonlyMap<string, NewDocument>,
	valueOf: (arg: readonly VbaToken[]) => number | undefined,
	push: PushFn,
): void {
	for (let i = 0; i + 4 < toks.length; i++) {
		const name = tokenName(toks[i])?.toLowerCase();
		const document = name ? documents.get(name) : undefined;
		if (!document || toks[i - 1]?.rawText === '.' || toks[i + 1].rawText !== '.' || toks[i + 3].rawText !== '(') {
			continue;
		}
		const member = tokenText(toks[i + 2]);
		const close = matchParenFrom(toks, i + 3);
		const args = close > i + 4 ? splitTopLevelTokenGroups(toks, i + 4, ',', close) : [];
		const at = { start: span.start + toks[i + 2].start, end: span.start + toks[close].end };
		const what = document.text ? `whose text the code set to ${document.text.length} character(s)` : 'which the code just added';
		if (member === 'range' && args.length === 2) {
			const end = valueOf(args[1]);
			const characters = document.text.length + 1;
			if (end !== undefined && end > characters) {
				push('hostArgumentOutOfRange', `'${toks[i].rawText}' is a new document ${what}, so it ends at position ${characters}, and Range ends at ${end}. This will raise Run-time error '4608': Value out of range.`, at);
			}
			continue;
		}
		const count = newDocumentCount(document, member);
		if (count === undefined || args.length !== 1) {
			continue;
		}
		const named = member === 'bookmarks' && args[0].length === 1 && args[0][0].kind === 'stringLiteral';
		const index = named ? undefined : valueOf(args[0]);
		if (named || (index !== undefined && index > count)) {
			const shown = toks.slice(i + 2, close + 1).map((tok) => tok.rawText).join('');
			push('hostArgumentOutOfRange', `'${toks[i].rawText}' is a new document ${what}, which has ${count} ${toks[i + 2].rawText}, so ${shown} does not exist. This will raise Run-time error '5941': The requested member of the collection does not exist.`, at);
		}
	}
}

function checkSpan(
	source: string,
	span: Span,
	host: 'Excel' | 'Word' | 'PowerPoint',
	model: HostObjectModel | undefined,
	memberCtx: MemberCompletionContext,
	env: ReadonlyMap<string, string>,
	arrays: ReadonlySet<string>,
	sourceNames: ReadonlySet<string>,
	valueOf: (arg: readonly VbaToken[]) => number | undefined,
	push: PushFn,
	stringOf: (arg: readonly VbaToken[]) => string | undefined = literalString,
): void {
	const toks = statementTokens(source, span);
	const at = (from: number, to: number): Span => ({ start: span.start + toks[from].start, end: span.start + toks[to].end });
	if (host === 'Excel') {
		checkSheetNameAssignment(source, span, toks, memberCtx, sourceNames, push);
		// `Range("A1:B2").Areas(2)`: an address with no comma is one area
		// (issue #278, measured in Excel 16.0: 1004).
		for (let i = 0; i + 8 < toks.length; i++) {
			if (tokenText(toks[i]) !== 'range' || (sourceNames.has('range') && toks[i - 1]?.rawText !== '.') || toks[i + 1].rawText !== '(' || toks[i + 2].kind !== 'stringLiteral' || toks[i + 3].rawText !== ')' || toks[i + 4].rawText !== '.'
				|| tokenText(toks[i + 5]) !== 'areas' || toks[i + 6].rawText !== '(' || toks[i + 7].kind !== 'integerLiteral' || toks[i + 8].rawText !== ')') {
				continue;
			}
			const address = toks[i + 2].rawText.slice(1, -1);
			const index = Number(toks[i + 7].rawText.replace(/[%&^]$/, ''));
			if (/^\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?$/.test(address) && index > 1) {
				push('hostArgumentOutOfRange', `Range("${address}") is one area, so Areas(${index}) names none. This will raise Run-time error '1004': Application-defined or object-defined error.`, at(i + 7, i + 7));
			}
		}
		checkBeforeAndAfter(source, span, toks, memberCtx, push);
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
		const relative = host === 'Excel' && callee.receiver === 'Excel.Range' && RANGE_RELATIVE_MEMBERS.has(lower);
		if (collection && callee.args.length === 1 && !RANGE_COORDINATE_MEMBERS.has(lower) && !relative) {
			const index = valueOf(callee.args[0]);
			if (index !== undefined && index < 1) {
				const error = collectionIndexError(host, collection, model);
				push(
					'hostArgumentOutOfRange',
					`Index ${index} is never an element: ${host} collections start at 1. This will raise Run-time error '${error.number}': ${error.text}.`,
					at(callee.openIndex + 1, callee.closeIndex - 1),
				);
				continue;
			}
			// `Worksheets(Worksheets.Count + 1)`: past the last element (issue
			// #309, measured in Excel and Word 16.0).
			const chain = chainText(toks, receiverStart(toks, callee.nameIndex), lower === 'item' ? callee.nameIndex - 2 : callee.nameIndex);
			const past = chain ? countOffset(callee.args[0], chain) : undefined;
			if (past !== undefined && past >= 1) {
				const error = collectionIndexError(host, collection, model);
				push(
					'hostArgumentOutOfRange',
					`${chain}.Count + ${past} is past the last element of ${chain}. This will raise Run-time error '${error.number}': ${error.text}.`,
					at(callee.openIndex + 1, callee.closeIndex - 1),
				);
				continue;
			}
		}
		checkArgumentLimits(span, callee, valueOf, push);
		checkInsertionPosition(span, toks, callee, valueOf, push);
		if (host === 'Word') {
			checkWordNames(span, callee, stringOf, push);
		}
		if (host === 'Excel') {
			checkExcelMethodArguments(span, toks, callee, stringOf, push);
			checkExcelCallee(source, span, toks, callee, calleeSpan, env, arrays, sourceNames, valueOf, push, stringOf);
		} else if (host === 'Word') {
			if (lower === 'range' && callee.receiver === 'Word.Document' && callee.openIndex > 0) {
				const start = callee.args[0] ? valueOf(callee.args[0]) : undefined;
				const end = callee.args[1] ? valueOf(callee.args[1]) : undefined;
				const bad = (start !== undefined && start < 0) || (end !== undefined && end < 0) || (start !== undefined && end !== undefined && end < start);
				if (bad) {
					push(
						'hostArgumentOutOfRange',
						`Document.Range takes character positions from 0, with End at or after Start. This will raise Run-time error '4608': Value out of range.`,
						at(callee.openIndex + 1, callee.closeIndex - 1),
					);
				}
			}
		}
	}
}

/**
 * Counts a host method refuses (issue #204, measured in Excel and Word 16.0):
 * `Worksheets.Add Count:=0` raises 1004, and Word's Tables.Add takes 1 to
 * 32767 rows and 1 to 63 columns (5148). Each argument is found by name or
 * by position.
 */
const ARGUMENT_LIMITS: ReadonlyArray<{
	receivers: readonly string[];
	member: string;
	parameter: string;
	position: number;
	runs: string;
	refused: ReadonlyArray<{ from?: number; to?: number }>;
	error: { number: string; text: string };
}> = [
	{
		receivers: ['Excel.Sheets', 'Excel.Worksheets'], member: 'add', parameter: 'Count', position: 2, runs: '1 or more',
		refused: [{ to: 0 }], error: { number: '1004', text: "Method 'Add' of object 'Sheets' failed" },
	},
	// Enum arguments (issue #244), swept from -100 to 999 in Excel 16.0 as
	// hostPropertyValues.ts's enum properties were.
	{
		receivers: ['Excel.Range'], member: 'borders', parameter: 'Index', position: 0, runs: '1 to 12, or an xlBordersIndex constant',
		refused: [{ from: -100, to: 0 }, { from: 13, to: 999 }], error: { number: '1004', text: 'Unable to get the Item property of the Borders class' },
	},
	{
		receivers: ['Excel.Range'], member: 'end', parameter: 'Direction', position: 0, runs: '1 to 4, or an xlDirection constant',
		refused: [{ from: -100, to: 0 }, { from: 5, to: 999 }], error: { number: '1004', text: 'Application-defined or object-defined error' },
	},
	{
		receivers: ['Excel.Range'], member: 'specialcells', parameter: 'Type', position: 0, runs: 'an xlCellType constant',
		refused: [{ from: -100, to: 0 }, { from: 13, to: 13 }, { from: 17, to: 999 }], error: { number: '1004', text: 'Application-defined or object-defined error' },
	},
	{
		receivers: ['Excel.Range'], member: 'sort', parameter: 'Order1', position: 1, runs: 'xlAscending (1) or xlDescending (2)',
		refused: [{ from: -100, to: 0 }, { from: 3, to: 999 }], error: { number: '1004', text: 'Application-defined or object-defined error' },
	},
	{
		receivers: ['Excel.Range'], member: 'pastespecial', parameter: 'Paste', position: 0, runs: 'an xlPasteType constant',
		refused: [{ from: -100, to: 0 }, { from: 9, to: 10 }, { from: 15, to: 999 }], error: { number: '1004', text: 'Application-defined or object-defined error' },
	},
	{
		receivers: ['Excel.Range'], member: 'insert', parameter: 'Shift', position: 0, runs: '1 to 4, or an xlInsertShiftDirection constant',
		refused: [{ from: -100, to: 0 }, { from: 5, to: 999 }], error: { number: '1004', text: 'Application-defined or object-defined error' },
	},
	{
		receivers: ['Excel.Range'], member: 'delete', parameter: 'Shift', position: 0, runs: '1 to 4, or an xlDeleteShiftDirection constant',
		refused: [{ from: -100, to: 0 }, { from: 5, to: 999 }], error: { number: '1004', text: 'Application-defined or object-defined error' },
	},
	{
		receivers: ['Excel.Range'], member: 'autofill', parameter: 'Type', position: 1, runs: '0 to 12, or an xlAutoFillType constant',
		refused: [{ from: -100, to: -1 }, { from: 13, to: 999 }], error: { number: '1004', text: 'Application-defined or object-defined error' },
	},
	// Word and PowerPoint (issue #245), swept from -100 to 999 in 16.0.
	{
		receivers: ['Word.Selection'], member: 'moveright', parameter: 'Unit', position: 0, runs: '1 to 3 or 16: wdCharacter, wdWord, wdSentence or wdCell',
		refused: [{ from: -100, to: 0 }, { from: 4, to: 11 }, { from: 13, to: 15 }, { from: 17, to: 999 }], error: { number: '4120', text: 'Bad parameter' },
	},
	{
		receivers: ['Word.Selection'], member: 'collapse', parameter: 'Direction', position: 0, runs: 'wdCollapseEnd (0) or wdCollapseStart (1)',
		refused: [{ from: -100, to: -1 }, { from: 2, to: 999 }], error: { number: '4120', text: 'Bad parameter' },
	},
	{
		receivers: ['Word.Selection'], member: 'insertbreak', parameter: 'Type', position: 0, runs: '0 to 11, a wdBreakType constant',
		refused: [{ from: -100, to: -1 }, { from: 12, to: 999 }], error: { number: '9118', text: 'Parameter value was out of acceptable range' },
	},
	{
		receivers: ['PowerPoint.Shapes'], member: 'addshape', parameter: 'Type', position: 0, runs: '1 to 183, an msoAutoShapeType constant',
		refused: [{ from: -100, to: 0 }, { from: 184, to: 999 }], error: { number: '-2147024809', text: 'The specified value is out of range' },
	},
	{
		receivers: ['PowerPoint.Slides'], member: 'add', parameter: 'Layout', position: 1, runs: 'a ppSlideLayout constant from 1',
		refused: [{ from: -100, to: 0 }, { from: 37, to: 999 }], error: { number: '-2147024809', text: 'Invalid enumeration value' },
	},
	{
		receivers: ['PowerPoint.Shapes'], member: 'addtable', parameter: 'NumRows', position: 0, runs: '1 to 75',
		refused: [{ from: -100, to: 0 }, { from: 76, to: 999 }], error: { number: '-2147188160', text: 'Integer out of range' },
	},
	{
		receivers: ['PowerPoint.Shapes'], member: 'addtable', parameter: 'NumColumns', position: 1, runs: '1 to 75',
		refused: [{ from: -100, to: 0 }, { from: 76, to: 999 }], error: { number: '-2147188160', text: 'Integer out of range' },
	},
	{
		receivers: ['PowerPoint.Slides'], member: 'range', parameter: 'Index', position: 0, runs: '1 or more',
		refused: [{ from: -100, to: 0 }], error: { number: '-2147188160', text: 'Invalid request' },
	},
	// Issue #311, measured in PowerPoint 16.0.
	{
		receivers: ['PowerPoint.Shapes'], member: 'addtextbox', parameter: 'Width', position: 3, runs: '0 or more',
		refused: [{ to: -1 }], error: { number: '-2147024809', text: 'The specified value is out of range' },
	},
	{
		receivers: ['PowerPoint.Shapes'], member: 'addtextbox', parameter: 'Height', position: 4, runs: '0 or more',
		refused: [{ to: -1 }], error: { number: '-2147024809', text: 'The specified value is out of range' },
	},
	{
		receivers: ['PowerPoint.Slide', 'PowerPoint.SlideRange'], member: 'moveto', parameter: 'ToPos', position: 0, runs: '1 or more',
		refused: [{ to: 0 }], error: { number: '-2147188160', text: 'Integer out of range' },
	},
	// Issue #610, measured in PowerPoint 16.0: orientations 1 and 6 run, 0,
	// -2 (mixed), 7 and 9 do not; a width or height of 0 runs.
	{
		receivers: ['PowerPoint.Shapes'], member: 'addtextbox', parameter: 'Orientation', position: 0, runs: '1 to 6, an msoTextOrientation constant',
		refused: [{ to: 0 }, { from: 7 }], error: { number: '-2147024809', text: 'The specified value is out of range' },
	},
	{
		receivers: ['PowerPoint.Shapes'], member: 'addshape', parameter: 'Width', position: 3, runs: '0 or more',
		refused: [{ to: -1 }], error: { number: '-2147024809', text: 'The specified value is out of range' },
	},
	{
		receivers: ['PowerPoint.Shapes'], member: 'addshape', parameter: 'Height', position: 4, runs: '0 or more',
		refused: [{ to: -1 }], error: { number: '-2147024809', text: 'The specified value is out of range' },
	},
	{
		receivers: ['Word.Tables'], member: 'add', parameter: 'NumRows', position: 1, runs: '1 to 32767',
		refused: [{ to: 0 }, { from: 32768 }], error: { number: '5148', text: 'The number must be between 1 and 32767' },
	},
	{
		receivers: ['Word.Tables'], member: 'add', parameter: 'NumColumns', position: 2, runs: '1 to 63',
		refused: [{ to: 0 }, { from: 64 }], error: { number: '5148', text: 'The number must be between 1 and 63' },
	},
];

function checkArgumentLimits(
	span: Span,
	callee: HostCallee,
	valueOf: (arg: readonly VbaToken[]) => number | undefined,
	push: PushFn,
): void {
	const lower = callee.name.toLowerCase();
	for (const limit of ARGUMENT_LIMITS) {
		if (limit.member !== lower || !limit.receivers.includes(callee.receiver)) {
			continue;
		}
		const arg = argumentByNameOrPosition(callee.args, limit.parameter, limit.position);
		const value = arg && arg.length > 0 ? valueOf(arg) : undefined;
		if (value === undefined || !limit.refused.some((range) => (range.from === undefined || value >= range.from) && (range.to === undefined || value <= range.to))) {
			continue;
		}
		push(
			'hostArgumentOutOfRange',
			`${callee.receiver.slice(callee.receiver.indexOf('.') + 1)}.${callee.name} takes ${limit.parameter} ${limit.runs}; ${value} is outside that. This will raise Run-time error '${limit.error.number}': ${limit.error.text}.`,
			argSpan(span, arg!),
		);
	}
}

/**
 * Names Word refuses (issue #311, measured in Word 16.0): a bookmark name
 * that is empty, starts with a digit, or holds anything but letters, digits
 * and underscores (5828), and an empty style name (5167). A bookmark name
 * past 40 letters is shortened, not refused. Issue #610 adds a style name of
 * blanks (5167) and a built-in style's name in any case (5173), and an empty
 * Variables name (-2147467259).
 */
function checkWordNames(
	span: Span,
	callee: HostCallee,
	stringOf: (arg: readonly VbaToken[]) => string | undefined,
	push: PushFn,
): void {
	if (callee.name.toLowerCase() !== 'add') {
		return;
	}
	const arg = argumentByNameOrPosition(callee.args, 'Name', 0);
	const name = arg?.length ? stringOf(arg) : undefined;
	if (name === undefined) {
		return;
	}
	if (callee.receiver === 'Word.Bookmarks') {
		const why = name === '' ? 'is empty' : /^\d/.test(name) ? 'starts with a digit' : /[^\p{L}\p{N}_]/u.test(name) ? 'holds a character other than a letter, a digit or _' : undefined;
		if (why) {
			push('hostArgumentOutOfRange', `The bookmark name "${name}" ${why}: a bookmark name starts with a letter and holds letters, digits and _. This will raise Run-time error '5828': Bad bookmark name.`, argSpan(span, arg!));
		}
	} else if (callee.receiver === 'Word.Styles' && name.trim() === '') {
		push('hostArgumentOutOfRange', `A style needs a name, and "${name}" is none. This will raise Run-time error '5167': This is not a valid style name.`, argSpan(span, arg!));
	} else if (callee.receiver === 'Word.Styles' && WORD_BUILTIN_STYLES.has(name.toLowerCase())) {
		push('hostArgumentOutOfRange', `"${name}" is the name of a built-in style, which a style the code adds cannot take, in any case. This will raise Run-time error '5173': This style name already exists or is reserved for a built-in style.`, argSpan(span, arg!));
	} else if (callee.receiver === 'Word.Variables' && name === '') {
		push('hostArgumentOutOfRange', `A document variable needs a name, and "" is none. This will raise Run-time error '-2147467259': Method 'Add' of object 'Variables' failed.`, argSpan(span, arg!));
	}
}

/**
 * Excel methods whose 1004 the arguments prove (issue #308, measured in
 * Excel 16.0): AutoFill into a range that does not hold its source, a Sort
 * of several cells keyed on a column outside them, and Names.Add of a name
 * Excel cannot hold.
 */
function checkExcelMethodArguments(
	span: Span,
	toks: readonly VbaToken[],
	callee: HostCallee,
	stringOf: (arg: readonly VbaToken[]) => string | undefined,
	push: PushFn,
): void {
	const lower = callee.name.toLowerCase();
	const given = (name: string, position: number): VbaToken[] | undefined => {
		const arg = argumentByNameOrPosition(callee.args, name, position);
		return arg && arg.length > 0 ? arg : undefined;
	};
	if (callee.receiver === 'Excel.Range' && lower === 'autofill') {
		const from = literalRangeReceiver(toks, callee.nameIndex - 1);
		const destination = given('Destination', 0);
		const to = destination ? literalRangeArgument(destination) : undefined;
		const holds = from && to && to.row <= from.row && to.column <= from.column && to.row + to.rows >= from.row + from.rows && to.column + to.width >= from.column + from.width;
		const same = holds && to.rows === from.rows && to.width === from.width;
		if (from && to && (!holds || same)) {
			push('hostArgumentOutOfRange', `AutoFill fills from ${from.text} into ${to.text}, which ${same ? 'is the source itself' : 'does not take it in'}: the destination must hold the source and reach past it. This will raise Run-time error '1004': AutoFill method of Range class failed.`, argSpan(span, destination!));
		}
		return;
	}
	if (callee.receiver === 'Excel.Range' && lower === 'sort' && !given('Orientation', 10)) {
		const block = literalRangeReceiver(toks, callee.nameIndex - 1);
		if (!block || (block.rows === 1 && block.width === 1)) {
			return; // one cell sorts its current region, which the data decides
		}
		for (const [name, position] of [['Key1', 0], ['Key2', 2], ['Key3', 5]] as const) {
			const key = given(name, position);
			const at = key ? literalRangeArgument(key) : undefined;
			if (at && (at.column + at.width <= block.column || at.column >= block.column + block.width)) {
				push('hostArgumentOutOfRange', `${name} ${at.text} lies outside the columns of ${block.text}, which is what is sorted. This will raise Run-time error '1004': The sort reference is not valid.`, argSpan(span, key!));
				return;
			}
		}
		return;
	}
	if (callee.receiver === 'Excel.Names' && lower === 'add') {
		const arg = given('Name', 0);
		const name = arg ? stringOf(arg) : undefined;
		const why = name === undefined ? undefined : refusedName(name);
		if (why) {
			push('hostArgumentOutOfRange', `"${name}" ${why}, so Excel holds no name of that spelling. This will raise Run-time error '1004': The syntax of this name isn't correct.`, argSpan(span, arg!));
		}
	}
}

/**
 * Why Excel refuses a name, where it is plain (issue #308, measured in
 * Excel 16.0): a space, a digit first, or the spelling of a cell, A1 or
 * R1C1. A0, XFE1 and A1048577 are names it accepts.
 */
function refusedName(name: string): string | undefined {
	if (/\s/.test(name)) {
		return 'holds a space';
	}
	if (/^\d/.test(name)) {
		return 'starts with a digit';
	}
	const cell = /^\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(name);
	if (cell && columnNumber(cell[1]) <= EXCEL_MAX_COLUMN && Number(cell[2]) >= 1 && Number(cell[2]) <= EXCEL_MAX_ROW) {
		return 'is the address of a cell';
	}
	if (/^[Rr]\d+[Cc]\d+$/.test(name)) {
		return 'is an R1C1 address';
	}
	return undefined;
}

/** `Range("B1:B3")` as a whole argument: its top-left cell and its size. */
function literalRangeArgument(arg: readonly VbaToken[]): { row: number; column: number; rows: number; width: number; text: string } | undefined {
	const toks = arg.filter((tok) => tok.kind !== 'comment');
	return toks.length === 4 ? literalRangeAt(toks, 0) : undefined;
}

const BEFORE_AND_AFTER_TYPES: ReadonlySet<string> = new Set(['Excel.Sheets', 'Excel.Worksheets', 'Excel.Charts', 'Excel.Worksheet', 'Excel.Chart']);

/**
 * `Worksheets.Add Before:=..., After:=...` and Move or Copy of a sheet
 * given both: the sheet goes before one or after one, never both (issue
 * #308, measured in Excel 16.0: 1004).
 */
function checkBeforeAndAfter(source: string, span: Span, toks: readonly VbaToken[], memberCtx: MemberCompletionContext, push: PushFn): void {
	for (let i = 1; i < toks.length; i++) {
		const lower = tokenText(toks[i]);
		if (toks[i - 1].rawText !== '.' || (lower !== 'add' && lower !== 'move' && lower !== 'copy')) {
			continue;
		}
		const open = toks[i + 1]?.rawText === '(' ? i + 1 : -1;
		const close = open > 0 ? matchParenFrom(toks, open) : toks.length;
		if (open < 0 && i !== firstExecutableTokenIndexOfMemberCall(toks, i)) {
			continue;
		}
		const args = close > (open > 0 ? open + 1 : i + 1) ? splitTopLevel(toks.slice(open > 0 ? open + 1 : i + 1, close)) : [];
		const before = argumentByNameOrPosition(args, 'Before', 0);
		const after = argumentByNameOrPosition(args, 'After', 1);
		if (!before?.length || !after?.length) {
			continue;
		}
		const parts = (resolveReceiverTypeAt(source, span.start + toks[i - 1].end, memberCtx) ?? '').replace(/^union:/, '').split('|');
		if (parts.every((part) => BEFORE_AND_AFTER_TYPES.has(part))) {
			push('hostArgumentOutOfRange', `${toks[i].rawText} takes Before or After, not both: a sheet goes before one sheet or after one. This will raise Run-time error '1004': Method '${toks[i].rawText}' failed.`, { start: span.start + toks[i].start, end: span.start + toks[close === toks.length ? close - 1 : close].end });
		}
	}
}

/**
 * Methods that insert at a position from 1 to Count + 1 (issue #309,
 * measured in Excel and PowerPoint 16.0): 0 and below, and Count + 2 or
 * more written against the collection's own Count, are refused.
 */
const INSERTIONS: ReadonlyArray<{ receiver: string; member: string; parameter: string; noun: string; error: { number: string; text: string } }> = [
	{ receiver: 'PowerPoint.Slides', member: 'add', parameter: 'Index', noun: 'slide', error: { number: '-2147188160', text: 'Integer out of range' } },
	{ receiver: 'PowerPoint.Slides', member: 'addslide', parameter: 'Index', noun: 'slide', error: { number: '-2147188160', text: 'Integer out of range' } },
	{ receiver: 'Excel.ListRows', member: 'add', parameter: 'Position', noun: 'row', error: { number: '9', text: 'Subscript out of range' } },
	{ receiver: 'Excel.ListColumns', member: 'add', parameter: 'Position', noun: 'column', error: { number: '9', text: 'Subscript out of range' } },
];

function checkInsertionPosition(
	span: Span,
	toks: readonly VbaToken[],
	callee: HostCallee,
	valueOf: (arg: readonly VbaToken[]) => number | undefined,
	push: PushFn,
): void {
	const lower = callee.name.toLowerCase();
	const insertion = INSERTIONS.find((entry) => entry.member === lower && entry.receiver === callee.receiver);
	const arg = insertion ? argumentByNameOrPosition(callee.args, insertion.parameter, 0) : undefined;
	if (!insertion || !arg?.length) {
		return;
	}
	const value = valueOf(arg);
	const chain = chainText(toks, receiverStart(toks, callee.nameIndex), callee.nameIndex - 2);
	const past = chain ? countOffset(arg, chain) : undefined;
	const shown = value !== undefined && value < 1 ? `${value}` : past !== undefined && past >= 2 ? `${chain}.Count + ${past}` : undefined;
	if (shown) {
		const type = callee.receiver.slice(callee.receiver.indexOf('.') + 1);
		push(
			'hostArgumentOutOfRange',
			`${type}.${callee.name} puts the new ${insertion.noun} at ${insertion.parameter}, from 1 to Count + 1; ${shown} is outside that. This will raise Run-time error '${insertion.error.number}': ${insertion.error.text}.`,
			argSpan(span, arg),
		);
	}
}

/** The text of `toks[from..to]` when it is names and dots only: `d.Paragraphs`. */
function chainText(toks: readonly VbaToken[], from: number, to: number): string | undefined {
	const part = toks.slice(from, to + 1);
	return part.length > 0 && part.every((tok, k) => (k % 2 === 0 ? tokenName(tok) !== undefined : tok.rawText === '.')) && part.length % 2 === 1
		? part.map((tok) => tok.rawText).join('')
		: undefined;
}

/** `k` for an argument written `<chain>.Count + k` (0 for `<chain>.Count`), or undefined. */
function countOffset(arg: readonly VbaToken[], chain: string): number | undefined {
	const toks = arg.filter((tok) => tok.kind !== 'comment');
	const count = toks.findIndex((tok, k) => tokenText(tok) === 'count' && toks[k - 1]?.rawText === '.');
	if (count < 2 || chainText(toks, 0, count - 2)?.toLowerCase() !== chain.toLowerCase()) {
		return undefined;
	}
	const rest = toks.slice(count + 1);
	if (rest.length === 0) {
		return 0;
	}
	const k = rest.length === 2 && rest[1].kind === 'integerLiteral' ? Number(rest[1].rawText) : undefined;
	return k === undefined ? undefined : rest[0].rawText === '+' ? k : rest[0].rawText === '-' ? -k : undefined;
}

/** An argument's value tokens, named (`Count:=0`) or at its position before any named one. */
function argumentByNameOrPosition(args: readonly VbaToken[][], name: string, position: number): VbaToken[] | undefined {
	for (const [index, arg] of args.entries()) {
		const toks = arg.filter((tok) => tok.kind !== 'comment');
		if (toks[1]?.rawText === ':=') {
			if (tokenName(toks[0])?.toLowerCase() === name.toLowerCase()) {
				return toks.slice(2);
			}
			continue;
		}
		if (index === position) {
			return toks;
		}
	}
	return undefined;
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
	const typed = hostReceiverTypes(resolveReceiverTypeAt(source, span.start + toks[i - 1].end, memberCtx), model);
	const resolved = typed.length > 0 ? typed : hostReceiverTypes(rangeItemType(source, span, toks, i - 1, memberCtx), model);
	// Of a union, only a part that has the member can run the call: the
	// member is judged on the one part that has it. `ActiveSheet` is a
	// Worksheet or a Chart, and only a Worksheet has Cells (issue #182).
	const having = resolved.filter((part) => resolveHostMember(part, name, model));
	// Parts that share the member and its type are judged as one:
	// `ActiveSheet.Shapes(0)` on a Worksheet or a Chart (issue #276).
	const returns = new Set(having.map((part) => resolveHostMember(part, name, model)!.returns));
	if (having.length === 0 || (having.length > 1 && (returns.size !== 1 || returns.has(undefined)))) {
		return undefined;
	}
	const receiver = having[0];
	const member = resolveHostMember(receiver, name, model)!;
	return { name, returns: member.returns, receiver, nameIndex: i, openIndex, closeIndex, args };
}

/**
 * `r.Item(1)` before the dot at `dotIndex`, with r a Range: a Range too,
 * though the model types Item as a Variant (issue #559, measured in Excel
 * 16.0).
 */
function rangeItemType(source: string, span: Span, toks: readonly VbaToken[], dotIndex: number, memberCtx: MemberCompletionContext): string | undefined {
	if (toks[dotIndex]?.rawText !== '.' || toks[dotIndex - 1]?.rawText !== ')') {
		return undefined;
	}
	const close = dotIndex - 1;
	const open = toks.findIndex((tok, k) => tok.rawText === '(' && matchParenFrom(toks, k) === close);
	if (open < 3 || tokenText(toks[open - 1]) !== 'item' || toks[open - 2].rawText !== '.') {
		return undefined;
	}
	return resolveReceiverTypeAt(source, span.start + toks[open - 2].end, memberCtx) === 'Excel.Range' ? 'Excel.Range' : undefined;
}

/**
 * The host types a resolved receiver may be. A one-part union - what a
 * collection's Object-declared Item gives, `Worksheets(1)` (issue #114) - is
 * its part. A union with any part the model does not know is not judged.
 */
function hostReceiverTypes(resolved: string | undefined, model: HostObjectModel | undefined): string[] {
	if (!resolved) {
		return [];
	}
	const parts = resolved.startsWith('union:') ? resolved.slice('union:'.length).split('|') : [resolved];
	return parts.every((part) => getHostType(part, model)) ? parts : [];
}

/** The member call's name index when the statement is `a.b.Name args`, else -1. */
function firstExecutableTokenIndexOfMemberCall(toks: readonly VbaToken[], nameIndex: number): number {
	// Back over each `.` to the name before it, stepping over a call's
	// parentheses: `Range("A1").Insert Shift:=9` (issue #244).
	let j = nameIndex;
	while (j >= 2 && toks[j - 1].rawText === '.') {
		let k = j - 2;
		if (toks[k].rawText === ')') {
			const close = k;
			const open = toks.findIndex((tok, m) => tok.rawText === '(' && matchParenFrom(toks, m) === close);
			if (open < 1) {
				return -1;
			}
			k = open - 1;
		}
		j = k;
	}
	// Inside a With, `.Add "a b", ...` starts at its leading dot (issue #610).
	const first = firstExecutableTokenIndex(toks);
	return j === first || (toks[j - 1]?.rawText === '.' && j - 1 === first) ? nameIndex : -1;
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
	arrays: ReadonlySet<string>,
	sourceNames: ReadonlySet<string>,
	valueOf: (arg: readonly VbaToken[]) => number | undefined,
	push: PushFn,
	stringOf: (arg: readonly VbaToken[]) => string | undefined = literalString,
): void {
	const lower = callee.name.toLowerCase();
	if (callee.openIndex < 0) {
		return;
	}
	const argsSpan = callee.closeIndex > callee.openIndex + 1
		? { start: span.start + toks[callee.openIndex + 1].start, end: span.start + toks[callee.closeIndex - 1].end }
		: calleeSpan;
	if (callee.receiver === 'Excel.WorksheetFunction') {
		const why = worksheetFunctionRefusal(lower, callee.args);
		if (why) {
			push('hostArgumentOutOfRange', `WorksheetFunction.${callee.name}: ${why}. The worksheet error is raised as Run-time error '1004': Unable to get the ${callee.name} property of the WorksheetFunction class.`, argsSpan);
		}
		return;
	}
	// A range counts its Cells, Rows and Columns from its own top-left cell,
	// so a literal `Range("B2")` receiver moves the far edge in. Any other
	// receiver starts at A1 or below, so the count alone past the edge is
	// already off the sheet (issue #182).
	const origin = singleCellReceiver(toks, callee.nameIndex - 1);
	// A chain from a literal range, or a block of several cells: each step
	// moves or resizes the block, and one that leaves the sheet raises 1004
	// (issue #508, measured in Excel 16.0). One step from one literal cell
	// is the checks below.
	const oneStep = origin !== undefined && lower !== 'item';
	if (callee.receiver === 'Excel.Range' && CHAIN_MEMBERS.has(lower) && !oneStep) {
		const block = rangeChainReceiver(toks, callee.nameIndex - 1, valueOf, sourceNames);
		const next = block ? chainStep(block, lower, callee.args, valueOf) : undefined;
		if (block && next && offSheet(next)) {
			const rows = next.rows === 1 ? `row ${next.row}` : `rows ${next.row} to ${next.row + next.rows - 1}`;
			const columns = next.width === 1 ? `column ${next.column}` : `columns ${next.column} to ${next.column + next.width - 1}`;
			push(
				'hostArgumentOutOfRange',
				`${callee.name}(${callee.args.map((arg) => (arg.length === 0 ? '' : valueOf(arg) ?? '...')).join(', ')}) on ${block.text} reaches ${rows}, ${columns}, off the sheet. This will raise Run-time error '1004': Application-defined or object-defined error.`,
				argsSpan,
			);
			return;
		}
	}
	// On a range, 0 and below reach above or left of it, and are judged
	// only on a range the code spells out (issue #275).
	if (callee.receiver === 'Excel.Range' && RANGE_RELATIVE_MEMBERS.has(lower) && checkBeforeRange(span, toks, callee, valueOf, push)) {
		return;
	}
	const fromRow = origin?.row ?? 1;
	const fromColumn = origin?.column ?? 1;
	const from = origin ? ` from ${origin.text}` : '';
	if (lower === 'cells' && callee.returns === 'Excel.Range') {
		// A column given by its letters: "AB" and "$a" run; "XFE", "AAAA", ""
		// and "A " raise 13; "A1" and "5" raise 1004 (issue #243).
		const column = callee.args[1];
		const columnText = column ? stringOf(column) : undefined;
		if (column?.length === 1 && columnText !== undefined) {
			const text = columnText;
			const letters = /^\$?([A-Za-z]{1,3})$/.exec(text);
			if (!letters || columnNumber(letters[1]) > EXCEL_MAX_COLUMN) {
				const digits = /\d/.test(text);
				push(
					'hostArgumentOutOfRange',
					digits
						? `Cells takes a column as a number or its letters, and "${text}" is neither. This will raise Run-time error '1004': Application-defined or object-defined error.`
						: `"${text}" names no column: the letters run A to XFD. This will raise Run-time error '13': Type mismatch.`,
					argSpan(span, column),
				);
				return;
			}
		}
		for (const arg of callee.args) {
			const value = valueOf(arg);
			if (value !== undefined && value < 1) {
				push('hostArgumentOutOfRange', `Cells takes a row and a column of at least 1; ${value} names no cell. This will raise Run-time error '1004': Application-defined or object-defined error.`, argSpan(span, arg));
				return;
			}
		}
		if (callee.args.length === 2) {
			const row = valueOf(callee.args[0]);
			const column = valueOf(callee.args[1]);
			const edge = pastSheetEdge(row === undefined ? undefined : fromRow + row - 1, column === undefined ? undefined : fromColumn + column - 1);
			if (edge) {
				push('hostArgumentOutOfRange', `Cells(${row ?? '...'}, ${column ?? '...'})${from} ${edge}. This will raise Run-time error '1004': Application-defined or object-defined error.`, argsSpan);
			}
		}
		return;
	}
	if ((lower === 'rows' || lower === 'columns') && callee.returns === 'Excel.Range' && callee.args.length === 1) {
		// `Columns("XFE")`: a letter past XFD names no column (issue #276,
		// measured in Excel 16.0: 13).
		const letters = lower === 'columns' ? stringOf(callee.args[0]) : undefined;
		if (letters !== undefined && /^[A-Za-z]{1,3}$/.test(letters) && fromColumn + columnNumber(letters) - 1 > EXCEL_MAX_COLUMN) {
			push('hostArgumentOutOfRange', `Columns("${letters}")${from} names a column past XFD, the last. This will raise Run-time error '13': Type mismatch.`, argsSpan);
			return;
		}
		const index = valueOf(callee.args[0]);
		const edge = index === undefined ? undefined : lower === 'rows' ? pastSheetEdge(fromRow + index - 1, undefined) : pastSheetEdge(undefined, fromColumn + index - 1);
		if (edge) {
			push('hostArgumentOutOfRange', `${callee.name}(${index})${from} ${edge}. This will raise Run-time error '1004': Application-defined or object-defined error.`, argsSpan);
		} else if (index !== undefined && index >= 1) {
			// A whole row or column is many cells (issue #454).
			checkMultiCellAsScalar(source, span, toks, callee, `${callee.name}(${index})`, env, arrays, push);
		}
		return;
	}
	if (lower === 'resize' && callee.receiver === 'Excel.Range') {
		for (const arg of callee.args) {
			const value = valueOf(arg);
			if (value !== undefined && value < 1) {
				push('hostArgumentOutOfRange', `Resize needs at least one row and one column; ${value} gives none. This will raise Run-time error '1004': Application-defined or object-defined error.`, argSpan(span, arg));
				return;
			}
		}
		const rows = callee.args[0] ? valueOf(callee.args[0]) : undefined;
		const columns = callee.args[1] ? valueOf(callee.args[1]) : undefined;
		const edge = pastSheetEdge(rows === undefined ? undefined : fromRow + rows - 1, columns === undefined ? undefined : fromColumn + columns - 1);
		if (edge) {
			push('hostArgumentOutOfRange', `Resize(${callee.args.map((arg) => valueOf(arg) ?? '...').join(', ')})${from} ${edge}. This will raise Run-time error '1004': Application-defined or object-defined error.`, argsSpan);
		}
		return;
	}
	if (lower === 'offset' && callee.receiver === 'Excel.Range') {
		const origin = singleCellReceiver(toks, callee.nameIndex - 1);
		if (!origin) {
			return;
		}
		const rowOffset = callee.args[0] ? valueOf(callee.args[0]) : 0;
		const columnOffset = callee.args[1] ? valueOf(callee.args[1]) : 0;
		if (rowOffset === undefined || columnOffset === undefined) {
			return;
		}
		const row = origin.row + rowOffset;
		const column = origin.column + columnOffset;
		if (row < 1 || column < 1 || row > EXCEL_MAX_ROW || column > EXCEL_MAX_COLUMN) {
			push('hostArgumentOutOfRange', `Offset(${rowOffset}, ${columnOffset}) from ${origin.text} lands at row ${row}, column ${column}, off the sheet. This will raise Run-time error '1004': Application-defined or object-defined error.`, argsSpan);
		}
		return;
	}
	if (lower === 'range' && callee.returns === 'Excel.Range') {
		const areas = callee.args.map((arg) => {
			const text = stringOf(arg);
			return text === undefined ? undefined : parseA1Address(text);
		});
		for (let k = 0; k < callee.args.length; k++) {
			const area = areas[k];
			if (area?.blank) {
				push('hostArgumentOutOfRange', `Range takes an address or a name, and "${area.text}" is blank. This will raise Run-time error '1004': Method 'Range' of object failed.`, argSpan(span, callee.args[k]));
				return;
			}
			if (area && !area.valid && !area.mayBeName) {
				push('hostArgumentOutOfRange', `"${area.text}" is not a cell address Excel accepts: rows run 1 to ${EXCEL_MAX_ROW} and columns A to XFD. This will raise Run-time error '1004': Method 'Range' of object failed.`, argSpan(span, callee.args[k]));
				return;
			}
		}
		if (callee.args.length === 1 && areas[0]?.valid && areas[0].multiCell) {
			checkMultiCellAsScalar(source, span, toks, callee, `Range("${areas[0].text}")`, env, arrays, push);
		}
		// `Range("A1", "B2")`: two different cells span a block (issue #454).
		const [a, b] = areas;
		if (callee.args.length === 2 && a?.valid && b?.valid && !a.multiCell && !b.multiCell && a.row !== undefined && b.row !== undefined
			&& (a.row !== b.row || a.column !== b.column)) {
			checkMultiCellAsScalar(source, span, toks, callee, `Range("${a.text}", "${b.text}")`, env, arrays, push);
		}
		// `Range(Cells(1, 1), Cells(2, 1))` the same way (issue #492).
		const [from, to] = callee.args.length === 2 ? callee.args.map((arg) => cellsCall(arg, valueOf)) : [];
		if (from && to && (from.row !== to.row || from.column !== to.column)) {
			checkMultiCellAsScalar(source, span, toks, callee, `Range(Cells(${from.row}, ${from.column}), Cells(${to.row}, ${to.column}))`, env, arrays, push);
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
	display: string,
	env: ReadonlyMap<string, string>,
	arrays: ReadonlySet<string>,
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
		`${display} read as a value is a two-dimensional array, ${use}. This will raise Run-time error '13': Type mismatch.`;
	// `Len(Range("A1:A2"))`, `InStr(1, Range("A1:A2"), "a")`: a whole
	// argument of a built-in that takes a single value (issue #454).
	const wholeArgument = ['(', ','].includes(toks[start - 1]?.rawText ?? '') && [')', ','].includes(toks[end + 1]?.rawText ?? '');
	const call = wholeArgument ? enclosingBuiltin(toks, start) : undefined;
	if (call && SCALAR_ARGUMENT_BUILTINS.has(call.lower)) {
		push('multiCellRangeAsScalar', message(`which ${call.display} cannot take as one value`), valueSpan);
		return;
	}
	// `Select Case Range("A1:A2")` compares the array with each Case.
	const head = firstExecutableTokenIndex(toks);
	if (tokenText(toks[head]) === 'select' && tokenText(toks[head + 1]) === 'case' && start === head + 2 && end === toks.length - 1) {
		push('multiCellRangeAsScalar', message('which Select Case cannot compare'), valueSpan);
		return;
	}
	const bare = bareAssignmentTarget(source, span);
	if (bare) {
		const eq = toks.findIndex((tok) => tok.rawText === '=');
		if (eq === start - 1 && end === toks.length - 1) {
			const target = normalizeType(env.get(bare.name.toLowerCase()));
			if (target && SCALAR_TYPES.has(target)) {
				// Into an array of another element type it is its Variant
				// elements that do not fit (issue #194).
				const holder = arrays.has(bare.name.toLowerCase())
					? `whose Variant elements an array of ${env.get(bare.name.toLowerCase())} cannot take`
					: `which a ${env.get(bare.name.toLowerCase())} variable cannot hold`;
				push('multiCellRangeAsScalar', message(holder), valueSpan);
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
		// Only the condition of a one-line If is judged here: a range after
		// Then or Else belongs to that branch's own statement, `If r Is
		// Nothing Then Set r = ws.Range("A1:P36")` (issue #140).
		const then = head === 'if' ? toks.findIndex((tok) => tokenText(tok) === 'then') : -1;
		if (then > 0 && start > then) {
			return;
		}
	}
	// The whole condition: `If ws.Range("A1:A2") Then` reads the array as
	// True or False (issue #492, measured in Excel 16.0).
	const opener = tokenText(toks[start - 1]);
	if (!bare && ['if', 'elseif', 'while', 'until'].includes(opener) && (end === toks.length - 1 || tokenText(toks[end + 1]) === 'then')) {
		push('multiCellRangeAsScalar', message(`which ${opener === 'elseif' ? 'ElseIf' : opener === 'if' ? 'If' : opener === 'while' ? 'While' : 'Until'} cannot read as True or False`), valueSpan);
		return;
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

/**
 * VBA built-ins that read each argument as one value, so a multi-cell Range
 * given whole raises 13 (issue #454, measured in Excel 16.0: CStr, Len,
 * LenB, Val, CLng, CDbl, CBool, Trim, UCase, Left, InStr, Abs, Int, Format
 * and Hex; the others here take the same kind of argument). IsEmpty,
 * IsNumeric, IsArray, TypeName, VarType and UBound take the array and run.
 */
const SCALAR_ARGUMENT_BUILTINS: ReadonlySet<string> = new Set([
	'cstr', 'len', 'lenb', 'val', 'clng', 'cint', 'cdbl', 'csng', 'ccur', 'cbyte', 'cbool', 'cdate',
	'trim', 'ltrim', 'rtrim', 'ucase', 'lcase', 'left', 'right', 'mid', 'instr', 'abs', 'int', 'fix', 'format', 'hex', 'oct',
]);

/** The VBA built-in whose argument list holds the token at `index`, bare or VBA-qualified, `$` spellings included. */
function enclosingBuiltin(toks: readonly VbaToken[], index: number): { lower: string; display: string } | undefined {
	let depth = 0;
	for (let i = index - 1; i >= 0; i--) {
		const raw = toks[i].rawText;
		if (raw === ')') {
			depth++;
		} else if (raw === '(') {
			if (depth === 0) {
				const nameAt = toks[i - 1]?.rawText === '$' ? i - 2 : i - 1;
				const name = tokenName(toks[nameAt]);
				const qualified = toks[nameAt - 1]?.rawText === '.';
				if (!name || (qualified && tokenText(toks[nameAt - 2]) !== 'vba')) {
					return undefined;
				}
				return { lower: name.toLowerCase(), display: toks.slice(nameAt, i).map((tok) => tok.rawText).join('') };
			}
			depth--;
		}
	}
	return undefined;
}

/** `Worksheets(1).Name = "a:b"`: the receiver's type and the literal decide. */
function checkSheetNameAssignment(
	source: string,
	span: Span,
	toks: readonly VbaToken[],
	memberCtx: MemberCompletionContext,
	sourceNames: ReadonlySet<string>,
	push: PushFn,
): void {
	const n = toks.length;
	const eq = toks.findIndex((tok) => tok.rawText === '=');
	if (eq < 2 || eq === n - 1 || tokenText(toks[eq - 1]) !== 'name' || toks[eq - 2].rawText !== '.') {
		return;
	}
	const name = spelledOutText(toks.slice(eq + 1), sourceNames);
	if (name === undefined) {
		return;
	}
	const resolved = resolveReceiverTypeAt(source, span.start + toks[eq - 2].end, memberCtx);
	// `Worksheets(1)` is a one-part union, `Sheets(1)` a Worksheet-or-Chart union: both are sheets.
	const parts = resolved ? (resolved.startsWith('union:') ? resolved.slice('union:'.length).split('|') : [resolved]) : [];
	if (parts.length === 0 || !parts.every((part) => part === 'Excel.Worksheet' || part === 'Excel.Chart')) {
		return;
	}
	let problem: string | undefined;
	let error = 'You typed an invalid name for a sheet or chart.';
	if (name.length === 0) {
		problem = 'a sheet name cannot be blank';
	} else if (name.length > SHEET_NAME_MAX) {
		problem = `a sheet name has at most ${SHEET_NAME_MAX} characters, and this one has ${name.length}`;
	} else if (SHEET_NAME_FORBIDDEN.test(name)) {
		problem = 'a sheet name cannot contain any of : \\ / ? * [ ]';
	} else if (name.toLowerCase() === 'history') {
		problem = 'Excel keeps History for itself, in any case';
		error = 'History is a reserved name.';
	} else if (name.startsWith("'") || name.endsWith("'")) {
		problem = 'a sheet name cannot start or end with an apostrophe';
	}
	if (problem) {
		push('sheetNameInvalid', `Excel refuses this name: ${problem}. This will raise Run-time error '1004': ${error}`, { start: span.start + toks[eq + 1].start, end: span.start + toks[n - 1].end });
	}
}

/**
 * The text an expression spells out from literals alone: a string
 * literal, `String$(32, "a")`, `Space$(3)`, and `&` between them (issue
 * #276). Undefined for anything else, or when the module declares its
 * own String or Space.
 */
function spelledOutText(toks: readonly VbaToken[], sourceNames: ReadonlySet<string>): string | undefined {
	const parts: VbaToken[][] = [[]];
	let depth = 0;
	for (const tok of toks) {
		if (tok.kind === 'comment') {
			continue;
		}
		depth += tok.rawText === '(' ? 1 : tok.rawText === ')' ? -1 : 0;
		if (depth === 0 && tok.rawText === '&') {
			parts.push([]);
		} else {
			parts[parts.length - 1].push(tok);
		}
	}
	let out = '';
	for (const written of parts) {
		// `String$` lexes as String and a `$`.
		const part = written[1]?.rawText === '$' ? [written[0], ...written.slice(2)] : written;
		if (part.length === 1 && part[0].kind === 'stringLiteral') {
			out += stringLiteralValue(part[0].rawText);
			continue;
		}
		const fn = tokenText(part[0]).replace(/\$$/, '');
		const count = part[2]?.kind === 'integerLiteral' ? parseVbaIntegerLiteral(part[2].rawText) : undefined;
		if (sourceNames.has(fn) || part[1]?.rawText !== '(' || count === undefined || count < 0 || count > 1000) {
			return undefined;
		}
		if (fn === 'space' && part.length === 4 && part[3].rawText === ')') {
			out += ' '.repeat(count);
		} else if (fn === 'string' && part.length === 6 && part[3].rawText === ',' && part[4].kind === 'stringLiteral' && part[5].rawText === ')' && stringLiteralValue(part[4].rawText) !== '') {
			out += stringLiteralValue(part[4].rawText)[0].repeat(count);
		} else {
			return undefined;
		}
	}
	return out;
}

/** Where a row or column past the bottom or right edge of the sheet lands, in words. */
function pastSheetEdge(row: number | undefined, column: number | undefined): string | undefined {
	if (row !== undefined && row > EXCEL_MAX_ROW) {
		return `reaches row ${row}, past the last row of the sheet, ${EXCEL_MAX_ROW}`;
	}
	if (column !== undefined && column > EXCEL_MAX_COLUMN) {
		return `reaches column ${column}, past the last column of the sheet, ${EXCEL_MAX_COLUMN} (XFD)`;
	}
	return undefined;
}

/** The single-cell literal `Range("B2")` ending at `toks[closeIndex]`, when that is the receiver. */
function singleCellReceiver(toks: readonly VbaToken[], dotIndex: number): { row: number; column: number; text: string } | undefined {
	const block = literalRangeReceiver(toks, dotIndex);
	return block && block.rows === 1 && block.width === 1 ? block : undefined;
}

/** `Cells(1, 1)` or `ws.Cells(1, 1)` with known numbers, as an argument of Range. */
function cellsCall(arg: readonly VbaToken[], valueOf: (arg: readonly VbaToken[]) => number | undefined): { row: number; column: number } | undefined {
	const toks = arg.filter((tok) => tok.kind !== 'comment');
	const at = toks.length > 2 && toks[1]?.rawText === '.' ? 2 : 0;
	if (tokenText(toks[at]) !== 'cells' || toks[at + 1]?.rawText !== '(' || matchParenFrom(toks, at + 1) !== toks.length - 1) {
		return undefined;
	}
	const args = splitTopLevelTokenGroups(toks, at + 2, ',', toks.length - 1);
	const row = args.length === 2 ? valueOf(args[0]) : undefined;
	const column = args.length === 2 ? valueOf(args[1]) : undefined;
	return row !== undefined && column !== undefined && row >= 1 && column >= 1 ? { row, column } : undefined;
}

/** A block of cells: its top-left row and column and its size. */
interface CellBlock {
	row: number;
	column: number;
	rows: number;
	width: number;
	/** The expression that names it, as written. */
	text: string;
	/**
	 * Rows or columns from EntireRow, EntireColumn, Rows(n) or Columns(n): one
	 * index in Item then counts rows or columns: `Columns(4).EntireColumn.Item(0)`
	 * is column C (issue #556, measured in Excel 16.0).
	 */
	mode?: 'rows' | 'columns';
}

/** The Range members a chain follows (issue #508). */
const CHAIN_MEMBERS: ReadonlySet<string> = new Set(['offset', 'resize', 'cells', 'item', 'rows', 'columns', 'entirerow', 'entirecolumn']);

/**
 * The block of cells the Range expression before the dot at `dotIndex`
 * names, followed through Offset, Resize, Cells, Item, Rows, Columns,
 * EntireRow and EntireColumn from a literal `Range("B2:C3")` (issue #508,
 * measured in Excel 16.0). Undefined where any step is not known, and
 * where a step already lands off the sheet, which that step reports.
 */
function rangeChainReceiver(
	toks: readonly VbaToken[],
	dotIndex: number,
	valueOf: (arg: readonly VbaToken[]) => number | undefined,
	sourceNames: ReadonlySet<string>,
): CellBlock | undefined {
	if (toks[dotIndex]?.rawText !== '.') {
		return undefined;
	}
	const last = toks[dotIndex - 1];
	const word = tokenText(last);
	if ((word === 'entirerow' || word === 'entirecolumn') && toks[dotIndex - 2]?.rawText === '.') {
		const inner = rangeChainReceiver(toks, dotIndex - 2, valueOf, sourceNames);
		if (!inner) {
			return undefined;
		}
		const text = `${inner.text}.${last.rawText}`;
		return word === 'entirerow'
			? { row: inner.row, column: 1, rows: inner.rows, width: EXCEL_MAX_COLUMN, text, mode: 'rows' }
			: { row: 1, column: inner.column, rows: EXCEL_MAX_ROW, width: inner.width, text, mode: 'columns' };
	}
	// `Cells`, `Rows` and `Columns` of a sheet, without an index: the whole
	// sheet (issue #628, measured in Excel 16.0).
	if ((word === 'cells' || word === 'rows' || word === 'columns') && ofSheet(toks, dotIndex - 1, sourceNames)) {
		const mode = word === 'rows' ? 'rows' : word === 'columns' ? 'columns' : undefined;
		return { row: 1, column: 1, rows: EXCEL_MAX_ROW, width: EXCEL_MAX_COLUMN, text: last.rawText, ...(mode ? { mode } : {}) };
	}
	if (last?.rawText !== ')') {
		return undefined;
	}
	const close = dotIndex - 1;
	const open = toks.findIndex((tok, k) => tok.rawText === '(' && matchParenFrom(toks, k) === close);
	const name = tokenText(toks[open - 1]);
	if (open < 1) {
		return undefined;
	}
	if (name === 'range') {
		// A Range on a range counts from that range: not followed.
		if (toks[open - 2]?.rawText === '.' && rangeChainReceiver(toks, open - 2, valueOf, sourceNames)) {
			return undefined;
		}
		const block = literalRangeReceiver(toks, dotIndex) ?? wholeLinesAt(toks, open - 1);
		return block ? { ...block } : undefined;
	}
	// `Cells(1, 2)`, `Rows(3)`, `Columns(2)` of the sheet: a cell, a whole row
	// or a whole column (issues #308 and #628, measured in Excel 16.0).
	if (ofSheet(toks, open - 1, sourceNames) && (name === 'cells' || name === 'rows' || name === 'columns')) {
		const args = close > open + 1 ? splitTopLevelTokenGroups(toks, open + 1, ',', close) : [];
		const first = args[0] ? valueOf(args[0]) : undefined;
		const text = toks.slice(open - 1, close + 1).map((tok) => tok.rawText).join('');
		if (name === 'cells' && args.length === 2) {
			const column = valueOf(args[1]);
			return first !== undefined && column !== undefined && first >= 1 && column >= 1 && first <= EXCEL_MAX_ROW && column <= EXCEL_MAX_COLUMN
				? { row: first, column, rows: 1, width: 1, text } : undefined;
		}
		if (args.length !== 1 || first === undefined || first < 1) {
			return undefined;
		}
		// Item on a row counts rows: `Rows(5).Item(2)` is row 6.
		if (name === 'rows') {
			return first <= EXCEL_MAX_ROW ? { row: first, column: 1, rows: 1, width: EXCEL_MAX_COLUMN, text, mode: 'rows' } : undefined;
		}
		return name === 'columns' && first <= EXCEL_MAX_COLUMN ? { row: 1, column: first, rows: EXCEL_MAX_ROW, width: 1, text, mode: 'columns' } : undefined;
	}
	if (!CHAIN_MEMBERS.has(name) || toks[open - 2]?.rawText !== '.') {
		return undefined;
	}
	const inner = rangeChainReceiver(toks, open - 2, valueOf, sourceNames);
	if (!inner) {
		return undefined;
	}
	const args = close > open + 1 ? splitTopLevelTokenGroups(toks, open + 1, ',', close) : [];
	const next = chainStep(inner, name, args, valueOf);
	if (!next || offSheet(next)) {
		return undefined;
	}
	return { ...next, text: `${inner.text}.${toks.slice(open - 1, close + 1).map((tok) => tok.rawText).join('')}` };
}

/**
 * Whether the member at `toks[at]` is the sheet's own: unqualified, or
 * after `ActiveSheet`, `Worksheets(...)` or `Sheets(...)` (issue #628).
 * Unqualified, a name the code declares is its own.
 */
function ofSheet(toks: readonly VbaToken[], at: number, sourceNames: ReadonlySet<string>): boolean {
	if (toks[at - 1]?.rawText !== '.') {
		return !sourceNames.has(tokenText(toks[at]));
	}
	const before = toks[at - 2];
	if (tokenText(before) === 'activesheet') {
		return true;
	}
	if (before?.rawText !== ')') {
		return false;
	}
	const open = toks.findIndex((tok, k) => tok.rawText === '(' && matchParenFrom(toks, k) === at - 2);
	const name = tokenText(toks[open - 1]);
	return open >= 1 && (name === 'worksheets' || name === 'sheets');
}

/** `Range("5:6")` or `Range("C:D")` starting at `toks[at]`: whole rows or whole columns (issue #628). */
function wholeLinesAt(toks: readonly VbaToken[], at: number): CellBlock | undefined {
	if (tokenText(toks[at]) !== 'range' || toks[at + 1]?.rawText !== '(' || toks[at + 2]?.kind !== 'stringLiteral' || toks[at + 3]?.rawText !== ')') {
		return undefined;
	}
	const text = stringLiteralValue(toks[at + 2].rawText);
	const rows = /^\$?(\d+):\$?(\d+)$/.exec(text);
	if (rows) {
		const [top, bottom] = [Number(rows[1]), Number(rows[2])].sort((a, b) => a - b);
		return top >= 1 && bottom <= EXCEL_MAX_ROW ? { row: top, column: 1, rows: bottom - top + 1, width: EXCEL_MAX_COLUMN, text: `Range("${text}")` } : undefined;
	}
	const columns = /^\$?([A-Za-z]{1,3}):\$?([A-Za-z]{1,3})$/.exec(text);
	if (columns) {
		const [left, right] = [columnNumber(columns[1]), columnNumber(columns[2])].sort((a, b) => a - b);
		return right <= EXCEL_MAX_COLUMN ? { row: 1, column: left, rows: EXCEL_MAX_ROW, width: right - left + 1, text: `Range("${text}")` } : undefined;
	}
	return undefined;
}

/** One member applied to a block: the block it names, or undefined where an argument is not known. */
function chainStep(
	block: CellBlock,
	name: string,
	args: readonly VbaToken[][],
	valueOf: (arg: readonly VbaToken[]) => number | undefined,
): Omit<CellBlock, 'text'> | undefined {
	const value = (k: number, missing: number): number | undefined => (args[k] === undefined || args[k].length === 0 ? missing : valueOf(args[k]));
	switch (name) {
		case 'offset': {
			const rows = value(0, 0);
			const columns = value(1, 0);
			return rows === undefined || columns === undefined ? undefined : { ...block, row: block.row + rows, column: block.column + columns };
		}
		case 'resize': {
			const rows = value(0, block.rows);
			const width = value(1, block.width);
			return rows === undefined || width === undefined || rows < 1 || width < 1 ? undefined : { ...block, rows, width };
		}
		case 'cells':
		case 'item': {
			if (args.length === 2) {
				const row = value(0, 1);
				const column = value(1, 1);
				return row === undefined || column === undefined ? undefined : { row: block.row + row - 1, column: block.column + column - 1, rows: 1, width: 1 };
			}
			const index = args.length === 1 ? value(0, 1) : undefined;
			if (index === undefined) {
				return undefined;
			}
			if (name === 'item' && block.mode === 'columns') {
				return { ...block, column: block.column + index - 1, width: 1 };
			}
			if (name === 'item' && block.mode === 'rows') {
				return { ...block, row: block.row + index - 1, rows: 1 };
			}
			const k = index - 1;
			return { row: block.row + Math.trunc(k / block.width), column: block.column + (k % block.width), rows: 1, width: 1 };
		}
		case 'rows': {
			const index = args.length === 1 ? value(0, 1) : undefined;
			return index === undefined ? undefined : { ...block, row: block.row + index - 1, rows: 1, mode: 'rows' };
		}
		case 'columns': {
			const index = args.length === 1 ? value(0, 1) : undefined;
			return index === undefined ? undefined : { ...block, column: block.column + index - 1, width: 1, mode: 'columns' };
		}
	}
	return undefined;
}

/** Whether any cell of the block is off the sheet. */
function offSheet(block: Omit<CellBlock, 'text'>): boolean {
	return block.row < 1 || block.column < 1 || block.row + block.rows - 1 > EXCEL_MAX_ROW || block.column + block.width - 1 > EXCEL_MAX_COLUMN;
}

/** The literal `Range("B2:C3")` before the dot at `dotIndex`: its top-left cell and its size. */
function literalRangeReceiver(toks: readonly VbaToken[], dotIndex: number): { row: number; column: number; rows: number; width: number; text: string } | undefined {
	if (toks[dotIndex]?.rawText !== '.' || toks[dotIndex - 1]?.rawText !== ')') {
		return undefined;
	}
	const close = dotIndex - 1;
	const open = toks.findIndex((tok, k) => tok.rawText === '(' && matchParenFrom(toks, k) === close);
	return open < 1 ? undefined : literalRangeAt(toks, open - 1);
}

/** `Range("B2:C3")` starting at `toks[at]`: its top-left cell and its size. */
function literalRangeAt(toks: readonly VbaToken[], at: number): { row: number; column: number; rows: number; width: number; text: string } | undefined {
	if (tokenText(toks[at]) !== 'range' || toks[at + 1]?.rawText !== '(' || toks[at + 2]?.kind !== 'stringLiteral' || toks[at + 3]?.rawText !== ')') {
		return undefined;
	}
	const area = parseA1Address(stringLiteralValue(toks[at + 2].rawText));
	if (!area?.valid || area.row === undefined || area.column === undefined) {
		return undefined;
	}
	const endRow = area.endRow ?? area.row;
	const endColumn = area.endColumn ?? area.column;
	return {
		row: Math.min(area.row, endRow),
		column: Math.min(area.column, endColumn),
		rows: Math.abs(endRow - area.row) + 1,
		width: Math.abs(endColumn - area.column) + 1,
		text: `Range("${area.text}")`,
	};
}

/**
 * Cells, Item, Rows or Columns on a range at 0 or below (issue #275,
 * measured in Excel 16.0). They count from the range's top-left cell,
 * so they raise 1004 only where they land above row 1 or left of
 * column A, which is known for a literal `Range("B2")` receiver alone.
 * One index over a range w columns wide is Cells((i - 1) \ w + 1,
 * (i - 1) Mod w + 1), with VBA's truncating \ and Mod. Returns true
 * when an index is 0 or below, reported or not, so the caller's checks
 * for a sheet's own Cells do not run.
 */
function checkBeforeRange(
	span: Span,
	toks: readonly VbaToken[],
	callee: HostCallee,
	valueOf: (arg: readonly VbaToken[]) => number | undefined,
	push: PushFn,
): boolean {
	const values = callee.args.map((arg) => valueOf(arg));
	if (!values.some((value) => value !== undefined && value < 1)) {
		return false;
	}
	const block = literalRangeReceiver(toks, callee.nameIndex - 1);
	const lower = callee.name.toLowerCase();
	if (!block || callee.args.length > 2 || ((lower === 'rows' || lower === 'columns') && callee.args.length !== 1)) {
		return true;
	}
	let row: number | undefined;
	let column: number | undefined;
	if (lower === 'rows') {
		row = block.row + values[0]! - 1;
	} else if (lower === 'columns') {
		column = block.column + values[0]! - 1;
	} else if (callee.args.length === 2) {
		row = values[0] === undefined ? undefined : block.row + values[0] - 1;
		column = values[1] === undefined ? undefined : block.column + values[1] - 1;
	} else {
		const k = values[0]! - 1;
		row = block.row + Math.trunc(k / block.width);
		column = block.column + (k % block.width);
	}
	if ((row !== undefined && row < 1) || (column !== undefined && column < 1)) {
		const where = row !== undefined && row < 1 ? `row ${row}, above row 1` : `column ${column}, left of column A`;
		push(
			'hostArgumentOutOfRange',
			`${callee.name}(${values.map((value) => value ?? '...').join(', ')}) counts from the top-left cell of ${block.text} and lands at ${where}. This will raise Run-time error '1004': Application-defined or object-defined error.`,
			{ start: span.start + toks[callee.openIndex + 1].start, end: span.start + toks[callee.closeIndex - 1].end },
		);
	}
	return true;
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
	/** Empty or spaces only: no address and no name (issue #276). */
	blank?: boolean;
	/**
	 * Every part past the sheet could be a workbook name: it starts with a
	 * letter and holds no `$`. A0, XFE1, A1048577 and XFE are names Excel
	 * accepts, and Range finds them, alone, in `A0:D4` or after `Sheet1!`.
	 */
	mayBeName?: boolean;
	row?: number;
	column?: number;
	/** The second cell of `A1:B2`. */
	endRow?: number;
	endColumn?: number;
}

/**
 * Parses an A1-style address literal: `A1`, `$A$1`, `A1:B2`, `A:A`, `1:1`,
 * with an optional `Sheet1!` or `'My Sheet'!` prefix. Anything else (a name,
 * an R1C1 address, a union) is not judged.
 */
function parseA1Address(text: string): A1Area | undefined {
	if (text.trim() === '') {
		return { text, valid: false, multiCell: false, blank: true };
	}
	const body = text.replace(/^(?:'[^']*'|[^!'\s]+)!/, '');
	// "R1C1" is an R1C1-style reference, which Range does not read and no
	// workbook name may be (issue #276, measured in Excel 16.0: 1004).
	if (/^R\d+C\d+$/i.test(body)) {
		return { text, valid: false, multiCell: false };
	}
	const cell = /^\$?([A-Za-z]{1,3})\$?(\d+)$/;
	const parts = body.split(':');
	if (parts.length > 2) {
		return undefined;
	}
	// "A1:", ":A1" and ":" leave a side empty (issue #243).
	if (parts.length === 2 && parts.some((part) => part === '')) {
		return { text, valid: false, multiCell: false };
	}
	const cells = parts.map((part) => cell.exec(part));
	if (cells.every((match) => match)) {
		const rows = cells.map((match) => Number(match![2]));
		const columns = cells.map((match) => columnNumber(match![1]));
		const inRange = parts.map((_, k) => rows[k] >= 1 && rows[k] <= EXCEL_MAX_ROW && columns[k] >= 1 && columns[k] <= EXCEL_MAX_COLUMN);
		const valid = inRange.every((ok) => ok);
		const mayBeName = parts.every((part, k) => inRange[k] || !part.includes('$'));
		const multiCell = cells.length === 2 && (rows[0] !== rows[1] || columns[0] !== columns[1]);
		return { text, valid, multiCell, mayBeName, row: rows[0], column: columns[0], endRow: rows[1], endColumn: columns[1] };
	}
	if (parts.length === 2) {
		const columnsOnly = parts.map((part) => /^\$?([A-Za-z]{1,3})$/.exec(part));
		if (columnsOnly.every((match) => match)) {
			const inRange = columnsOnly.map((match) => columnNumber(match![1]) <= EXCEL_MAX_COLUMN);
			const valid = inRange.every((ok) => ok);
			const mayBeName = parts.every((part, k) => inRange[k] || !part.includes('$'));
			return { text, valid, multiCell: true, mayBeName };
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
/** The saved workbook's sheets, and what the project's code may do to them. */
export interface WorkbookSheetsCheck {
	sheets: readonly WorkbookSheetInfo[];
	changes: SheetChanges;
}

/** Both halves or nothing: sheets alone cannot say what code adds at run time. */
export function workbookSheetsToCheck(opts: AnalyzeModuleOptions): WorkbookSheetsCheck | undefined {
	return opts.workbookSheets && opts.projectSheetChanges
		? { sheets: opts.workbookSheets, changes: opts.projectSheetChanges }
		: undefined;
}

/** The sheet kinds each of ThisWorkbook's sheet collections holds; undefined is every kind. */
const SHEET_COLLECTION_KINDS: ReadonlyMap<string, WorkbookSheetInfo['kind'] | undefined> = new Map([
	['sheets', undefined],
	['worksheets', 'worksheet'],
	['charts', 'chartsheet'],
]);

const SHEET_KIND_WORDS: Readonly<Record<WorkbookSheetInfo['kind'], string>> = {
	worksheet: 'worksheet',
	chartsheet: 'chart sheet',
	dialogsheet: 'dialog sheet',
	macrosheet: 'macro sheet',
};

/** Names Excel gives a sheet that code adds or copies: Sheet4, Chart2, Sheet1 (2). */
const MADE_SHEET_NAME = /^(sheet|chart|dialog|macro)\d+$|\s\(\d+\)$/i;

/**
 * `ThisWorkbook.Sheets("Missing")` and `ThisWorkbook.Worksheets(9)` on a
 * workbook without that sheet raise 9 (issue #229). ThisWorkbook only: a bare
 * `Sheets` is the active workbook's, which may be any workbook. A name or an
 * index that code in the project could have made - by adding, copying or
 * naming a sheet - is left alone.
 */
function checkWorkbookSheetAccess(
	source: string,
	span: Span,
	workbook: WorkbookSheetsCheck,
	sourceNames: ReadonlySet<string>,
	valueOf: (arg: readonly VbaToken[]) => number | undefined,
	push: PushFn,
): void {
	if (sourceNames.has('thisworkbook')) {
		return;
	}
	const toks = statementTokens(source, span);
	const lower = (i: number): string => toks[i]?.rawText.toLowerCase() ?? '';
	for (let i = 0; i < toks.length; i++) {
		if (lower(i) !== 'thisworkbook' || lower(i + 1) !== '.') {
			continue;
		}
		// `Application.ThisWorkbook` is the same object; `x.ThisWorkbook` is not known.
		if (lower(i - 1) === '.' && !(lower(i - 2) === 'application' && lower(i - 3) !== '.')) {
			continue;
		}
		const collection = lower(i + 2);
		if (!SHEET_COLLECTION_KINDS.has(collection)) {
			continue;
		}
		const kind = SHEET_COLLECTION_KINDS.get(collection);
		let open = i + 3;
		if (lower(open) === '.' && lower(open + 1) === 'item') {
			open += 2;
		}
		if (lower(open) !== '(') {
			continue;
		}
		const close = matchParenFrom(toks, open);
		if (close < 0) {
			continue;
		}
		const args = splitTopLevel(toks.slice(open + 1, close));
		if (args.length !== 1) {
			continue;
		}
		const held = workbook.sheets.filter((sheet) => kind === undefined || sheet.kind === kind);
		const argTokens = args[0].filter((tok) => tok.kind !== 'comment');
		const where = { start: span.start + toks[open + 1].start, end: span.start + toks[close - 1].end };
		const what = collection === 'worksheets' ? 'worksheet' : collection === 'charts' ? 'chart sheet' : 'sheet';
		if (argTokens.length === 1 && argTokens[0].kind === 'stringLiteral') {
			const name = argTokens[0].rawText.slice(1, -1).replace(/""/g, '"');
			// Excel matches names without regard to case; outside ASCII its rule is not known here.
			if (/[^\x20-\x7e]/.test(name) || held.some((sheet) => sheet.name.toLowerCase() === name.toLowerCase())) {
				continue;
			}
			const changes = workbook.changes;
			if (changes.assignsComputedName || changes.namesAssigned.has(name.toLowerCase()) || (changes.addsSheets && MADE_SHEET_NAME.test(name))) {
				continue;
			}
			const other = workbook.sheets.find((sheet) => sheet.name.toLowerCase() === name.toLowerCase());
			const detail = other ? `'${other.name}' is a ${SHEET_KIND_WORDS[other.kind]}, not a ${what}` : `this workbook has no ${what} named '${name}'`;
			push('sheetNotInWorkbook', `${capitalize(detail)}. This will raise Run-time error '9': Subscript out of range.`, where);
			continue;
		}
		const index = valueOf(args[0]);
		// Index 0 and below are host-argument-out-of-range's.
		if (index === undefined || index < 1 || index <= held.length || workbook.changes.addsSheets) {
			continue;
		}
		const count = held.length === 1 ? `1 ${what}` : `${held.length} ${what}s`;
		push('sheetNotInWorkbook', `This workbook has ${count}, so index ${index} is past the last. This will raise Run-time error '9': Subscript out of range.`, where);
	}
}

function capitalize(text: string): string {
	return text.charAt(0).toUpperCase() + text.slice(1);
}

/** The text of an argument that is one string literal. */
function literalString(arg: readonly VbaToken[]): string | undefined {
	const toks = arg.filter((tok) => tok.kind !== 'comment');
	return toks.length === 1 && toks[0].kind === 'stringLiteral' ? stringLiteralValue(toks[0].rawText) : undefined;
}

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
