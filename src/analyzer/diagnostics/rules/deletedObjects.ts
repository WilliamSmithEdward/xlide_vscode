// Rule: an object used after the statement that deletes or closes it (issue
// #294, each measured in Excel 16.0).
//
// `ws.Delete` leaves the variable referring to a sheet that is gone: a
// member of it then raises -2147221080, "Method 'Name' of object
// '_Worksheet' failed", and a Range taken from it before raises 424. A
// closed Workbook raises the same -2147221080; a deleted Shape or Name 424;
// an Unlisted ListObject 1004. `ws Is Nothing` still runs and gives False,
// and a new `Set` ends what is known.
//
// A sheet or Range taken from a workbook before it closed is gone with it,
// and so are a closed Word document (5825) and the Ranges taken from it, a
// closed PowerPoint presentation, a deleted Slide and a Slide of a closed
// presentation (-2147188720) (issue #683, measured in Excel, Word and
// PowerPoint 16.0).
//
// What Erase leaves in a Variant that held an array is the array rules'
// (issue #420).
//
// Only a straight run of a procedure's statements is followed: a block that
// names a variable ends what is known of it.

import type { ConditionalActivityTracker } from '../../conditional/conditionalCompilation';
import type { VbaToken } from '../../lexer/tokenKinds';
import type { BodyNode, ModuleNode } from '../../parser/nodes';
import { isLeafStatement } from '../../parser/nodes';
import type { buildModuleSymbols } from '../../symbols/buildModuleSymbols';
import { procedureSymbolFor, type PushFn } from '../analysisContext';
import { normalizeType } from '../typeInference';
import { statementLabelDeclaration } from '../../flow/procedureLabels';
import { activeModuleMembers, isInactiveNode, matchParenFrom, statementTokensAfterLeadingLabel, tokenName, tokenText } from '../walker';

/** What deletes or closes each type, and what a member of it then raises. */
const ENDINGS: Readonly<Record<string, { verb: string; error: string }>> = {
	worksheet: { verb: 'delete', error: "'-2147221080': Method 'MEMBER' of object '_Worksheet' failed" },
	workbook: { verb: 'close', error: "'-2147221080': Method 'MEMBER' of object '_Workbook' failed" },
	shape: { verb: 'delete', error: "'424': Object required" },
	name: { verb: 'delete', error: "'424': Object required" },
	listobject: { verb: 'unlist', error: "'1004': Application-defined or object-defined error" },
	// Word and PowerPoint (issue #683, measured in Word and PowerPoint 16.0).
	document: { verb: 'close', error: "'5825': Object has been deleted" },
	presentation: { verb: 'close', error: "'-2147188720': Presentation (unknown member) : Object does not exist" },
	slide: { verb: 'delete', error: "'-2147188720': Slide (unknown member) : Object does not exist" },
};

/** Members of a sheet that give a Range of it. */
const SHEET_RANGES: ReadonlySet<string> = new Set(['range', 'cells', 'rows', 'columns', 'usedrange']);

/** Members of a workbook that give one of its sheets. */
const WORKBOOK_SHEETS: ReadonlySet<string> = new Set(['sheets', 'worksheets', 'activesheet']);

/**
 * What an object taken from another becomes when that one ends, by the
 * owner's type and its own (issues #294 and #683, measured in Excel, Word
 * and PowerPoint 16.0): a Range of a deleted sheet or of a closed
 * workbook's sheet raises 424, the sheet itself the Worksheet's error; a
 * Range of a closed document 5825; a Slide of a closed presentation the
 * Slide's error. `chain` is the member names after the owner, `ws.Range`
 * as ['range'].
 */
function derivedEnding(ownerType: string, type: string, chain: readonly string[]): { kind: string; error: string } | undefined {
	if (ownerType === 'worksheet' && type === 'range' && SHEET_RANGES.has(chain[0])) {
		return { kind: 'a range', error: "'424': Object required" };
	}
	if (ownerType === 'workbook' && WORKBOOK_SHEETS.has(chain[0])) {
		if (type === 'worksheet' && chain.length === 1) {
			return { kind: 'a sheet', error: ENDINGS.worksheet.error };
		}
		if (type === 'range' && chain.length === 2 && SHEET_RANGES.has(chain[1])) {
			return { kind: 'a range', error: "'424': Object required" };
		}
	}
	if (ownerType === 'document' && type === 'range') {
		return { kind: 'a range', error: ENDINGS.document.error };
	}
	if (ownerType === 'presentation' && type === 'slide' && chain[0] === 'slides') {
		return { kind: 'a slide', error: ENDINGS.slide.error };
	}
	return undefined;
}

interface Ended {
	/** How it ended, for the message: "deleted on line 7". */
	how: string;
	error: string;
}

interface Derived {
	owner: string;
	/** How the message names it: "a range". */
	kind: string;
	error: string;
}

export function checkDeletedObjects(
	source: string,
	mod: ModuleNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		const locals = (procedureSymbolFor(symbols, member)?.children ?? []).filter((child) => child.kind === 'localVariable' && child.visibility !== 'Static');
		const typeOf = new Map(locals.filter((child) => !child.isArray).map((child) => [child.name.toLowerCase(), normalizeType(child.asType) ?? 'variant']));
		if (![...typeOf.values()].some((type) => ENDINGS[type] || type === 'range' || type === 'variant')) {
			continue;
		}
		const run = (body: readonly BodyNode[]): void => {
			const ended = new Map<string, Ended>();
			const rangesOf = new Map<string, Derived>();
			const forget = (lower: string): void => {
				ended.delete(lower);
			};
			for (const node of body) {
				if (isInactiveNode(activity, node)) {
					continue;
				}
				if (!isLeafStatement(node)) {
					// A block may change what it names; its own run starts afresh.
					const text = source.slice(node.span.start, node.span.end).toLowerCase();
					for (const lower of [...ended.keys(), ...rangesOf.keys()]) {
						if (new RegExp(`\\b${lower}\\b`).test(text)) {
							forget(lower);
							rangesOf.delete(lower);
						}
					}
					if ('body' in node && Array.isArray(node.body)) {
						run(node.body as BodyNode[]);
					}
					continue;
				}
				// A label is reached from wherever a jump to it runs: an error
				// handler from any line after its On Error GoTo, the Delete's
				// included (issue #584). Nothing known before it holds there.
				if (statementLabelDeclaration(source, node.span)) {
					ended.clear();
					rangesOf.clear();
				}
				const toks = statementTokensAfterLeadingLabel(source, node.span).filter((tok) => tok.kind !== 'comment');
				const line = source.slice(0, node.span.start).split('\n').length;
				const head = tokenText(toks[0]);
				// `Set x = ...` gives x a new object.
				if (head === 'set' && tokenName(toks[1]) && toks[2]?.rawText === '=') {
					const target = toks[1].rawText.toLowerCase();
					forget(target);
					const owner = tokenName(toks[3])?.toLowerCase();
					const derived = owner ? derivedEnding(typeOf.get(owner) ?? '', typeOf.get(target) ?? '', memberChain(toks, 4)) : undefined;
					if (owner && derived) {
						rangesOf.set(target, { owner, ...derived });
					} else {
						rangesOf.delete(target);
					}
					continue;
				}
				// `ws.Delete`, `wb.Close False`, `lo.Unlist`.
				const subject = tokenName(toks[0])?.toLowerCase();
				const ending = subject ? ENDINGS[typeOf.get(subject) ?? ''] : undefined;
				if (subject && ending && toks[1]?.rawText === '.' && tokenText(toks[2]) === ending.verb && (toks.length === 3 || ending.verb === 'close')) {
					const how = `${ending.verb === 'close' ? 'closed' : ending.verb === 'unlist' ? 'unlisted' : 'deleted'} on line ${line}`;
					ended.set(subject, { how, error: ending.error });
					// What was taken from it, and from that in turn: a sheet of a
					// closed workbook and a range of that sheet.
					const endFrom = (owner: string, shown: string): void => {
						for (const [taken, derived] of rangesOf) {
							if (derived.owner === owner && !ended.has(taken)) {
								ended.set(taken, { how: `${derived.kind} of '${shown}', which was ${how}`, error: derived.error });
								endFrom(taken, shown);
							}
						}
					};
					endFrom(subject, toks[0].rawText);
					continue;
				}
				report(toks, node.span.start, ended, push);
				// What the statement assigns or passes whole is no longer known.
				const assigned = tokenName(toks[head === 'let' ? 1 : 0])?.toLowerCase();
				const assignAt = head === 'let' ? 2 : 1;
				if (assigned && toks[assignAt]?.rawText === '=') {
					forget(assigned);
				}
				toks.forEach((tok, i) => {
					const lower = tokenName(tok)?.toLowerCase();
					if (lower && (toks[i - 1]?.rawText === '(' || toks[i - 1]?.rawText === ',') && (toks[i + 1]?.rawText === ')' || toks[i + 1]?.rawText === ',' || !toks[i + 1])) {
						forget(lower);
					}
				});
			}
		};
		run(member.body);
	}
}

/** The member names of the chain starting at `.` at `from`, arguments skipped: `.Sheets(1).Range("A1")` gives ['sheets', 'range']. */
function memberChain(toks: readonly VbaToken[], from: number): string[] {
	const out: string[] = [];
	let i = from;
	while (toks[i]?.rawText === '.' && tokenName(toks[i + 1])) {
		out.push(tokenText(toks[i + 1]));
		i += 2;
		if (toks[i]?.rawText === '(') {
			const close = matchParenFrom(toks, i);
			if (close < 0) {
				return [];
			}
			i = close + 1;
		}
	}
	return i === toks.length ? out : [];
}

function report(toks: readonly VbaToken[], start: number, ended: Map<string, Ended>, push: PushFn): void {
	toks.forEach((tok, i) => {
		const lower = tokenName(tok)?.toLowerCase();
		if (!lower || toks[i - 1]?.rawText === '.' || toks[i - 1]?.rawText === '!') {
			return;
		}
		const at = { start: start + tok.start, end: start + tok.end };
		const gone = ended.get(lower);
		if (gone && toks[i + 1]?.rawText === '.' && tokenName(toks[i + 2])) {
			const name = toks[i + 2].rawText;
			const what = /^a (?:range|sheet|slide) /.test(gone.how) ? `'${tok.rawText}' is ${gone.how}` : `'${tok.rawText}' was ${gone.how}`;
			push('objectUsedAfterDelete', `${what}, so its ${name} is gone. This will raise Run-time error ${gone.error.replace('MEMBER', name)}.`, at);
			ended.delete(lower);
		}
	});
}
