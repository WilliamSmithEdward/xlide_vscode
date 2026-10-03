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
import { activeModuleMembers, isInactiveNode, statementTokensAfterLeadingLabel, tokenName, tokenText } from '../walker';

/** What deletes or closes each type, and what a member of it then raises. */
const ENDINGS: Readonly<Record<string, { verb: string; error: string }>> = {
	worksheet: { verb: 'delete', error: "'-2147221080': Method 'MEMBER' of object '_Worksheet' failed" },
	workbook: { verb: 'close', error: "'-2147221080': Method 'MEMBER' of object '_Workbook' failed" },
	shape: { verb: 'delete', error: "'424': Object required" },
	name: { verb: 'delete', error: "'424': Object required" },
	listobject: { verb: 'unlist', error: "'1004': Application-defined or object-defined error" },
};

/** Members of a sheet that give a Range of it. */
const SHEET_RANGES: ReadonlySet<string> = new Set(['range', 'cells', 'rows', 'columns', 'usedrange']);

interface Ended {
	/** How it ended, for the message: "deleted on line 7". */
	how: string;
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
			const rangesOf = new Map<string, string>();
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
					if (typeOf.get(target) === 'range' && owner && typeOf.get(owner) === 'worksheet' && toks[4]?.rawText === '.' && SHEET_RANGES.has(tokenText(toks[5]))) {
						rangesOf.set(target, owner);
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
					if (typeOf.get(subject) === 'worksheet') {
						for (const [range, owner] of rangesOf) {
							if (owner === subject) {
								ended.set(range, { how: `a range of '${toks[0].rawText}', which was deleted on line ${line}`, error: "'424': Object required" });
							}
						}
					}
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
			const what = gone.how.startsWith('a range') ? `'${tok.rawText}' is ${gone.how}` : `'${tok.rawText}' was ${gone.how}`;
			push('objectUsedAfterDelete', `${what}, so its ${name} is gone. This will raise Run-time error ${gone.error.replace('MEMBER', name)}.`, at);
			ended.delete(lower);
		}
	});
}
