// Rule family: a name a Word document or a PowerPoint presentation already
// holds, given again (issue #311). Measured in Word and PowerPoint 16.0
// (2026-10-02); each compiles and raises every time it runs.
//
//  - Word: `d.Variables.Add "zq", 1` twice raises 5903, `d.Styles.Add "Zq"`
//    twice 5173, and `d.CustomDocumentProperties.Add Name:="zq", ...` twice
//    -2147467259. The names ignore case.
//  - PowerPoint: two slides the procedure added, `Set a = p.Slides.Add(...)`,
//    cannot share a name, and slide names ignore case: `a.Name = "Zq"` then
//    `b.Name = "zq"` raises -2147188160.
//
// Issue #610 adds, each measured: a variable added again after its value
// was written, or through another name for the document, `Set d2 = d`
// (5903); a custom property named "" (-2147418113); and, on a
// presentation the procedure made with Presentations.Add. Default names,
// index order and positive bounds are not inferred: new-slide events can
// change them before Add returns.
//
// What is known is followed in a straight line. Any other mention of the
// document or slide variable, a label, a block or a call into the
// project's own code ends it.

import type { ConditionalActivityTracker } from '../../conditional/conditionalCompilation';
import type { MemberCompletionContext } from '../../completion/memberAccess';
import { resolveReceiverTypeAt } from '../../completion/memberAccess';
import { statementLabelDeclaration } from '../../flow/procedureLabels';
import type { VbaToken } from '../../lexer/tokenKinds';
import type { BodyNode, ModuleNode } from '../../parser/nodes';
import { isLeafStatement } from '../../parser/nodes';
import type { PushFn } from '../analysisContext';
import { walkEnteringBlocks } from '../dataflow';
import { stringLiteralValue } from '../typeInference';
import {
	activeModuleMembers,
	matchParenFrom,
	setAssignmentTarget,
	statementTokensAfterLeadingLabel,
	tokenName,
	tokenText,
} from '../walker';
import { namesIn } from './shared';

/** Word collections whose Add refuses a name already in, by lowercased member. */
const WORD_NAMED: ReadonlyMap<string, { error: string; noun: string }> = new Map([
	['variables', { error: "'5903': The Variable name already exists", noun: 'variable' }],
	['styles', { error: "'5173': This style name already exists or is reserved for a built-in style", noun: 'style' }],
	['customdocumentproperties', { error: "'-2147467259': Automation error", noun: 'custom property' }],
]);

/** The state key of a slide the procedure added; a variable name holds no `#`. */
const SLIDE = '#slide:';

export function checkDocumentNames(
	source: string,
	mod: ModuleNode,
	host: string | undefined,
	callables: ReadonlySet<string>,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
	memberCtx: MemberCompletionContext,
): void {
	if (host !== 'Word' && host !== 'PowerPoint') {
		return;
	}
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		// Failed mutations can be skipped by an error handler.
		if (/\bon\s+error\b/i.test(source.slice(member.span.start, member.span.end))) { continue; }
		// `d.variables` -> the names added; `#slide:a` -> the name slide a was given ('' none).
		const state = new Map<string, Set<string>>();
		// By presentation variable (issue #610).
		const presentations = new Set<string>();
		// The presentation each followed slide is in, by slide state key.
		const slideIn = new Map<string, string>();
		const forgetAll = (): void => {
			state.clear();
			presentations.clear();
			slideIn.clear();
		};
		const forget = (names: Iterable<string>): void => {
			const gone = new Set(names);
			for (const key of [...state.keys()]) {
				const head = key.startsWith(SLIDE) ? key.slice(SLIDE.length) : key.split('.')[0];
				if (gone.has(head)) {
					state.delete(key);
					slideIn.delete(key);
				}
			}
			for (const lower of gone) {
				presentations.delete(lower);
			}
		};
		// Duplicate names require fixed slides in the same known presentation.
		const holder = (key: string, lower: string): string | undefined => [...state]
			.find(([other, held]) => other.startsWith(SLIDE) && other !== key
				&& slideIn.has(key) && slideIn.get(other) === slideIn.get(key) && held.has(lower))?.[0].slice(SLIDE.length);
		const nameSlide = (key: string, name: string, nameAt: VbaToken, at: (tok: VbaToken) => { start: number; end: number }): void => {
			const lower = name.toLowerCase();
			const taken = holder(key, lower);
			if (taken !== undefined) {
				const who = taken.startsWith('#') ? 'a slide the code added' : `'${taken}', another slide the code added`;
				push('hostArgumentOutOfRange', `"${name}" is the name of ${who}, and slide names ignore case. This will raise Run-time error '-2147188160': Another slide already has this name.`, at(nameAt));
			}
			state.set(key, new Set([lower]));
		};
		// New-slide events can change names and order before Add returns.
		const addSlide = (pres: string | undefined, key: string): void => {
			for (const [other] of state) {
				if (other.startsWith(SLIDE)) { state.set(other, new Set()); }
			}
			state.set(key, new Set());
			if (pres && presentations.has(pres)) { slideIn.set(key, pres); }
		};
		const visit = (node: BodyNode): void => {
			if (!isLeafStatement(node)) {
				return;
			}
			const toks = statementTokensAfterLeadingLabel(source, node.span);
			if (statementLabelDeclaration(source, node.span) || tokenText(toks[0]) === 'gosub'
				|| toks.some((tok) => tokenName(tok) !== undefined && callables.has(tokenText(tok)))) {
				forgetAll();
			}
			if (node.kind === 'Statement' && node.singleLineIfBranches) {
				forgetAll();
				return;
			}
			const at = (tok: VbaToken) => ({ start: node.span.start + tok.start, end: node.span.start + tok.end });
			if (host === 'Word') {
				const candidate = wordAdd(toks);
				const collection = toks.findIndex(tok => WORD_NAMED.has(tokenText(tok)));
				const actualDocument = collection > 0 && resolveReceiverTypeAt(source, node.span.start + toks[collection - 1].end, memberCtx) === 'Word.Document';
				const add = actualDocument ? candidate : undefined;
				if (add && add.name === '' && add.key.endsWith('.customdocumentproperties')) {
					push('hostArgumentOutOfRange', `A custom property needs a name, and "" is none. This will raise Run-time error '-2147418113': Automation error.`, at(add.nameToken));
				}
				if (add) {
					// Arguments can invoke helpers/getters before Add runs.
					const addIndex = toks.findIndex(tok => tokenText(tok) === 'add');
					if (toks.slice(addIndex + 1).some((tok, i, args) => tok.kind === 'identifier' && args[i + 1]?.rawText !== ':='
						&& !['msopropertytypeboolean', 'msopropertytypedate', 'msopropertytypefloat', 'msopropertytypenumber', 'msopropertytypestring'].includes(tokenText(tok)))) {
						forgetAll();
						return;
					}
					const held = state.get(add.key) ?? new Set<string>();
					const name = add.name.toLowerCase();
					if (held.has(name)) {
						const kind = WORD_NAMED.get(add.key.slice(add.key.lastIndexOf('.') + 1))!;
						push('hostArgumentOutOfRange', `The ${kind.noun} "${add.name}" was already added to ${add.chain}, and the names ignore case. This will raise Run-time error ${kind.error}.`, at(add.nameToken));
					}
					forget([...namesIn(source, node.span)].filter((lower) => lower !== add.key.split('.')[0]));
					held.add(name);
					state.set(add.key, held);
					return;
				}
				// `d.Variables("zq").Value = 3` changes the value, not the names;
				// `.Delete` takes the name out (issue #610).
				const use = variableUse(toks);
				if (use && state.has(use.key)) {
					const eq = toks.findIndex(tok => tok.rawText === '=');
					if (eq >= 0 && toks.slice(eq + 1).some(tok => tok.kind === 'identifier')) { forgetAll(); return; }
					if (use.deletes) {
						state.get(use.key)!.delete(use.name);
					}
					forget([...namesIn(source, node.span)].filter((lower) => lower !== use.key.split('.')[0]));
					return;
				}
			}
			const set = setAssignmentTarget(source, node.span);
			if (set) {
				const lower = set.name.toLowerCase();
				const value = set.valueTokens.filter((tok) => tok.kind !== 'comment');
				// `Set d2 = d`: another name for the document, whose names it
				// shares (issue #610).
				const alias = value.length === 1 ? tokenName(value[0])?.toLowerCase() : undefined;
				const shared = alias ? [...state].filter(([key]) => !key.startsWith(SLIDE) && key.split('.')[0] === alias) : [];
				const call = value.findIndex((tok, k) => value[k - 1]?.rawText === '.' && value[k - 2] && tokenText(value[k - 2]) === 'slides' && ['add', 'addslide'].includes(tokenText(tok)));
				// `Set a = p.Slides.Add(1, ...)` keeps what is known of p.
				const pres = call === 4 && value[1].rawText === '.' && value[3].rawText === '.' ? tokenName(value[0])?.toLowerCase() : undefined;
				forget([...namesIn(source, node.span)].filter((name) => name !== pres || name === lower));
				for (const [key, held] of shared) {
					state.set(key, held);
					state.set(lower + key.slice(alias!.length), held);
				}
				if (host === 'PowerPoint' && call > 0 && value[call + 1]?.rawText === '(' && matchParenFrom(value, call + 1) === value.length - 1
					&& resolveReceiverTypeAt(source, node.span.start + value[call - 1].end, memberCtx) === 'PowerPoint.Slides') {
					addSlide(pres, SLIDE + lower);
				}
				// `Set p = Presentations.Add(...)`: a new presentation, no slides.
				if (host === 'PowerPoint' && tokenText(value[0]) === 'presentations' && value[1]?.rawText === '.' && tokenText(value[2]) === 'add'
					&& resolveReceiverTypeAt(source, node.span.start + value[1].end, memberCtx) === 'PowerPoint.Presentations'
					&& (value.length === 3 || (value[3]?.rawText === '(' && matchParenFrom(value, 3) === value.length - 1))) {
					presentations.add(lower);
				}
				return;
			}
			if (host === 'PowerPoint') {
				// `p.Slides.Add 3, ppLayoutBlank` as a statement.
				const pres = tokenName(toks[0])?.toLowerCase();
				if (pres && presentations.has(pres) && toks[1]?.rawText === '.' && tokenText(toks[2]) === 'slides' && toks[3]?.rawText === '.' && tokenText(toks[4]) === 'add' && toks[5]?.rawText !== '(') {
					addSlide(pres, `${SLIDE}#${node.span.start}`);
					return;
				}

			}
			// `a.Name = "Zq"` or `a.Name = b.Name` on a slide the procedure added.
			const slide = toks[1]?.rawText === '.' && tokenText(toks[2]) === 'name' && toks[3]?.rawText === '=' ? tokenName(toks[0])?.toLowerCase() : undefined;
			if (slide !== undefined && state.has(SLIDE + slide) && (toks.length === 5 || toks.length === 7)) {
				const other = toks.length === 7 && toks[5].rawText === '.' && tokenText(toks[6]) === 'name' ? tokenName(toks[4])?.toLowerCase() : undefined;
				const otherName = other !== undefined ? [...(state.get(SLIDE + other) ?? [])][0] : undefined;
				const name = toks.length === 5 && toks[4].kind === 'stringLiteral' ? stringLiteralValue(toks[4].rawText) : otherName;
				if (name !== undefined) {
					nameSlide(SLIDE + slide, name, toks[4], at);
				} else {
					state.set(SLIDE + slide, new Set());
				}
				return;
			}
			forgetAll();
		};
		walkEnteringBlocks(source, member.body, (node) => activity?.isInactive(node.span) === true, visit, {
			// A block forgets the presentations; what it names it forgets too.
			snapshot: () => new Map([...state].map(([key, held]) => [key, new Set(held)])),
			restore: (saved) => {
				state.clear();
				presentations.clear();
				for (const [key, held] of saved) {
					state.set(key, new Set(held));
				}
			},
			forget,
			touches: (stmt) => new Set([...namesIn(source, stmt.span), ...presentations.keys(), ...[...state.keys()].map(key => key.startsWith(SLIDE) ? key.slice(SLIDE.length) : key.split('.')[0])]),
		});
	}
}

/**
 * `d.Variables("zq").Value = 3`, `d.Variables("zq") = 3` or
 * `d.Variables("zq").Delete` on a receiver of names and dots: the state key
 * and the name, lower case, and whether it deletes.
 */
function variableUse(toks: readonly VbaToken[]): { key: string; name: string; deletes: boolean } | undefined {
	const at = toks.findIndex((tok, k) => tokenText(tok) === 'variables' && toks[k - 1]?.rawText === '.' && toks[k + 1]?.rawText === '(' && toks[k + 2]?.kind === 'stringLiteral' && toks[k + 3]?.rawText === ')');
	if (at < 2) {
		return undefined;
	}
	let start = at - 2;
	while (start >= 2 && toks[start - 1].rawText === '.' && tokenName(toks[start - 2]) !== undefined) {
		start -= 2;
	}
	if (start !== 0 || tokenName(toks[0]) === undefined) {
		return undefined;
	}
	const rest = toks.slice(at + 4).map((tok) => tokenText(tok) || tok.rawText);
	const writes = (rest[0] === '.' && rest[1] === 'value' && rest[2] === '=') || rest[0] === '=';
	const deletes = rest.length === 2 && rest[0] === '.' && rest[1] === 'delete';
	if (!writes && !deletes) {
		return undefined;
	}
	const key = toks.slice(0, at + 1).map((tok) => tok.rawText).join('').toLowerCase();
	return { key, name: stringLiteralValue(toks[at + 2].rawText).toLowerCase(), deletes };
}

/**
 * `d.Variables.Add "zq", 1`, `d.Styles.Add("Zq")` or
 * `d.CustomDocumentProperties.Add Name:="zq", ...` with a literal name, on a
 * receiver of names and dots: the state key, the receiver as written, and
 * the name.
 */
function wordAdd(toks: readonly VbaToken[]): { key: string; chain: string; name: string; nameToken: VbaToken } | undefined {
	const add = toks.findIndex((tok, k) => tokenText(tok) === 'add' && toks[k - 1]?.rawText === '.' && WORD_NAMED.has(tokenText(toks[k - 2])));
	if (add < 3) {
		return undefined;
	}
	// The receiver: names and dots from the statement's start, or after `=`.
	let start = add - 2;
	while (start >= 2 && toks[start - 1].rawText === '.' && tokenName(toks[start - 2]) !== undefined) {
		start -= 2;
	}
	const before = toks[start - 1]?.rawText;
	if (start === add - 2 || (start !== 0 && before !== '=' && tokenText(toks[start - 1]) !== 'call')) {
		return undefined;
	}
	const open = toks[add + 1]?.rawText === '(' ? add + 1 : -1;
	const close = open > 0 ? matchParenFrom(toks, open) : toks.length;
	const args: VbaToken[][] = [[]];
	for (let k = open > 0 ? open + 1 : add + 1, depth = 0; k < close; k++) {
		const raw = toks[k].rawText;
		depth += raw === '(' ? 1 : raw === ')' ? -1 : 0;
		if (raw === ',' && depth === 0) {
			args.push([]);
			continue;
		}
		args[args.length - 1].push(toks[k]);
	}
	const named = args.find((arg) => arg[1]?.rawText === ':=' && tokenText(arg[0]) === 'name');
	const arg = named ? named.slice(2) : args[0][1]?.rawText === ':=' ? undefined : args[0];
	if (arg?.length !== 1 || arg[0].kind !== 'stringLiteral') {
		return undefined;
	}
	const chain = toks.slice(start, add - 1).map((tok) => tok.rawText).join('');
	return { key: chain.toLowerCase(), chain: chain.slice(0, chain.lastIndexOf('.')), name: stringLiteralValue(arg[0].rawText), nameToken: arg[0] };
}
