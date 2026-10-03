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
// presentation the procedure made with Presentations.Add, the slides it
// adds: each is named Slide1, Slide2 in the order it was added, so
// `a.Name = b.Name` and `a.Name = "Slide2"` raise -2147188160, as does
// `p.Slides(2).Name` given a name another slide has; and `p.Slides.Add 3`
// or `a.MoveTo 2` past the count raise -2147188160 too.
//
// What is known is followed in a straight line. Any other mention of the
// document or slide variable, a label, a block or a call into the
// project's own code ends it.

import type { ConditionalActivityTracker } from '../../conditional/conditionalCompilation';
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

/** A presentation the procedure made, and the slides it holds, in order. */
interface PresentationState {
	/** Slide state keys in order; undefined for a slide not followed. */
	order: (string | undefined)[];
	/** How many slides were added, which numbers the next default name. */
	added: number;
}

export function checkDocumentNames(
	source: string,
	mod: ModuleNode,
	host: string | undefined,
	callables: ReadonlySet<string>,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	if (host !== 'Word' && host !== 'PowerPoint') {
		return;
	}
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		// `d.variables` -> the names added; `#slide:a` -> the name slide a was given ('' none).
		const state = new Map<string, Set<string>>();
		// By presentation variable (issue #610).
		const presentations = new Map<string, PresentationState>();
		// The presentation each followed slide is in, by slide state key.
		const slideIn = new Map<string, string>();
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
		// The slide `p.Slides(2)` names, by its state key, or undefined.
		const slideAt = (pres: string, index: number): string | undefined => presentations.get(pres)?.order[index - 1];
		// A name another followed slide holds, in any case: its variable.
		const holder = (key: string, lower: string): string | undefined => [...state]
			.find(([other, held]) => other.startsWith(SLIDE) && other !== key && held.has(lower))?.[0].slice(SLIDE.length);
		const nameSlide = (key: string, name: string, nameAt: VbaToken, at: (tok: VbaToken) => { start: number; end: number }): void => {
			const lower = name.toLowerCase();
			const taken = holder(key, lower);
			if (taken !== undefined) {
				const who = taken.startsWith('#') ? 'a slide the code added' : `'${taken}', another slide the code added`;
				push('hostArgumentOutOfRange', `"${name}" is the name of ${who}, and slide names ignore case. This will raise Run-time error '-2147188160': Another slide already has this name.`, at(nameAt));
			}
			state.set(key, new Set([lower]));
		};
		// `Slides.Add(i, ...)` on a presentation the procedure made: a new
		// slide at i, named Slide<n> by the order it was added.
		const addSlide = (pres: string, index: number | undefined, key: string, indexAt: VbaToken | undefined, at: (tok: VbaToken) => { start: number; end: number }): void => {
			const held = presentations.get(pres);
			if (!held) {
				return;
			}
			if (index === undefined || index < 1) {
				presentations.delete(pres);
				return;
			}
			if (index > held.order.length + 1) {
				push('hostArgumentOutOfRange', `'${pres}' holds ${held.order.length} slide${held.order.length === 1 ? '' : 's'}, so a new one goes at 1 to ${held.order.length + 1}; ${index} is past that. This will raise Run-time error '-2147188160': Integer out of range.`, at(indexAt!));
				presentations.delete(pres);
				return;
			}
			held.added++;
			held.order.splice(index - 1, 0, key);
			state.set(key, new Set([`slide${held.added}`]));
			slideIn.set(key, pres);
		};
		const visit = (node: BodyNode): void => {
			if (!isLeafStatement(node)) {
				return;
			}
			const toks = statementTokensAfterLeadingLabel(source, node.span).filter((tok) => tok.kind !== 'comment');
			if (statementLabelDeclaration(source, node.span) || tokenText(toks[0]) === 'gosub'
				|| toks.some((tok) => tokenName(tok) !== undefined && callables.has(tokenText(tok)))) {
				state.clear();
			}
			if (node.kind === 'Statement' && node.singleLineIfBranches) {
				forget(namesIn(source, node.span));
				return;
			}
			const at = (tok: VbaToken) => ({ start: node.span.start + tok.start, end: node.span.start + tok.end });
			if (host === 'Word') {
				const add = wordAdd(toks);
				if (add && add.name === '' && add.key.endsWith('.customdocumentproperties')) {
					push('hostArgumentOutOfRange', `A custom property needs a name, and "" is none. This will raise Run-time error '-2147418113': Automation error.`, at(add.nameToken));
				}
				if (add) {
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
				if (host === 'PowerPoint' && call > 0 && value[call + 1]?.rawText === '(' && matchParenFrom(value, call + 1) === value.length - 1) {
					state.set(SLIDE + lower, new Set());
					// `p.Slides.Add(1, ...)` on a presentation the procedure made.
					if (pres && presentations.has(pres)) {
						const index = value[call + 2];
						addSlide(pres, index?.kind === 'integerLiteral' ? Number(index.rawText) : undefined, SLIDE + lower, index, at);
					}
				}
				// `Set p = Presentations.Add(...)`: a new presentation, no slides.
				if (host === 'PowerPoint' && tokenText(value[0]) === 'presentations' && value[1]?.rawText === '.' && tokenText(value[2]) === 'add'
					&& (value.length === 3 || (value[3]?.rawText === '(' && matchParenFrom(value, 3) === value.length - 1))) {
					presentations.set(lower, { order: [], added: 0 });
				}
				return;
			}
			if (host === 'PowerPoint') {
				// `p.Slides.Add 3, ppLayoutBlank` as a statement.
				const pres = tokenName(toks[0])?.toLowerCase();
				if (pres && presentations.has(pres) && toks[1]?.rawText === '.' && tokenText(toks[2]) === 'slides' && toks[3]?.rawText === '.' && tokenText(toks[4]) === 'add' && toks[5]?.rawText !== '(') {
					const index = toks[5];
					addSlide(pres, index?.kind === 'integerLiteral' ? Number(index.rawText) : undefined, `${SLIDE}#${node.span.start}`, index, at);
					return;
				}
				// `p.Slides(2).Name = "zq"`.
				if (pres && presentations.has(pres) && toks[1]?.rawText === '.' && tokenText(toks[2]) === 'slides' && toks[3]?.rawText === '(' && toks[4]?.kind === 'integerLiteral' && toks[5]?.rawText === ')'
					&& toks[6]?.rawText === '.' && tokenText(toks[7]) === 'name' && toks[8]?.rawText === '=' && toks.length === 10 && toks[9].kind === 'stringLiteral') {
					const key = slideAt(pres, Number(toks[4].rawText));
					if (key) {
						nameSlide(key, stringLiteralValue(toks[9].rawText), toks[9], at);
						return;
					}
				}
				// `a.MoveTo 2` past the slides of the presentation a is in.
				const moved = tokenName(toks[0])?.toLowerCase();
				const movedIn = moved ? slideIn.get(SLIDE + moved) : undefined;
				const count = movedIn ? presentations.get(movedIn)?.order.length : undefined;
				if (count !== undefined && toks[1]?.rawText === '.' && tokenText(toks[2]) === 'moveto' && toks.length === 4 && toks[3].kind === 'integerLiteral' && Number(toks[3].rawText) > count) {
					push('hostArgumentOutOfRange', `'${movedIn}' holds ${count} slide${count === 1 ? '' : 's'}, so ${toks[3].rawText} is past the last. This will raise Run-time error '-2147188160': Integer out of range.`, at(toks[3]));
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
			forget(namesIn(source, node.span));
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
			touches: (stmt) => namesIn(source, stmt.span),
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
