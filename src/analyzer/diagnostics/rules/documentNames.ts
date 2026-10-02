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
		const forget = (names: Iterable<string>): void => {
			const gone = new Set(names);
			for (const key of [...state.keys()]) {
				const head = key.startsWith(SLIDE) ? key.slice(SLIDE.length) : key.split('.')[0];
				if (gone.has(head)) {
					state.delete(key);
				}
			}
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
			}
			const set = setAssignmentTarget(source, node.span);
			if (set) {
				const lower = set.name.toLowerCase();
				forget(namesIn(source, node.span));
				const value = set.valueTokens.filter((tok) => tok.kind !== 'comment');
				const call = value.findIndex((tok, k) => value[k - 1]?.rawText === '.' && value[k - 2] && tokenText(value[k - 2]) === 'slides' && ['add', 'addslide'].includes(tokenText(tok)));
				if (host === 'PowerPoint' && call > 0 && value[call + 1]?.rawText === '(' && matchParenFrom(value, call + 1) === value.length - 1) {
					state.set(SLIDE + lower, new Set());
				}
				return;
			}
			// `a.Name = "Zq"` on a slide the procedure added.
			const slide = toks.length === 5 && toks[1].rawText === '.' && tokenText(toks[2]) === 'name' && toks[3].rawText === '=' ? tokenName(toks[0])?.toLowerCase() : undefined;
			if (slide !== undefined && state.has(SLIDE + slide)) {
				const name = toks[4].kind === 'stringLiteral' ? stringLiteralValue(toks[4].rawText) : undefined;
				const lower = name?.toLowerCase();
				const taken = lower ? [...state].find(([key, held]) => key.startsWith(SLIDE) && key !== SLIDE + slide && held.has(lower)) : undefined;
				if (taken) {
					push('hostArgumentOutOfRange', `"${name}" is the name the code gave '${taken[0].slice(SLIDE.length)}', another slide it added, and slide names ignore case. This will raise Run-time error '-2147188160': Another slide already has this name.`, at(toks[4]));
				}
				state.set(SLIDE + slide, new Set(lower ? [lower] : []));
				return;
			}
			forget(namesIn(source, node.span));
		};
		walkEnteringBlocks(source, member.body, (node) => activity?.isInactive(node.span) === true, visit, {
			snapshot: () => new Map([...state].map(([key, held]) => [key, new Set(held)])),
			restore: (saved) => {
				state.clear();
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
