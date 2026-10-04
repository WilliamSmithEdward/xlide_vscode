// What a UserForm's designer puts in its controls, against what the form's
// own code asks of them (issue #315, each case measured in Excel 16.0):
//
//   Mp.Pages(9) and Mp.Pages("Page9") on a two-page MultiPage   error 5
//   Mp.Value = 5 on the same MultiPage                          error 380
//   L1.AddItem "a" then L1.ListIndex = 5 on an unbound ListBox  error 380
//   L1.Selected(5) = True after the same AddItem                error 380
//   L1.List(5) read after the same AddItem                      error 381
//   L1.ListIndex = L1.ListCount, whatever the list holds        error 380
//
// The designer knows a MultiPage's pages and whether a list has a RowSource.
// Code anywhere can add pages or items, so a control is judged only inside
// the one procedure of its form that names it: no other procedure or module
// names it, and nothing in the project reaches a form's Controls collection.

import type { ConditionalActivityTracker } from '../../conditional/conditionalCompilation';
import { statementLabelDeclaration } from '../../flow/procedureLabels';
import { matchParenFrom, splitTopLevelTokenGroups, identifierWords } from '../../lexer/tokenHelpers';
import { tokenizeCached } from '../../lexer/tokenize';
import type { VbaToken } from '../../lexer/tokenKinds';
import type { BodyNode, ModuleNode, ProcedureNode, Span } from '../../parser/nodes';
import { isLeafStatement } from '../../parser/nodes';
import type { FormControlInfo } from '../../symbols/projectIndex';
import type { PushFn } from '../analysisContext';
import { stringLiteralValue } from '../typeInference';
import { activeModuleMembers, blockHeaderLineSpan, isInactiveNode, statementAndBranchSpans, statementTokens, tokenName, tokenText } from '../walker';

/** A list or MultiPage the designer knows the contents of, by lowercased name. */
type Judged = ReadonlyMap<string, FormControlInfo>;

/** Members of a list control that read it without changing what it holds. */
const LIST_READS: ReadonlySet<string> = new Set([
	'listcount', 'listindex', 'value', 'text', 'name', 'visible', 'enabled', 'tag', 'locked',
	'top', 'left', 'width', 'height', 'setfocus', 'boundcolumn', 'textcolumn', 'multiselect',
]);

export function checkFormContents(
	source: string,
	mod: ModuleNode,
	controls: readonly FormControlInfo[] | undefined,
	nameMentions: ReadonlyMap<string, number> | undefined,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	if (!controls || controls.length === 0) {
		return;
	}
	const lists = new Map<string, FormControlInfo>();
	const pageSets = new Map<string, FormControlInfo>();
	for (const control of controls) {
		const lower = control.name.toLowerCase();
		if ((control.type === 'MSForms.ListBox' || control.type === 'MSForms.ComboBox')) {
			lists.set(lower, control);
		} else if (control.type === 'MSForms.MultiPage' && control.pages) {
			pageSets.set(lower, control);
		}
	}
	if (lists.size === 0 && pageSets.size === 0) {
		return;
	}
	// Code that reaches a form's Controls, here or in any other module, can
	// reach every control without naming it.
	const controlsReached = (nameMentions?.get('controls') ?? 0) > 0;
	const where = mentionSpans(source, new Set([...lists.keys(), ...pageSets.keys()]));
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		const ownOnly = (lower: string): boolean => nameMentions !== undefined && !controlsReached
			&& (nameMentions.get(lower) ?? 0) === 1
			&& (where.get(lower) ?? []).every((at) => at >= member.span.start && at < member.span.end)
			&& !declaresName(member, lower);
		checkListAgainstCount(source, member, lists, ownOnly, activity, push);
		checkPages(source, member, pageSets, ownOnly, activity, push);
	}
}

/** Where the module names each of `names`, as an identifier or a word in a string. */
function mentionSpans(source: string, names: ReadonlySet<string>): Map<string, number[]> {
	const out = new Map<string, number[]>();
	for (const token of tokenizeCached(source)) {
		const words = token.kind === 'identifier' ? [token.rawText.toLowerCase()]
			: token.kind === 'stringLiteral' ? identifierWords(token.rawText) : [];
		for (const word of words) {
			if (names.has(word)) {
				out.set(word, [...(out.get(word) ?? []), token.start]);
			}
		}
	}
	return out;
}

/** Whether the procedure declares a parameter or local of this name, which hides the control. */
function declaresName(member: ProcedureNode, lower: string): boolean {
	if (member.params.some((param) => param.name.toLowerCase() === lower)) {
		return true;
	}
	let found = false;
	const visit = (list: readonly BodyNode[]): void => {
		for (const node of list) {
			if (node.kind === 'VariableGroup' && node.declarations.some((decl) => decl.name.toLowerCase() === lower)) {
				found = true;
			}
			if ('body' in node && Array.isArray(node.body)) {
				visit(node.body as BodyNode[]);
			}
			if (node.kind === 'IfBlock') {
				for (const branch of node.branches) {
					visit(branch.body);
				}
			}
		}
	};
	visit(member.body);
	return found;
}

/**
 * The control a token names: `L1` or `Me.L1`, never `f.L1`, which is another
 * instance's. Undefined for any other token.
 */
function namedControl(toks: readonly VbaToken[], i: number, controls: Judged): FormControlInfo | undefined {
	const lower = tokenName(toks[i])?.toLowerCase();
	if (!lower || toks[i].kind !== 'identifier' || !controls.has(lower)) {
		return undefined;
	}
	if (toks[i - 1]?.rawText === '.' && !(tokenText(toks[i - 2]) === 'me' && toks[i - 3]?.rawText !== '.')) {
		return undefined;
	}
	return controls.get(lower);
}

/** A whole-number literal, a minus sign allowed. */
function signedLiteral(group: readonly VbaToken[]): number | undefined {
	const toks = group.filter((tok) => tok.kind !== 'comment');
	const negative = toks.length === 2 && toks[0].rawText === '-';
	const tok = toks[negative ? 1 : 0];
	if (!tok || toks.length !== (negative ? 2 : 1) || tok.kind !== 'integerLiteral' || !/^\d+$/.test(tok.rawText)) {
		return undefined;
	}
	return negative ? -Number(tok.rawText) : Number(tok.rawText);
}

/** Whether the procedure has a label or a jump, which can run a line twice. */
function jumps(source: string, member: ProcedureNode): boolean {
	let found = false;
	const visit = (list: readonly BodyNode[]): void => {
		for (const node of list) {
			if (isLeafStatement(node)) {
				const head = tokenText(statementTokens(source, node.span)[0]);
				if (statementLabelDeclaration(source, node.span) || head === 'goto' || head === 'gosub' || head === 'resume' || head === 'on') {
					found = true;
				}
			} else if ('body' in node && Array.isArray(node.body)) {
				visit(node.body as BodyNode[]);
			}
			if (node.kind === 'IfBlock') {
				for (const branch of node.branches) {
					visit(branch.body);
				}
			}
		}
	};
	visit(member.body);
	return found;
}

/**
 * Follows each unbound list's item count down the procedure's top level,
 * from empty: AddItem adds one, Clear empties it, `List = Array(...)` sets
 * it. A line that may change it any other way, or a block that names it,
 * ends what is known.
 */
function checkListAgainstCount(
	source: string,
	member: ProcedureNode,
	lists: Judged,
	ownOnly: (lower: string) => boolean,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	if (lists.size === 0) {
		return;
	}
	const counts = new Map<string, number | undefined>();
	for (const [lower, control] of lists) {
		counts.set(lower, control.listStartsEmpty === true && ownOnly(lower) ? 0 : undefined);
	}
	const judged = !jumps(source, member);
	const visit = (node: BodyNode, topLevel: boolean): void => {
		if (isInactiveNode(activity, node)) {
			return;
		}
		if (!isLeafStatement(node)) {
			// A block may run its lines any number of times.
			for (const span of [node.span]) {
				const toks = statementTokens(source, span);
				for (let i = 0; i < toks.length; i++) {
					const control = namedControl(toks, i, lists);
					if (control) {
						counts.set(control.name.toLowerCase(), undefined);
					}
				}
			}
			const children = node.kind === 'IfBlock' ? node.branches.flatMap((branch) => branch.body) : 'body' in node && Array.isArray(node.body) ? node.body as BodyNode[] : [];
			for (const child of children) {
				visit(child, false);
			}
			return;
		}
		const spans = statementAndBranchSpans(node);
		for (const span of spans) {
			const toks = statementTokens(source, span);
			listStatement(toks, span, lists, judged && topLevel && spans.length === 1 ? counts : new Map(), counts, push);
		}
	};
	for (const node of member.body) {
		visit(node, true);
	}
}

/**
 * One statement against the counts known before it: reports what it asks
 * past the list, then moves `counts` on. `known` is empty where the
 * statement may run other than once, in order.
 */
function listStatement(
	toks: readonly VbaToken[],
	span: Span,
	lists: Judged,
	known: ReadonlyMap<string, number | undefined>,
	counts: Map<string, number | undefined>,
	push: PushFn,
): void {
	const at = (from: VbaToken, to: VbaToken = from): Span => ({ start: span.start + from.start, end: span.start + to.end });
	for (let i = 0; i < toks.length; i++) {
		const control = namedControl(toks, i, lists);
		if (!control) {
			continue;
		}
		const lower = control.name.toLowerCase();
		// `L1.ListIndex = L1.ListCount` is past any list (measured).
		const member = toks[i + 1]?.rawText === '.' ? tokenText(toks[i + 2]) : undefined;
		const statementStart = tokenText(toks[0]) === 'me' ? 2 : 0;
		const target = i === statementStart;
		if (target && member === 'listindex' && toks[i + 3]?.rawText === '=' && toks.length >= i + 7) {
			const rest = toks.slice(i + 4);
			const restStart = tokenText(rest[0]) === 'me' && rest[1]?.rawText === '.' ? 2 : 0;
			if (rest.length === restStart + 3 && tokenName(rest[restStart])?.toLowerCase() === lower && rest[restStart + 1].rawText === '.' && tokenText(rest[restStart + 2]) === 'listcount') {
				push('hostPropertyValueOutOfRange', `ListIndex runs from -1 to ${control.name}.ListCount - 1, so ${control.name}.ListCount is past the list. This will raise Run-time error '380': Could not set the ListIndex property. Invalid property value.`, at(toks[i], toks[toks.length - 1]));
				return;
			}
		}
		const count = known.get(lower);
		if (!target) {
			// Read in an expression: `L1.List(5)` past the list raises 381.
			if (member === 'list' && toks[i + 3]?.rawText === '(' && count !== undefined) {
				const close = matchParenFrom(toks, i + 3);
				const args = close > 0 ? splitTopLevelTokenGroups(toks, i + 4, ',', close) : [];
				const index = args.length === 1 ? signedLiteral(args[0]) : undefined;
				if (index !== undefined && (index < 0 || index >= count)) {
					push('hostPropertyValueOutOfRange', `${control.name} ${holds(count)} here, so List(${index}) is past it. This will raise Run-time error '381': Could not get the List property. Invalid property array index.`, at(toks[i], toks[close]));
				}
				continue;
			}
			if (member === undefined || !LIST_READS.has(member) && member !== 'list' && member !== 'column' && member !== 'selected') {
				counts.set(lower, undefined);
			}
			continue;
		}
		const args = splitTopLevelTokenGroups(toks, i + 3, ',');
		switch (member) {
			case 'additem':
				counts.set(lower, count === undefined ? undefined : count + 1);
				break;
			case 'clear':
				counts.set(lower, counts.get(lower) === undefined ? undefined : 0);
				break;
			case 'removeitem': {
				const index = args.length === 1 ? signedLiteral(args[0]) : undefined;
				counts.set(lower, count !== undefined && index !== undefined && index >= 0 && index < count ? count - 1 : undefined);
				break;
			}
			case 'list': {
				// `L1.List = Array("a", "b")` holds two.
				const value = toks.slice(i + 4);
				const close = tokenText(value[0]) === 'array' && value[1]?.rawText === '(' ? matchParenFrom(value, 1) : -1;
				const items = close === value.length - 1 && toks[i + 3]?.rawText === '=' ? splitTopLevelTokenGroups(value, 2, ',', close) : undefined;
				counts.set(lower, items && counts.get(lower) !== undefined ? (close === 2 ? 0 : items.length) : undefined);
				break;
			}
			case 'listindex': {
				const value = toks[i + 3]?.rawText === '=' ? signedLiteral(toks.slice(i + 4)) : undefined;
				if (count !== undefined && value !== undefined && (value < -1 || value >= count)) {
					push('hostPropertyValueOutOfRange', `${control.name} ${holds(count)} here, so ListIndex runs ${count === 0 ? 'only to -1' : `from -1 to ${count - 1}`}; ${value} is outside it. This will raise Run-time error '380': Could not set the ListIndex property. Invalid property value.`, at(toks[i], toks[toks.length - 1]));
				}
				break;
			}
			case 'selected': {
				const close = toks[i + 3]?.rawText === '(' ? matchParenFrom(toks, i + 3) : -1;
				const index = close > 0 ? signedLiteral(toks.slice(i + 4, close)) : undefined;
				if (control.type === 'MSForms.ListBox' && count !== undefined && index !== undefined && (index < 0 || index >= count) && toks[close + 1]?.rawText === '=') {
					push('hostPropertyValueOutOfRange', `${control.name} ${holds(count)} here, so Selected(${index}) is past it. This will raise Run-time error '380': Could not set the Selected property. Invalid property value.`, at(toks[i], toks[close]));
				}
				break;
			}
			default:
				if (member === undefined || !LIST_READS.has(member)) {
					counts.set(lower, undefined);
				}
		}
	}
}

function holds(count: number): string {
	return count === 0 ? 'holds no items' : `holds ${count} item${count === 1 ? '' : 's'}`;
}

/**
 * `Mp.Pages(9)`, `Mp.Pages("Page9")` and `Mp.Value = 5` on a MultiPage
 * whose pages the designer lists, in a procedure that cannot have changed
 * them: it names the MultiPage only as `Mp.Pages(...)`, `Mp.Pages.Count`
 * or one of its own properties.
 */
function checkPages(
	source: string,
	member: ProcedureNode,
	pageSets: Judged,
	ownOnly: (lower: string) => boolean,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	const candidates = new Map([...pageSets].filter(([lower]) => ownOnly(lower)));
	if (candidates.size === 0) {
		return;
	}
	const uses: Array<{ control: FormControlInfo; toks: readonly VbaToken[]; i: number; span: Span }> = [];
	const unsure = new Set<string>();
	const visit = (list: readonly BodyNode[]): void => {
		for (const node of list) {
			if (isInactiveNode(activity, node)) {
				continue;
			}
			// A block's opening line names what it reads like any statement;
			// a With over the MultiPage can then reach it unnamed.
			const spans = isLeafStatement(node) ? statementAndBranchSpans(node) : [blockHeaderLineSpan(source, node.span)];
			for (const span of spans) {
				const toks = statementTokens(source, span);
				const withHeader = !isLeafStatement(node) && tokenText(toks[0]) === 'with';
				for (let i = 0; i < toks.length; i++) {
					const control = namedControl(toks, i, candidates);
					if (!control) {
						continue;
					}
					const next = toks[i + 1]?.rawText === '.' ? tokenText(toks[i + 2]) : undefined;
					if (withHeader || next === undefined || (next === 'pages' && toks[i + 3]?.rawText !== '(' && !(toks[i + 3]?.rawText === '.' && tokenText(toks[i + 4]) === 'count'))) {
						unsure.add(control.name.toLowerCase());
					} else {
						uses.push({ control, toks, i, span });
					}
				}
			}
			if (!isLeafStatement(node)) {
				const children = node.kind === 'IfBlock' ? node.branches.flatMap((branch) => branch.body) : 'body' in node && Array.isArray(node.body) ? node.body as BodyNode[] : [];
				visit(children);
			}
		}
	};
	visit(member.body);
	for (const { control, toks, i, span } of uses) {
		if (unsure.has(control.name.toLowerCase())) {
			continue;
		}
		const pages = control.pages!;
		const at = (tok: VbaToken): Span => ({ start: span.start + tok.start, end: span.start + tok.end });
		const next = tokenText(toks[i + 2]);
		if (next === 'pages' && toks[i + 3]?.rawText === '(') {
			const close = matchParenFrom(toks, i + 3);
			const arg = close > 0 ? toks.slice(i + 4, close) : [];
			const index = signedLiteral(arg);
			if (index !== undefined && (index < 0 || index >= pages.length)) {
				push('runtimeArgumentValue', `The MultiPage ${control.name} has ${pages.length} page${pages.length === 1 ? '' : 's'}, indexed 0 to ${pages.length - 1}; ${index} is none of them. This will raise Run-time error '5': Invalid procedure call or argument.`, at(arg[arg.length - 1]));
			} else if (arg.length === 1 && arg[0].kind === 'stringLiteral') {
				const name = stringLiteralValue(arg[0].rawText);
				if (!pages.some((page) => page.toLowerCase() === name.toLowerCase())) {
					push('runtimeArgumentValue', `The MultiPage ${control.name} has no page named "${name}". This will raise Run-time error '5': Invalid procedure call or argument.`, at(arg[0]));
				}
			}
		} else if (next === 'value' && toks[i + 3]?.rawText === '=' && (i === 0 || (i === 2 && tokenText(toks[0]) === 'me'))) {
			const value = signedLiteral(toks.slice(i + 4));
			if (value !== undefined && value >= pages.length) {
				push('hostPropertyValueOutOfRange', `The MultiPage ${control.name} has ${pages.length} page${pages.length === 1 ? '' : 's'}, so Value runs from 0 to ${pages.length - 1}; ${value} is past it. This will raise Run-time error '380': Could not set the Value property. Invalid property value.`, { start: span.start + toks[i].start, end: span.start + toks[toks.length - 1].end });
			}
		}
	}
}
