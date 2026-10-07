// Only collection-independent failures are diagnosed for form contents.
// Designer pages and initial list contents cannot establish runtime bounds:
// instances persist, events run, and other code can change the collections.
import type { ConditionalActivityTracker } from '../../conditional/conditionalCompilation';
import type { VbaToken } from '../../lexer/tokenKinds';
import type { BodyNode, ModuleNode, ProcedureNode } from '../../parser/nodes';
import type { FormControlInfo } from '../../symbols/projectIndex';
import type { PushFn } from '../analysisContext';
import { activeModuleMembers, forEachStatement, matchParenFrom, statementAndBranchSpans, statementTokens, tokenName, tokenText } from '../walker';

type Judged = ReadonlyMap<string, FormControlInfo>;

export function checkFormContents(
	source: string,
	mod: ModuleNode,
	controls: readonly FormControlInfo[] | undefined,
	_nameMentions: ReadonlyMap<string, number> | undefined,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	if (!controls?.length) { return; }
	const known = new Map(controls.map(control => [control.name.toLowerCase(), control]));
	for (const proc of activeModuleMembers(mod, activity)) {
		if (proc.kind !== 'Procedure') { continue; }
		const available = new Map([...known].filter(([name]) => !declaresName(proc, name)));
		forEachStatement(proc.body, stmt => {
			for (const span of statementAndBranchSpans(stmt)) {
				const toks = statementTokens(source, span);
				const start = tokenText(toks[0]) === 'me' && toks[1]?.rawText === '.' ? 2 : 0;
				for (let i = 0; i < toks.length; i++) {
					const control = namedControl(toks, i, available);
					if (!control || toks[i + 1]?.rawText !== '.') { continue; }
					const member = tokenText(toks[i + 2]);
					const list = control.type === 'MSForms.ListBox' || control.type === 'MSForms.ComboBox';
					const at = (end: number) => ({ start: span.start + toks[i].start, end: span.start + toks[end].end });
					if (list && i === start && member === 'listindex' && toks[i + 3]?.rawText === '=') {
						const rhs = toks.slice(i + 4);
						const r = tokenText(rhs[0]) === 'me' && rhs[1]?.rawText === '.' ? 2 : 0;
						const sameCount = rhs.length === r + 3 && tokenName(rhs[r])?.toLowerCase() === control.name.toLowerCase()
							&& rhs[r + 1]?.rawText === '.' && tokenText(rhs[r + 2]) === 'listcount';
						const value = signedLiteral(rhs);
						if (sameCount || value !== undefined && value < -1) {
							push('hostPropertyValueOutOfRange', 'ListIndex runs from -1 to ' + control.name + ".ListCount - 1; this value is outside it. This will raise Run-time error '380': Could not set the ListIndex property. Invalid property value.", at(toks.length - 1));
						}
					}
					// Negative row/page indexes cannot become valid as objects are added.
					if ((list && (member === 'list' || member === 'selected') || control.type === 'MSForms.MultiPage' && member === 'pages') && toks[i + 3]?.rawText === '(') {
						const close = matchParenFrom(toks, i + 3);
						const index = close > 0 ? signedLiteral(toks.slice(i + 4, close)) : undefined;
						if (index === undefined || index >= 0) { continue; }
						if (member === 'pages') {
							push('runtimeArgumentValue', "Pages indexes start at 0. This will raise Run-time error '5': Invalid procedure call or argument.", at(close));
						} else if (member === 'selected' && control.type === 'MSForms.ListBox' && toks[close + 1]?.rawText === '=') {
							push('hostPropertyValueOutOfRange', "Selected indexes start at 0. This will raise Run-time error '380': Could not set the Selected property. Invalid property value.", at(close));
						} else if (member === 'list' && i !== start) {
							push('hostPropertyValueOutOfRange', "List row indexes start at 0. This will raise Run-time error '381': Could not get the List property. Invalid property array index.", at(close));
						}
					}
				}
			}
		}, activity);
	}
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
