// Where a statement reads a bare name as a truth value or a Boolean operand
// (issue #424): the whole condition of If, ElseIf, Do While/Until, Loop
// While/Until and While, a Select Case subject, IIf's first argument, and an
// operand of Not, And, Or, Xor, Eqv or Imp. An object, an array or a Variant
// holding an array has no value there, and what each raises was measured in
// Excel 16.0; the rules that know each kind of name judge it.

import type { VbaToken } from '../lexer/tokenKinds';
import { tokenName, tokenText } from './walker';

export type ConditionForm = 'condition' | 'select' | 'iif' | 'not' | 'logical';

export interface ConditionOperand {
	index: number;
	form: ConditionForm;
}

const LOGICAL: ReadonlySet<string> = new Set(['and', 'or', 'xor', 'eqv', 'imp']);
const CONDITION_HEADS: ReadonlyArray<readonly string[]> = [
	['do', 'while'], ['do', 'until'], ['loop', 'while'], ['loop', 'until'], ['while'], ['select', 'case'],
];

/** The bare names a statement's tokens read as a condition or a Boolean operand. */
export function conditionOperands(toks: readonly VbaToken[]): ConditionOperand[] {
	const out: ConditionOperand[] = [];
	const bare = (i: number): boolean => tokenName(toks[i]) !== undefined && toks[i].kind !== 'keyword'
		&& !['.', '!'].includes(toks[i - 1]?.rawText ?? '') && !['(', '.', '!', '$'].includes(toks[i + 1]?.rawText ?? '');
	const head = tokenText(toks[0]);
	// `If x Then` and `ElseIf x Then`, a block's or a one-line If's.
	if ((head === 'if' || head === 'elseif') && tokenText(toks[2]) === 'then' && bare(1)) {
		out.push({ index: 1, form: 'condition' });
	}
	for (const words of CONDITION_HEADS) {
		if (toks.length === words.length + 1 && words.every((word, k) => tokenText(toks[k]) === word) && bare(words.length)) {
			out.push({ index: words.length, form: words[0] === 'select' ? 'select' : 'condition' });
		}
	}
	// An operand ends where the expression does or another Boolean operator starts.
	const ends = (tok: VbaToken | undefined): boolean => !tok || [')', ','].includes(tok.rawText) || tokenText(tok) === 'then' || LOGICAL.has(tokenText(tok));
	const starts = (tok: VbaToken | undefined): boolean => !tok || ['(', ',', '='].includes(tok.rawText)
		|| ['then', 'if', 'elseif', 'while', 'until', 'case', 'not', 'else'].includes(tokenText(tok)) || LOGICAL.has(tokenText(tok));
	for (let i = 1; i < toks.length; i++) {
		if (!bare(i) || out.some((hit) => hit.index === i)) {
			continue;
		}
		if (tokenText(toks[i - 1]) === 'not' && ends(toks[i + 1])) {
			out.push({ index: i, form: 'not' });
		} else if ((LOGICAL.has(tokenText(toks[i - 1])) && ends(toks[i + 1])) || (LOGICAL.has(tokenText(toks[i + 1])) && starts(toks[i - 1]))) {
			out.push({ index: i, form: 'logical' });
		} else if (toks[i - 1]?.rawText === '(' && tokenText(toks[i - 2]) === 'iif' && toks[i + 1]?.rawText === ',' && toks[i - 3]?.rawText !== '.') {
			out.push({ index: i, form: 'iif' });
		}
	}
	return out;
}
