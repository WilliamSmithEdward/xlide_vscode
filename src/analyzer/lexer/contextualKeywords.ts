// Contextual keywords are keywords only inside the statement that makes them
// keywords (issue #86). Everywhere else they are names, and the VBE spells a
// variable called `text`, `binary` or `output` the way it was declared: it
// capitalizes Text in `Option Compare Text` and nowhere else.
//
//   Explicit, Base, Compare, Binary, Text   Option (Binary also in Open ... For Binary)
//   Lib, Alias, PtrSafe                     Declare
//   Step                                    For
//   Error                                   On Error, and the Error statement
//   Output, Append, Random, Read            Open (Read after Access or Lock)
//
// The lexer calls this once it has the whole token stream, so every consumer
// sees the same answer: a keyword token where the word is grammar, an
// identifier token where it is a name.

import type { VbaToken } from './tokenKinds';

/** The words, in the canonical spelling the lexer gives their keyword tokens. */
const STATEMENT_BOUND: ReadonlySet<string> = new Set([
	'Explicit', 'Base', 'Compare', 'Binary', 'Text',
	'Lib', 'Alias', 'PtrSafe',
	'Step',
	'Error',
	'Output', 'Append', 'Random', 'Read',
]);

/** Turns each statement-bound keyword standing outside its statement into an identifier. */
export function settleContextualKeywords(tokens: readonly VbaToken[]): void {
	let statementStart = 0;
	for (let i = 0; i < tokens.length; i++) {
		const token = tokens[i];
		if (token.kind === 'newline' || token.kind === 'colon') {
			statementStart = i + 1;
		} else if (token.kind === 'keyword' && STATEMENT_BOUND.has(token.canonicalText ?? '')) {
			// Only what comes before the word decides it, and the words are rare,
			// so the statement is gathered for them alone.
			const statement = tokens.slice(statementStart, i + 1).filter((t) => t.kind !== 'comment');
			if (!isKeywordAt(statement, statement.length - 1)) {
				token.kind = 'identifier';
				delete token.canonicalText;
			}
		}
	}
}

function isKeywordAt(statement: readonly VbaToken[], i: number): boolean {
	const prev = statement[i - 1];
	if (isMemberAccess(prev)) {
		return false;
	}
	const before = word(prev);
	switch (word(statement[i])) {
		case 'explicit':
		case 'base':
		case 'compare':
			return before === 'option';
		case 'text':
			return before === 'compare' && word(statement[i - 2]) === 'option';
		case 'binary':
			return (before === 'compare' && word(statement[i - 2]) === 'option')
				|| (before === 'for' && inOpenStatement(statement, i));
		case 'output':
		case 'append':
		case 'random':
			return before === 'for' && inOpenStatement(statement, i);
		case 'read':
			return (before === 'access' || before === 'lock') && inOpenStatement(statement, i);
		case 'ptrsafe':
		case 'lib':
		case 'alias':
			return inDeclareHeader(statement, i);
		case 'step':
			return inForHeader(statement, i) && endsOperand(prev);
		case 'error':
			return before === 'on'
				|| (before === 'local' && word(statement[i - 2]) === 'on')
				|| startsStatement(statement, i);
		default:
			return true;
	}
}

/** The word a token spells, lower-cased: its canonical text when it has one. */
function word(token: VbaToken | undefined): string {
	return (token?.canonicalText ?? token?.rawText ?? '').toLowerCase();
}

function isMemberAccess(token: VbaToken | undefined): boolean {
	return !!token && ((token.kind === 'punctuation' && token.rawText === '.') || (token.kind === 'operator' && token.rawText === '!'));
}

/** At the start of a statement: first, after a line number, or after a single-line If's Then or Else. */
function startsStatement(statement: readonly VbaToken[], i: number): boolean {
	if (i === 0) {
		return true;
	}
	const prev = statement[i - 1];
	if (i === 1 && prev.kind === 'integerLiteral' && /^\d+$/.test(prev.rawText)) {
		return true;
	}
	return prev.kind === 'keyword' && (word(prev) === 'then' || word(prev) === 'else');
}

/** A `Open` statement runs before `i`: the reserved word, not a member called Open. */
function inOpenStatement(statement: readonly VbaToken[], i: number): boolean {
	for (let j = i - 1; j >= 0; j--) {
		if (word(statement[j]) === 'open' && statement[j].kind === 'keyword' && !isMemberAccess(statement[j - 1])) {
			return true;
		}
	}
	return false;
}

/** Inside a Declare header, before its parameter list. */
function inDeclareHeader(statement: readonly VbaToken[], i: number): boolean {
	let declare = false;
	for (let j = 0; j < i; j++) {
		const token = statement[j];
		if (token.kind === 'punctuation' && token.rawText === '(') {
			return false;
		}
		if (word(token) === 'declare' && token.kind === 'keyword') {
			declare = true;
		}
	}
	return declare;
}

/** After the For and To of a counted For header, outside any parentheses. */
function inForHeader(statement: readonly VbaToken[], i: number): boolean {
	let depth = 0;
	let sawTo = false;
	for (let j = i - 1; j >= 0; j--) {
		const token = statement[j];
		if (token.kind === 'punctuation' && token.rawText === ')') {
			depth++;
		} else if (token.kind === 'punctuation' && token.rawText === '(') {
			if (depth === 0) {
				return false;
			}
			depth--;
		} else if (depth === 0 && token.kind === 'keyword' && word(token) === 'to') {
			sawTo = true;
		} else if (depth === 0 && token.kind === 'keyword' && word(token) === 'for') {
			return sawTo && word(statement[j - 1]) !== 'exit';
		}
	}
	return false;
}

/** The token can end an operand, so what follows it is not the next part of one. */
function endsOperand(token: VbaToken | undefined): boolean {
	if (!token) {
		return false;
	}
	switch (token.kind) {
		case 'identifier':
		case 'bracketedIdentifier':
		case 'integerLiteral':
		case 'floatLiteral':
		case 'stringLiteral':
		case 'dateLiteral':
			return true;
		case 'punctuation':
			return token.rawText === ')';
		case 'keyword':
			return ['true', 'false', 'nothing', 'empty', 'null', 'me'].includes(word(token));
		default:
			return false;
	}
}
