import type { VbaToken } from '../lexer/tokenKinds';
import { tokensWithoutLeadingLineNumber } from '../lexer/tokenHelpers';
import { parseExpression } from '../parser/parseExpression';

/** Recognize an assignment slot using only the current statement tokens. */
export function assignmentTargetFromTokens(statement: readonly VbaToken[]): VbaToken[] | undefined {
	let tokens = tokensWithoutLeadingLineNumber(statement);
	if (tokens.at(-1)?.rawText !== '=') { return undefined; }
	// Only the consequent of a single-line If is an assignment; its condition is not.
	let branch = -1;
	for (let i = 0; i < tokens.length; i++) {
		if (tokens[0]?.rawText.toLowerCase() === 'if' && tokens[i].kind === 'keyword' && /^(Then|Else)$/i.test(tokens[i].rawText)) { branch = i; }
	}
	if (branch >= 0) { tokens = tokens.slice(branch + 1); }
	tokens = tokens.slice(0, -1);
	if (tokens[0]?.rawText.toLowerCase() === 'let') { tokens.shift(); }
	const first = tokens[0];
	if (!first || (first.kind !== 'identifier' && first.kind !== 'bracketedIdentifier'
		&& first.rawText !== '.' && !/^(Me|ThisWorkbook)$/i.test(first.rawText))) { return undefined; }
	// The expression parser's statement callers normally bind Me as a receiver.
	// This isolated syntactic check needs only its identifier-shaped grammar.
	const syntaxTokens = tokens.map(t => t.kind === 'keyword' && t.rawText.toLowerCase() === 'me'
		? { ...t, kind: 'identifier' as const } : t);
	const parsed = parseExpression(syntaxTokens);
	return parsed.expr && parsed.endIndex === tokens.length && !parsed.diagnostics.length
		&& ['IdentifierExpr', 'MemberAccessExpr', 'IndexExpr'].includes(parsed.expr.exprKind) ? tokens : undefined;
}
