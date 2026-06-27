// Error-tolerant VBA value-expression parser (MS-VBAL §5.6).
//
// Turns a slice of significant tokens into the ExprNode hierarchy defined in
// nodes.ts. This is the critical-path keystone for roadmap 2.4.0: procedure
// bodies currently keep non-declaration statements as raw text, so flow-sensitive
// binding, branch modeling, and arbitrary-expression typing all wait on a real
// expression AST.
//
// Verified against MS-VBAL.pdf, v20250520:
//   - 5.6     Expressions (value expressions)
//   - 5.6.6   Operator expressions / precedence and associativity
//   - 5.6.9   Index / member-access / function-call expressions
//   - 5.6.10  New / AddressOf / TypeOf...Is
//
// Design notes:
//   * Never throws. On an unexpected token the parser stops at the last good
//     position, records a diagnostic, and returns a best-effort node so callers
//     (and recovery) keep working - the same Phase 3 contract parseModule uses.
//   * Every node carries an absolute source span built from token offsets.
//   * Scope: literals, identifiers, parenthesised expressions, member-access
//     chains (including the leading-dot With form), index/call expressions with
//     positional, named (`name:=expr`), and omitted (`f(1, , 3)`) arguments,
//     unary -/+/Not, and the full binary operator precedence ladder, plus
//     New / AddressOf / TypeOf...Is, and bang member access (`obj!name`).
//   * The §5.6 expression grammar is now fully modeled; statements no longer fall
//     back to raw for named/omitted arguments or bang access.

import { VbaToken } from '../lexer/tokenKinds';
import { isIdentLike, tokenName, tokenWord } from '../lexer/tokenHelpers';
import {
	AddressOfExpr,
	Argument,
	BinaryExpr,
	BinaryOperator,
	ExprNode,
	IdentifierExpr,
	IndexExpr,
	LiteralExpr,
	LiteralKind,
	MemberAccessExpr,
	NewExpr,
	ParenExpr,
	ParseDiagnostic,
	Span,
	TypeOfIsExpr,
	UnaryExpr,
	UnaryOperator,
} from './nodes';

/** Result of parsing an expression from a token slice. */
export interface ExprParseResult {
	/** The parsed expression, or null when the slice held no expression token. */
	expr: ExprNode | null;
	/** Diagnostics raised while parsing (unexpected/trailing tokens, etc.). */
	diagnostics: ParseDiagnostic[];
	/**
	 * Index into the input token array just past the last consumed token. A
	 * caller wiring statements can compare this to the slice end to detect
	 * trailing tokens it does not yet model and fall back accordingly.
	 */
	endIndex: number;
}

/**
 * Binary-operator binding power (MS-VBAL 5.6.6). Higher binds tighter. All
 * binary operators here are left-associative. Exponentiation (`^`) and the
 * unary operators are handled structurally in parseUnary/parsePower because `^`
 * binds tighter than unary minus while still allowing a signed exponent.
 */
const BINARY_PRECEDENCE: Readonly<Record<string, number>> = {
	Imp: 1,
	Eqv: 2,
	Xor: 3,
	Or: 4,
	And: 5,
	// comparison operators share one level
	'=': 6,
	'<>': 6,
	'<': 6,
	'>': 6,
	'<=': 6,
	'>=': 6,
	Like: 6,
	Is: 6,
	'&': 7,
	'+': 8,
	'-': 8,
	Mod: 9,
	'\\': 10,
	'*': 11,
	'/': 11,
};

/** Word-operators that arrive as keyword tokens, mapped to canonical text. */
const WORD_BINARY_OPS: Readonly<Record<string, BinaryOperator>> = {
	mod: 'Mod',
	and: 'And',
	or: 'Or',
	xor: 'Xor',
	eqv: 'Eqv',
	imp: 'Imp',
	like: 'Like',
	is: 'Is',
};

const LITERAL_KEYWORDS: Readonly<Record<string, LiteralKind>> = {
	true: 'boolean',
	false: 'boolean',
	nothing: 'nothing',
	null: 'null',
	empty: 'empty',
};

/** Type-declaration characters that are never binary operators (safe as suffix). */
const UNAMBIGUOUS_SUFFIXES = new Set(['$', '%', '@']);

/**
 * Binding power of the `Not` prefix (MS-VBAL 5.6.6): below the comparison
 * operators but above `And`, so `Not a = b` is `Not (a = b)` while `Not a And b`
 * is `(Not a) And b`.
 */
const NOT_PRECEDENCE = 6;

/**
 * Parse a value expression from tokens[from, to). The token slice must already
 * exclude comments and statement separators (use codeTokens upstream).
 */
export function parseExpression(
	tokens: readonly VbaToken[],
	from = 0,
	to = tokens.length,
): ExprParseResult {
	const parser = new ExpressionParser(tokens, from, to);
	const expr = parser.parse();
	return { expr, diagnostics: parser.diagnostics, endIndex: parser.index };
}

/**
 * Parse a parenless call-statement argument list from tokens[from, to). Returns
 * null when any present argument is malformed (the caller falls back to a raw
 * statement). Supports positional, named (`name:=value`), and omitted arguments.
 */
export function parseParenlessArguments(
	tokens: readonly VbaToken[],
	from: number,
	to: number,
): Argument[] | null {
	return new ExpressionParser(tokens, from, to).parseParenlessArgumentList();
}

class ExpressionParser {
	readonly diagnostics: ParseDiagnostic[] = [];
	index: number;

	constructor(
		private readonly tokens: readonly VbaToken[],
		from: number,
		private readonly to: number,
	) {
		this.index = from;
	}

	parse(): ExprNode | null {
		if (this.atEnd()) {
			return null;
		}
		const expr = this.parseBinary(0);
		return expr;
	}

	// --- token cursor -------------------------------------------------------

	private atEnd(): boolean {
		return this.index >= this.to;
	}

	private peek(): VbaToken | undefined {
		return this.index < this.to ? this.tokens[this.index] : undefined;
	}

	private next(): VbaToken | undefined {
		return this.index < this.to ? this.tokens[this.index++] : undefined;
	}

	private raw(token: VbaToken | undefined): string {
		return token?.rawText ?? '';
	}

	private spanFrom(startToken: VbaToken, endToken: VbaToken): Span {
		return { start: startToken.start, end: endToken.end };
	}

	private spanOf(expr: ExprNode): Span {
		return expr.span;
	}

	private diag(token: VbaToken | undefined, message: string): void {
		const span: Span = token
			? { start: token.start, end: token.end }
			: { start: 0, end: 0 };
		this.diagnostics.push({ span, message, severity: 'error', specRef: 'MS-VBAL 5.6' });
	}

	// --- precedence-climbing binary layer -----------------------------------

	private parseBinary(minPrec: number): ExprNode | null {
		let left: ExprNode | null;
		// `Not` is a low-precedence prefix: recognise it only where a Not-expression
		// is allowed (at or below its binding power), so comparisons inside it bind
		// first while `And`/`Or` outside it bind after.
		const lead = this.peek();
		if (
			lead &&
			lead.kind === 'keyword' &&
			tokenWord(lead) === 'not' &&
			minPrec <= NOT_PRECEDENCE
		) {
			this.next();
			const operand = this.parseBinary(NOT_PRECEDENCE);
			if (!operand) {
				this.diag(lead, "Expected an expression after 'Not'.");
				return null;
			}
			left = {
				exprKind: 'UnaryExpr',
				operator: 'Not',
				operand,
				span: { start: lead.start, end: this.spanOf(operand).end },
			} satisfies UnaryExpr;
		} else {
			left = this.parseUnary();
		}
		if (!left) {
			return null;
		}
		for (;;) {
			const opToken = this.peek();
			const op = this.binaryOperator(opToken);
			if (!op) {
				break;
			}
			const prec = BINARY_PRECEDENCE[op];
			if (prec === undefined || prec < minPrec) {
				break;
			}
			this.next(); // consume operator
			// Left-associative: the right side binds operators strictly tighter.
			const right = this.parseBinary(prec + 1);
			if (!right) {
				this.diag(opToken, `Expected an expression after '${op}'.`);
				return left;
			}
			const node: BinaryExpr = {
				exprKind: 'BinaryExpr',
				operator: op,
				left,
				right,
				span: { start: this.spanOf(left).start, end: this.spanOf(right).end },
			};
			left = node;
		}
		return left;
	}

	/** Canonical binary operator for a token, or null when it is not one. */
	private binaryOperator(token: VbaToken | undefined): BinaryOperator | null {
		if (!token) {
			return null;
		}
		if (token.kind === 'keyword') {
			const word = tokenWord(token);
			return WORD_BINARY_OPS[word] ?? null;
		}
		if (token.kind === 'operator') {
			const raw = token.rawText;
			if (raw in BINARY_PRECEDENCE) {
				return raw as BinaryOperator;
			}
		}
		return null;
	}

	// --- unary / exponent ---------------------------------------------------

	private parseUnary(): ExprNode | null {
		const token = this.peek();
		if (!token) {
			this.diag(undefined, 'Expected an expression.');
			return null;
		}
		const unaryOp = this.prefixOperator(token);
		if (unaryOp) {
			this.next();
			const operand = this.parseUnary();
			if (!operand) {
				this.diag(token, `Expected an expression after '${this.raw(token)}'.`);
				return null;
			}
			const node: UnaryExpr = {
				exprKind: 'UnaryExpr',
				operator: unaryOp,
				operand,
				span: { start: token.start, end: this.spanOf(operand).end },
			};
			return node;
		}
		return this.parsePower();
	}

	/** Prefix `-`, `+`, or `Not` (MS-VBAL 5.6.6). */
	private prefixOperator(token: VbaToken): UnaryOperator | null {
		if (token.kind === 'operator' && (token.rawText === '-' || token.rawText === '+')) {
			return token.rawText as UnaryOperator;
		}
		// `Not` is handled at its own (low) precedence in parseBinary, not here.
		return null;
	}

	/**
	 * Exponentiation binds tighter than unary minus but its right operand may be
	 * signed (`2 ^ -3`). Left-associative: `2 ^ 3 ^ 2` is `(2 ^ 3) ^ 2`.
	 */
	private parsePower(): ExprNode | null {
		let base = this.parsePostfix();
		if (!base) {
			return null;
		}
		while (this.peek()?.kind === 'operator' && this.peek()!.rawText === '^') {
			const opToken = this.next()!;
			const exponent = this.parseSignedPrimary();
			if (!exponent) {
				this.diag(opToken, "Expected an expression after '^'.");
				return base;
			}
			const node: BinaryExpr = {
				exprKind: 'BinaryExpr',
				operator: '^',
				left: base,
				right: exponent,
				span: { start: this.spanOf(base).start, end: this.spanOf(exponent).end },
			};
			base = node;
		}
		return base;
	}

	/** A postfix primary optionally preceded by sign(s) - the operand of `^`. */
	private parseSignedPrimary(): ExprNode | null {
		const token = this.peek();
		if (token && token.kind === 'operator' && (token.rawText === '-' || token.rawText === '+')) {
			this.next();
			const operand = this.parseSignedPrimary();
			if (!operand) {
				return null;
			}
			const node: UnaryExpr = {
				exprKind: 'UnaryExpr',
				operator: token.rawText as UnaryOperator,
				operand,
				span: { start: token.start, end: this.spanOf(operand).end },
			};
			return node;
		}
		return this.parsePostfix();
	}

	// --- postfix: member access and index/call ------------------------------

	private parsePostfix(): ExprNode | null {
		let expr = this.parsePrimary();
		if (!expr) {
			return null;
		}
		for (;;) {
			const token = this.peek();
			if (!token) {
				break;
			}
			if (token.kind === 'punctuation' && token.rawText === '.') {
				this.next();
				const memberToken = this.peek();
				if (!memberToken || !this.isMemberName(memberToken)) {
					this.diag(memberToken ?? token, "Expected a member name after '.'.");
					break;
				}
				this.next();
				const member = tokenName(memberToken) ?? memberToken.rawText;
				const node: MemberAccessExpr = {
					exprKind: 'MemberAccessExpr',
					object: expr,
					member,
					memberSpan: { start: memberToken.start, end: memberToken.end },
					span: { start: this.spanOf(expr).start, end: memberToken.end },
				};
				expr = node;
				continue;
			}
			if (token.kind === 'punctuation' && token.rawText === '(') {
				const indexed = this.parseIndex(expr);
				if (!indexed) {
					break;
				}
				expr = indexed;
				continue;
			}
			// Bang member access: `receiver!name` / `receiver![Bracketed Name]`.
			// Disambiguated from the `!` Single type-declaration suffix by requiring
			// the `!` to be glued to both the receiver and a plain (non-keyword)
			// identifier or bracketed name - canonical bang syntax. A `!` followed by
			// an operator, a keyword (`a! And b`), whitespace, or end-of-statement is
			// a type-suffix / stray operator and is left to the caller (stays raw).
			const bang = this.bangMemberAccess(expr, token);
			if (bang) {
				expr = bang;
				continue;
			}
			break;
		}
		return expr;
	}

	/** Build a `receiver!name` bang member-access node, or null when `!` is not a bang. */
	private bangMemberAccess(receiver: ExprNode, bangToken: VbaToken): MemberAccessExpr | null {
		if (bangToken.kind !== 'operator' || bangToken.rawText !== '!') {
			return null;
		}
		const nameTok = this.peekAt(1);
		if (
			!nameTok ||
			!(nameTok.kind === 'identifier' || nameTok.kind === 'bracketedIdentifier') ||
			bangToken.start !== this.spanOf(receiver).end || // `!` glued to receiver
			nameTok.start !== bangToken.end // name glued to `!`
		) {
			return null;
		}
		this.next(); // consume !
		this.next(); // consume name
		const member = tokenName(nameTok) ?? nameTok.rawText;
		return {
			exprKind: 'MemberAccessExpr',
			object: receiver,
			member,
			memberSpan: { start: nameTok.start, end: nameTok.end },
			accessKind: 'bang',
			span: { start: this.spanOf(receiver).start, end: nameTok.end },
		};
	}

	/** `callee(args)` - positional, named (`name:=expr`), and omitted (`f(1, , 3)`) arguments. */
	private parseIndex(callee: ExprNode): IndexExpr | null {
		const open = this.next()!; // consume '('
		const args: Argument[] = [];
		// Empty argument list: `callee()`.
		if (this.peek()?.kind === 'punctuation' && this.peek()!.rawText === ')') {
			const close = this.next()!;
			return {
				exprKind: 'IndexExpr',
				callee,
				args,
				span: { start: this.spanOf(callee).start, end: close.end },
			};
		}
		for (;;) {
			const arg = this.parseArgument(')');
			if (!arg) {
				this.diag(this.peek() ?? open, 'Expected an argument expression.');
				return null;
			}
			args.push(arg);
			const sep = this.peek();
			if (sep?.kind === 'punctuation' && sep.rawText === ',') {
				this.next();
				continue;
			}
			if (sep?.kind === 'punctuation' && sep.rawText === ')') {
				const close = this.next()!;
				return {
					exprKind: 'IndexExpr',
					callee,
					args,
					span: { start: this.spanOf(callee).start, end: close.end },
				};
			}
			this.diag(sep ?? open, "Expected ',' or ')' in argument list.");
			return null;
		}
	}

	/**
	 * Parse a parenless (call-statement) argument list from the parser's remaining
	 * tokens. Returns null when any present argument is malformed so the caller can
	 * fall back to a raw statement; an empty trailing slot after a `,` (`Foo a,`) is
	 * treated as malformed rather than an omission, preserving the conservative
	 * no-regression boundary.
	 */
	parseParenlessArgumentList(): Argument[] | null {
		const args: Argument[] = [];
		for (;;) {
			const arg = this.parseArgument(undefined);
			if (!arg) {
				return null;
			}
			args.push(arg);
			const sep = this.peek();
			if (!sep) {
				return args; // reached the end of the slice
			}
			if (sep.kind === 'punctuation' && sep.rawText === ',') {
				this.next();
				continue;
			}
			return null; // unexpected trailing token - malformed
		}
	}

	/**
	 * Parse a single argument: an optional `name:=` prefix then a value, or an
	 * omitted slot (an empty position before a `,` or the terminator). Returns null
	 * only when a value is expected but fails to parse, so the caller falls back to
	 * raw. `terminator` is `)` for a parenthesised list, or undefined for a parenless
	 * list (which ends at the slice boundary, never an omission).
	 */
	private parseArgument(terminator: ')' | undefined): Argument | null {
		const head = this.peek();
		// Omitted slot: the position is empty (the next token is a separator or the
		// list terminator). Modeled as a zero-width span at that position.
		if (this.atArgumentBoundary(head, terminator)) {
			const pos = head!.start;
			return { value: null, span: { start: pos, end: pos } };
		}
		if (!head) {
			return null;
		}
		// Named argument: `name := value`.
		let name: string | undefined;
		let nameSpan: Span | undefined;
		const colonEq = this.peekAt(1);
		if (
			this.isMemberName(head) &&
			colonEq &&
			colonEq.kind === 'operator' &&
			colonEq.rawText === ':='
		) {
			name = tokenName(head) ?? head.rawText;
			nameSpan = { start: head.start, end: head.end };
			this.next(); // consume name
			this.next(); // consume :=
		}
		const value = this.parseBinary(0);
		if (!value) {
			return null;
		}
		const start = nameSpan ? nameSpan.start : this.spanOf(value).start;
		return { name, nameSpan, value, span: { start, end: this.spanOf(value).end } };
	}

	/** True when the current position is an empty argument slot (a `,` or terminator). */
	private atArgumentBoundary(token: VbaToken | undefined, terminator: ')' | undefined): boolean {
		if (!token || token.kind !== 'punctuation') {
			return false;
		}
		return token.rawText === ',' || (terminator !== undefined && token.rawText === terminator);
	}

	private peekAt(offset: number): VbaToken | undefined {
		const i = this.index + offset;
		return i < this.to ? this.tokens[i] : undefined;
	}

	private isMemberName(token: VbaToken): boolean {
		return isIdentLike(token) || token.kind === 'bracketedIdentifier';
	}

	// --- primary ------------------------------------------------------------

	private parsePrimary(): ExprNode | null {
		const token = this.peek();
		if (!token) {
			this.diag(undefined, 'Expected an expression.');
			return null;
		}

		// Parenthesised expression.
		if (token.kind === 'punctuation' && token.rawText === '(') {
			this.next();
			const inner = this.parseBinary(0);
			if (!inner) {
				return null;
			}
			const close = this.peek();
			if (close?.kind === 'punctuation' && close.rawText === ')') {
				this.next();
				const node: ParenExpr = {
					exprKind: 'ParenExpr',
					inner,
					span: { start: token.start, end: close.end },
				};
				return node;
			}
			this.diag(close ?? token, "Expected ')'.");
			return inner;
		}

		// Leading-dot member access inside a With block (`.Member`).
		if (token.kind === 'punctuation' && token.rawText === '.') {
			this.next();
			const memberToken = this.peek();
			if (!memberToken || !this.isMemberName(memberToken)) {
				this.diag(memberToken ?? token, "Expected a member name after '.'.");
				return null;
			}
			this.next();
			const member = tokenName(memberToken) ?? memberToken.rawText;
			const node: MemberAccessExpr = {
				exprKind: 'MemberAccessExpr',
				object: null,
				member,
				memberSpan: { start: memberToken.start, end: memberToken.end },
				span: { start: token.start, end: memberToken.end },
			};
			return node;
		}

		// Literals.
		const literal = this.literalFor(token);
		if (literal) {
			this.next();
			return literal;
		}

		// Keyword-led primaries: New / AddressOf / TypeOf.
		if (token.kind === 'keyword') {
			const word = tokenWord(token);
			if (word === 'new') {
				return this.parseNew(token);
			}
			if (word === 'addressof') {
				return this.parseAddressOf(token);
			}
			if (word === 'typeof') {
				return this.parseTypeOf(token);
			}
		}

		// Identifier reference.
		if (token.kind === 'identifier' || token.kind === 'bracketedIdentifier') {
			this.next();
			return this.identifierExpr(token);
		}

		this.diag(token, `Unexpected token '${token.rawText}' in expression.`);
		return null;
	}

	private identifierExpr(token: VbaToken): IdentifierExpr {
		const name = tokenName(token) ?? token.rawText;
		const node: IdentifierExpr = {
			exprKind: 'IdentifierExpr',
			name,
			span: { start: token.start, end: token.end },
		};
		// Capture an immediately-adjacent unambiguous type-declaration character
		// ($/%/@). &/!/#/^ are left to the operator layer to avoid misreading a
		// concatenation, bang, date, or exponent.
		const suffix = this.peek();
		if (
			suffix &&
			suffix.kind === 'unknown' &&
			UNAMBIGUOUS_SUFFIXES.has(suffix.rawText) &&
			suffix.start === token.end
		) {
			this.next();
			node.typeSuffix = suffix.rawText;
			node.span = { start: token.start, end: suffix.end };
		}
		return node;
	}

	private literalFor(token: VbaToken): LiteralExpr | null {
		let literalKind: LiteralKind | null = null;
		switch (token.kind) {
			case 'integerLiteral':
				literalKind = 'integer';
				break;
			case 'floatLiteral':
				literalKind = 'float';
				break;
			case 'stringLiteral':
				literalKind = 'string';
				break;
			case 'dateLiteral':
				literalKind = 'date';
				break;
			case 'keyword': {
				literalKind = LITERAL_KEYWORDS[tokenWord(token)] ?? null;
				break;
			}
			default:
				literalKind = null;
		}
		if (!literalKind) {
			return null;
		}
		return {
			exprKind: 'LiteralExpr',
			literalKind,
			raw: token.rawText,
			span: { start: token.start, end: token.end },
		};
	}

	private parseNew(keyword: VbaToken): NewExpr | null {
		this.next(); // consume New
		const typeToken = this.peek();
		if (!typeToken || !this.isMemberName(typeToken)) {
			this.diag(typeToken ?? keyword, "Expected a type name after 'New'.");
			return null;
		}
		// A New type name may be a dotted library type (e.g. Scripting.Dictionary).
		let endToken = this.next()!;
		let typeName = tokenName(endToken) ?? endToken.rawText;
		while (this.peek()?.kind === 'punctuation' && this.peek()!.rawText === '.') {
			this.next();
			const part = this.peek();
			if (!part || !this.isMemberName(part)) {
				break;
			}
			this.next();
			typeName += `.${tokenName(part) ?? part.rawText}`;
			endToken = part;
		}
		return {
			exprKind: 'NewExpr',
			typeName,
			typeNameSpan: { start: keyword.start, end: endToken.end },
			span: { start: keyword.start, end: endToken.end },
		};
	}

	private parseAddressOf(keyword: VbaToken): AddressOfExpr | null {
		this.next(); // consume AddressOf
		const targetToken = this.peek();
		if (!targetToken || !(targetToken.kind === 'identifier' || isIdentLike(targetToken))) {
			this.diag(targetToken ?? keyword, "Expected a procedure name after 'AddressOf'.");
			return null;
		}
		this.next();
		const target = this.identifierExpr(targetToken);
		return {
			exprKind: 'AddressOfExpr',
			target,
			span: { start: keyword.start, end: target.span.end },
		};
	}

	private parseTypeOf(keyword: VbaToken): TypeOfIsExpr | null {
		this.next(); // consume TypeOf
		// Operand runs up to the `Is` keyword; parse it as a postfix expression
		// (a receiver chain), which is what `TypeOf <expr> Is <Type>` allows.
		const operand = this.parsePostfix();
		if (!operand) {
			this.diag(keyword, "Expected an expression after 'TypeOf'.");
			return null;
		}
		const isToken = this.peek();
		if (!isToken || isToken.kind !== 'keyword' || tokenWord(isToken) !== 'is') {
			this.diag(isToken ?? keyword, "Expected 'Is' in a 'TypeOf ... Is' expression.");
			return null;
		}
		this.next(); // consume Is
		const typeToken = this.peek();
		if (!typeToken || !this.isMemberName(typeToken)) {
			this.diag(typeToken ?? isToken, "Expected a type name after 'Is'.");
			return null;
		}
		this.next();
		const typeName = tokenName(typeToken) ?? typeToken.rawText;
		return {
			exprKind: 'TypeOfIsExpr',
			operand,
			typeName,
			typeNameSpan: { start: typeToken.start, end: typeToken.end },
			span: { start: keyword.start, end: typeToken.end },
		};
	}
}
