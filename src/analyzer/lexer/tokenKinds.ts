// VBA lexical token kinds.
//
// Verified against MS-VBAL.pdf, v20250520 (Release: May 20, 2025),
// section 3.3 "Lexical Tokens" and its subsections 3.3.1 - 3.3.5.
//
// The lexer is loss-aware: every character of the source appears in exactly one
// token's rawText or in a token's leadingTrivia, so concatenating the stream
// reproduces the original source exactly (round-trippable; Phase 1 acceptance
// criterion).

/** The category of a lexical token. */
export type TokenKind =
	// End of a logical line - the line-terminator (MS-VBAL 3.2.2 / 3.3.1 EOL).
	| 'newline'
	// Apostrophe comment or Rem comment (MS-VBAL 3.3.1 comment-body / 3.3.5.2
	// rem-keyword). Non-syntactic but preserved.
	| 'comment'
	// A reserved identifier or contextual keyword (MS-VBAL 3.3.5.2). canonicalText
	// holds the canonical capitalization.
	| 'keyword'
	// An <IDENTIFIER>: a lex-identifier that is not a reserved-identifier
	// (MS-VBAL 3.3.5 / 3.3.5.2).
	| 'identifier'
	// A FOREIGN-NAME: "[" foreign-identifier "]" (MS-VBAL 3.3.5.3).
	| 'bracketedIdentifier'
	// INTEGER token (MS-VBAL 3.3.2), including any %/&/^ type suffix.
	| 'integerLiteral'
	// FLOAT token (MS-VBAL 3.3.2), including any !/#/@ type suffix.
	| 'floatLiteral'
	// DATE token delimited by '#' (MS-VBAL 3.3.3 date-or-time / 5.6.5).
	| 'dateLiteral'
	// STRING token (MS-VBAL 3.3.4).
	| 'stringLiteral'
	// An arithmetic, comparison, logical, or concatenation symbol drawn from
	// special-token (MS-VBAL 3.3.1), e.g. = < > <= >= <> + - * / \ ^ & := !
	| 'operator'
	// Structural punctuation from special-token: , . ( ) ;
	| 'punctuation'
	// The ':' statement separator (MS-VBAL 3.3.1 EOS). ':=' is an 'operator'.
	| 'colon'
	// A '#' that begins a conditional-compilation directive line
	// (MS-VBAL 3.4: #Const / #If / #ElseIf / #Else / #EndIf).
	| 'directive'
	// Any character sequence that does not match a known lexical rule.
	| 'unknown';

/** Trivia kinds: insignificant text attached to the following token. */
export type TriviaKind =
	// One or more WSC characters (MS-VBAL 3.2.2 WSC): tab, space, etc.
	| 'whitespace'
	// A line-continuation: 1*WSC underscore line-terminator (MS-VBAL 3.2.2).
	| 'lineContinuation';

/** Insignificant text (whitespace or line continuation) attached to a token. */
export interface Trivia {
	kind: TriviaKind;
	text: string;
	/** Absolute UTF-16 offset of the first character. */
	start: number;
	/** Absolute UTF-16 offset just past the last character. */
	end: number;
}

/** A VBA lexical token. */
export interface VbaToken {
	kind: TokenKind;
	/** The exact source text of the token. */
	rawText: string;
	/**
	 * Canonical capitalization for keyword tokens (MS-VBAL 3.3.5.2). Undefined
	 * for all non-keyword tokens. Never set for text inside comments or strings.
	 */
	canonicalText?: string;
	/** Absolute UTF-16 offset of the first character of rawText. */
	start: number;
	/** Absolute UTF-16 offset just past the last character of rawText. */
	end: number;
	/** Zero-based line of the token start (VS Code Position convention). */
	line: number;
	/** Zero-based UTF-16 column of the token start. */
	character: number;
	/** Whitespace / line continuations immediately preceding this token. */
	leadingTrivia?: Trivia[];
	/**
	 * Whitespace / line continuations following this token. Only populated for
	 * trailing trivia at end-of-file (which has no following token to lead), so
	 * that the token stream remains perfectly round-trippable.
	 */
	trailingTrivia?: Trivia[];
}

/** WSC - whitespace characters (MS-VBAL 3.2.2). Excludes line terminators. */
export function isWsc(ch: string): boolean {
	switch (ch) {
		case '\t': // %x0009 tab-character
		case '\u0019': // %x0019 eom-character
		case ' ': // %x0020 space-character
		case '\u3000': // %x3000 DBCS-whitespace
			return true;
		default:
			// most-Unicode-class-Zs: other Unicode space separators.
			return /\s/.test(ch) && ch !== '\n' && ch !== '\r' && ch !== '\v' && ch !== '\f';
	}
}

/** True if ch is a line terminator (CR or LF). */
export function isLineTerminator(ch: string): boolean {
	return ch === '\r' || ch === '\n';
}
