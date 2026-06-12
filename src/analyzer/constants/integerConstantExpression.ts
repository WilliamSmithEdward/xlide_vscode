// Shared VBA integer-constant expression evaluation.
//
// One evaluator for declared-constant/enum-member integer expressions, used by
// both the project-wide symbol graph (exported constant surfaces) and the
// diagnostics engine (fixed-length strings, runtime argument bounds, division
// by zero). Keeping a single copy guarantees the project-visible constant
// values and the diagnostics rules can never disagree on the same expression.
//
// The grammar is deliberately conservative: +, -, * (binary and unary +/-),
// parentheses, integer literals (decimal, &H hex, &O octal, with an optional
// %/&/^ type suffix), bare constant names, and `Module.Constant` qualified
// names. Anything else evaluates to undefined so callers never guess.

import { tokenize } from '../lexer/tokenize';
import type { VbaToken } from '../lexer/tokenKinds';
import { tokenName } from '../lexer/tokenHelpers';

/** Lookup of integer constant values by lowercased (possibly qualified) name. */
export interface IntegerConstantLookup {
	get(name: string): number | undefined;
}

/** Parses an unsigned decimal integer literal, rejecting unsafe magnitudes. */
export function parseDecimalIntegerLiteral(raw: string): number | undefined {
	if (!/^\d+$/.test(raw)) {
		return undefined;
	}
	const value = Number(raw);
	return Number.isSafeInteger(value) ? value : undefined;
}

/** Parses a VBA integer literal (decimal, &H, &O; optional %/&/^ suffix). */
export function parseVbaIntegerLiteral(raw: string): number | undefined {
	const text = raw.trim().replace(/[%&^]$/, '');
	const hex = /^&[hH]([0-9A-Fa-f]+)$/.exec(text);
	if (hex) {
		const value = Number.parseInt(hex[1], 16);
		return Number.isSafeInteger(value) ? value : undefined;
	}
	const octal = /^&[oO]([0-7]+)$/.exec(text);
	if (octal) {
		const value = Number.parseInt(octal[1], 8);
		return Number.isSafeInteger(value) ? value : undefined;
	}
	return parseDecimalIntegerLiteral(text);
}

/** Clamps arithmetic results to safe integers; undefined when out of range. */
export function safeInteger(value: number): number | undefined {
	return Number.isSafeInteger(value) ? value : undefined;
}

/**
 * Raw value expression of an enum member: the explicit initializer when
 * present, otherwise the implicit MS-VBAL rule of previous member + 1 (first
 * member defaults to 0).
 */
export function enumMemberRawExpression(
	explicitRaw: string | undefined,
	previousName: string | undefined,
): string {
	return explicitRaw ?? (previousName ? `${previousName} + 1` : '0');
}

/** Evaluates one raw constant expression against already-known constants. */
export function evaluateIntegerConstantExpression(
	raw: string,
	constants: IntegerConstantLookup,
): number | undefined {
	return new IntegerConstantExpressionParser(raw, constants).parse();
}

class IntegerConstantExpressionParser {
	private readonly tokens: VbaToken[];
	private index = 0;

	constructor(
		raw: string,
		private readonly constants: IntegerConstantLookup,
	) {
		this.tokens = tokenize(raw).filter((token) => token.kind !== 'comment' && token.kind !== 'newline');
	}

	parse(): number | undefined {
		if (this.tokens.length === 0) {
			return undefined;
		}
		const value = this.expression();
		return value !== undefined && !this.current() ? value : undefined;
	}

	private expression(): number | undefined {
		let value = this.term();
		while (value !== undefined) {
			if (this.accept('+')) {
				const right = this.term();
				value = right === undefined ? undefined : safeInteger(value + right);
				continue;
			}
			if (this.accept('-')) {
				const right = this.term();
				value = right === undefined ? undefined : safeInteger(value - right);
				continue;
			}
			break;
		}
		return value;
	}

	private term(): number | undefined {
		let value = this.factor();
		while (value !== undefined) {
			if (!this.accept('*')) {
				break;
			}
			const right = this.factor();
			value = right === undefined ? undefined : safeInteger(value * right);
		}
		return value;
	}

	private factor(): number | undefined {
		if (this.accept('+')) {
			return this.factor();
		}
		if (this.accept('-')) {
			const value = this.factor();
			return value === undefined ? undefined : safeInteger(-value);
		}
		if (this.accept('(')) {
			const value = this.expression();
			return value !== undefined && this.accept(')') ? value : undefined;
		}
		const token = this.current();
		if (!token) {
			return undefined;
		}
		if (token.kind === 'integerLiteral') {
			this.index++;
			return parseVbaIntegerLiteral(token.rawText);
		}
		const qualified = this.qualifiedName();
		if (qualified) {
			return this.constants.get(qualified.toLowerCase());
		}
		const name = tokenName(token);
		if (name) {
			this.index++;
			return this.constants.get(name.toLowerCase());
		}
		return undefined;
	}

	private qualifiedName(): string | undefined {
		const qualifier = tokenName(this.current());
		const dot = this.tokens[this.index + 1];
		const member = tokenName(this.tokens[this.index + 2]);
		if (!qualifier || dot?.rawText !== '.' || !member) {
			return undefined;
		}
		this.index += 3;
		return `${qualifier}.${member}`;
	}

	private current(): VbaToken | undefined {
		return this.tokens[this.index];
	}

	private accept(raw: string): boolean {
		if (this.current()?.rawText !== raw) {
			return false;
		}
		this.index++;
		return true;
	}
}

/**
 * Resolves raw constant expressions (lowercased name -> raw text, undefined
 * for ambiguous duplicates) to integer values, memoized with cycle detection.
 * Names absent from `rawConstants` fall back to the optional `base` map of
 * already-resolved values; the returned map only contains `rawConstants` keys.
 */
export function resolveRawIntegerConstants(
	rawConstants: ReadonlyMap<string, string | undefined>,
	base: ReadonlyMap<string, number | undefined> = new Map(),
): Map<string, number | undefined> {
	const resolved = new Map<string, number | undefined>();
	const resolving = new Set<string>();
	const resolve = (name: string): number | undefined => {
		const key = name.toLowerCase();
		if (resolved.has(key)) {
			return resolved.get(key);
		}
		if (!rawConstants.has(key)) {
			return base.get(key);
		}
		if (resolving.has(key)) {
			resolved.set(key, undefined);
			return undefined;
		}
		const raw = rawConstants.get(key);
		if (raw === undefined) {
			resolved.set(key, undefined);
			return undefined;
		}
		resolving.add(key);
		const value = evaluateIntegerConstantExpression(raw, { get: resolve });
		resolving.delete(key);
		resolved.set(key, value);
		return value;
	};
	for (const key of rawConstants.keys()) {
		resolve(key);
	}
	return resolved;
}
