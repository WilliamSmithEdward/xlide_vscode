import type { VbaToken } from '../lexer/tokenKinds';
import { tokenWord } from '../lexer/tokenHelpers';

export interface FixedLengthStringType {
	typeStart: number;
	starIndex: number;
	lengthIndex: number;
	endIndex: number;
}

export function parseFixedLengthStringType(
	tokens: readonly VbaToken[],
	typeStart: number,
): FixedLengthStringType | undefined {
	if (tokenWord(tokens[typeStart]) !== 'string') {
		return undefined;
	}
	if (tokens[typeStart + 1]?.rawText !== '*') {
		return undefined;
	}
	if (!isFixedLengthStringLengthToken(tokens[typeStart + 2])) {
		return undefined;
	}
	return {
		typeStart,
		starIndex: typeStart + 1,
		lengthIndex: typeStart + 2,
		endIndex: typeStart + 3,
	};
}

function isFixedLengthStringLengthToken(token: VbaToken | undefined): boolean {
	if (!token) {
		return false;
	}
	return (
		token.kind === 'integerLiteral' ||
		token.kind === 'identifier' ||
		token.kind === 'bracketedIdentifier'
	);
}
