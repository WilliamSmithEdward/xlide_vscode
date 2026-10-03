import { expect, it } from 'vitest';
import { topLevelEqualsIndex } from '../src/analyzer/lexer/tokenHelpers';
import { tokenize } from '../src/analyzer/lexer/tokenize';

it.each([
	['target = Call(x = 1)', 1],
	['target(index = 1) = 2', 6],
	['Call(x = 1)', -1],
	['target = (1 + 2', 1],
	['target) = 2', -1],
	['', -1],
])('retains raw-token assignment scanning for %s', (source, expected) => {
	const tokens = tokenize(source).filter(token => token.kind !== 'whitespace' && token.kind !== 'eof');
	expect(topLevelEqualsIndex(tokens)).toBe(expected);
});
