import { readFileSync } from 'fs';
import { describe, expect, it } from 'vitest';

interface GrammarRule {
	comment?: string;
	match?: string;
	patterns?: GrammarRule[];
}

interface Grammar {
	repository: Record<string, GrammarRule>;
}

function loadGrammar(): Grammar {
	return JSON.parse(readFileSync('syntaxes/vba.tmLanguage.json', 'utf8')) as Grammar;
}

describe('VBA grammar on long lines', () => {
	it('searches long lines of [ or blanks quickly', () => {
		// The foreign-name rule let a [ stand inside the name, so on a line of [ with
		// no ] the search from every [ ran on to the end of the line: about 0.3 s for
		// 16,000 of them in Node. The line-continuation rule did the same over a run of
		// blanks with no _. JavaScript's regex engine is the one these patterns meet
		// outside VS Code, and both are valid JavaScript patterns. Each line below takes
		// a few milliseconds when the time grows with the line's length, far under the
		// bound even on a busy machine.
		const grammar = loadGrammar();
		const foreign = grammar.repository.identifiers.patterns?.find((rule) => rule.comment?.startsWith('Foreign name'))?.match;
		const continuation = grammar.repository['line-continuation'].match;
		expect(foreign).toBeDefined();
		expect(continuation).toBeDefined();
		const patterns = [new RegExp(foreign ?? '', 'u'), new RegExp(continuation ?? '', 'u')];
		const n = 20000;
		for (const prefix of ['', 'x = ', "x = ' "]) {
			for (const fragment of ['[', '[a', '[ ', ' ', '\t', ' [']) {
				for (const ending of ['', 'x', ']']) {
					const line = `${prefix}${fragment.repeat(n)}${ending}`;
					const start = performance.now();
					for (const pattern of patterns) {
						pattern.test(line);
					}
					expect(performance.now() - start, JSON.stringify(`${prefix}${fragment}...${ending}`)).toBeLessThan(250);
				}
			}
		}
	}, 30_000);
});
