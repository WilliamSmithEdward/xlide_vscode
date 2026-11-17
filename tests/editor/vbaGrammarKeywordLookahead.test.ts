import { readFileSync } from 'fs';
import { describe, expect, it } from 'vitest';

interface GrammarRule {
	match?: string;
	patterns?: GrammarRule[];
}

interface Grammar {
	repository: Record<string, GrammarRule>;
}

/** A rule's pattern as a JavaScript RegExp: its leading (?i) becomes the i flag. */
function firstPattern(grammar: Grammar, name: string): RegExp {
	const source = grammar.repository[name].patterns?.[0]?.match ?? '';
	expect(source.startsWith('(?i)'), name).toBe(true);
	return new RegExp(source.slice('(?i)'.length), 'giu');
}

describe('VBA grammar keyword rules on long lines', () => {
	it('searches a long line of True( or CInt( quickly', () => {
		// The rules for True, False, Nothing, Null and Empty and for the conversions,
		// LBound, UBound and Debug looked past a following ( for a ) and As, for an array
		// Type field. On a line with no ), every such word searched to the end of the line:
		// 1.4 to 3.8 s for 100,000 characters in Node. The VBE takes an array field named
		// like one of these words for an error, so the rules no longer look past the (.
		// Each line below takes a few milliseconds when the time grows with the line's
		// length, far under the bound even on a busy machine.
		const grammar = loadGrammar();
		const patterns = [firstPattern(grammar, 'constants'), firstPattern(grammar, 'reserved-names')];
		const n = 100000;
		for (const word of ['True', 'Nothing', 'CInt', 'LBound', 'Debug']) {
			for (const fragment of [`${word}(`, `${word} (`, `${word}(a `]) {
				const line = `x = ${fragment.repeat(Math.ceil(n / fragment.length))}`;
				const start = performance.now();
				for (const pattern of patterns) {
					pattern.lastIndex = 0;
					while (pattern.exec(line)) {
						// Every word is a match; the search is what is timed.
					}
				}
				expect(performance.now() - start, JSON.stringify(`${fragment}...`)).toBeLessThan(250);
			}
		}
	}, 30_000);
});

function loadGrammar(): Grammar {
	return JSON.parse(readFileSync('syntaxes/vba.tmLanguage.json', 'utf8')) as Grammar;
}
