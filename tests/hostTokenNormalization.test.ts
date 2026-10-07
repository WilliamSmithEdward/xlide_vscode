import {expect, it} from 'vitest';
import {EMPTY_HOST_MODEL, hostObjectModelForTokens} from '../src/analyzer/host/hostRegistry';
import {withResolvedHostModel} from '../src/analyzer/diagnostics/analyzeModule';
import {analyzeModule} from '../src/analyzer';

const cases = [['word', 'excel'], ['excel', 'word'], ['word', 'excel', 'powerpoint']].flatMap(tokens => ['upper', 'spaces', 'mixed'].map(mode => ({tokens, mode})));
it.each(cases)('normalizes $tokens using $mode before selecting and caching merged providers', ({tokens, mode}) => {
	const expected = hostObjectModelForTokens(tokens)!;
	const input = Object.freeze(tokens.map((token, i) => mode === 'upper' ? token.toUpperCase() : mode === 'spaces' ? '  ' + token + '\t' : i % 2 ? token.toUpperCase() : ' ' + token + ' '));
	const before = [...input], actual = hostObjectModelForTokens(input);
	expect(actual?.source).toBe(expected.source);
	expect(actual?.hostName).toBe(tokens[0] === 'word' ? 'Word' : 'Excel');
	expect(actual?.aliases.range).toBe(tokens[0] === 'word' ? 'Word.Range' : 'Excel.Range');
	expect(actual === expected).toBe(true);
	expect(input).toEqual(before);
});

it('resolves normalized hosts and references through analyzer options', () => {
	const canonical = {host: 'word', referencedHosts: ['excel', 'powerpoint']}, supplied = {host: ' Word ', referencedHosts: Object.freeze([' EXCEL ', 'PowerPoint'])};
	expect(withResolvedHostModel(supplied).hostModel === withResolvedHostModel(canonical).hostModel).toBe(true);
	const source = 'Option Explicit\nSub S()\nDim xl As Excel.Application\nDim slide As PowerPoint.Slide\nxl.Calculate\nslide.Copy\nEnd Sub\n';
	const failures: unknown[] = [];
	expect(analyzeModule(source, {...supplied, onInternalError: (error, where) => {failures.push([String(error), where]);}})).toEqual(analyzeModule(source, canonical));
	expect(failures).toEqual([]);
});

it('keeps unknown, duplicate and single-token default behavior', () => {
	expect(hostObjectModelForTokens([])).toBeUndefined();
	expect(hostObjectModelForTokens([' EXCEL '])).toBeUndefined();
	expect(hostObjectModelForTokens([' UNKNOWN '])).toBe(EMPTY_HOST_MODEL);
	expect(hostObjectModelForTokens(['word', 'word'])?.source).toBe(hostObjectModelForTokens(['word'])!.source + ' + ' + hostObjectModelForTokens(['word'])!.source);
	expect(hostObjectModelForTokens(['unknown', 'word', 'excel'])).toBe(hostObjectModelForTokens(['word', 'excel']));
});
