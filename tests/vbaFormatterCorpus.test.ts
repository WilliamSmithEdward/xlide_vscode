// The formatter over every VBA sample the repository holds: the output must
// lex to the input's tokens, formatting twice must equal formatting once, and
// the formatter must never decline on real code.

import { describe, expect, it } from 'vitest';
import { formatVbaModule, tokenStreamDifference } from '../src/analyzer/format/formatModule';
import { allStructuralComparisonSamples } from './helpers/structuralCorpus';

const OPTIONS = { tabSize: 4, insertSpaces: true };

describe('formatVbaModule over the corpus', () => {
	const samples = allStructuralComparisonSamples();

	it('has samples to run over', () => {
		expect(samples.length).toBeGreaterThan(100);
	});

	it('never declines and never changes a token', () => {
		const refusals: string[] = [];
		for (const sample of samples) {
			const result = formatVbaModule(sample.source, OPTIONS);
			if (result.text === undefined) {
				refusals.push(`${sample.id}: ${result.refusal}`);
				continue;
			}
			const difference = tokenStreamDifference(sample.source, result.text);
			if (difference) {
				refusals.push(`${sample.id}: ${difference}`);
			}
		}
		expect(refusals).toEqual([]);
	});

	it('is idempotent', () => {
		const unstable: string[] = [];
		for (const sample of samples) {
			const once = formatVbaModule(sample.source, OPTIONS).text;
			if (once === undefined) {
				continue;
			}
			const twice = formatVbaModule(once, OPTIONS).text;
			if (twice !== once) {
				unstable.push(sample.id);
			}
		}
		expect(unstable).toEqual([]);
	});

	it('keeps every line count', () => {
		for (const sample of samples) {
			const out = formatVbaModule(sample.source, OPTIONS).text ?? sample.source;
			expect(out.split(/\r\n|\r|\n/).length).toBe(sample.source.split(/\r\n|\r|\n/).length);
		}
	});
});
