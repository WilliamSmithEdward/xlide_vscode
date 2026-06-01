import { readFileSync } from 'fs';
import { describe, expect, it } from 'vitest';

interface VbaLanguageConfiguration {
	indentationRules?: Record<string, string>;
	onEnterRules?: Array<{ beforeText?: string; afterText?: string }>;
}

function loadConfig(): VbaLanguageConfiguration {
	return JSON.parse(
		readFileSync('language-configuration/vba-language-configuration.json', 'utf8'),
	) as VbaLanguageConfiguration;
}

describe('VBA language configuration', () => {
	it('uses JavaScript-compatible indentation and enter regexes', () => {
		const config = loadConfig();
		const patterns = [
			...Object.values(config.indentationRules ?? {}),
			...(config.onEnterRules ?? []).flatMap((rule) =>
				[rule.beforeText, rule.afterText].filter((pattern): pattern is string => Boolean(pattern)),
			),
		];

		for (const pattern of patterns) {
			expect(() => new RegExp(pattern)).not.toThrow();
		}
	});

	it('matches lowercase block openers for editor auto-indent', () => {
		const config = loadConfig();
		const rule = new RegExp(config.indentationRules?.increaseIndentPattern ?? '');

		expect(rule.test('    if ready then')).toBe(true);
		expect(rule.test('    for each item in collection')).toBe(true);
		expect(rule.test('    with activesheet')).toBe(true);
	});
});
