import { readFileSync } from 'fs';
import { describe, expect, it } from 'vitest';
import { detectSmartBlockOpener } from '../src/vbaLinter';

interface VbaLanguageConfiguration {
	indentationRules?: Record<string, string>;
	folding?: { markers?: Record<string, string> };
	onEnterRules?: Array<{ beforeText?: string; afterText?: string }>;
}

function loadConfig(): VbaLanguageConfiguration {
	return JSON.parse(
		readFileSync('language-configuration/vba-language-configuration.json', 'utf8'),
	) as VbaLanguageConfiguration;
}

function enterRuleMatches(config: VbaLanguageConfiguration, line: string): boolean {
	return (config.onEnterRules ?? []).some((rule) =>
		rule.beforeText ? new RegExp(rule.beforeText).test(line) : false,
	);
}

describe('VBA language configuration', () => {
	it('uses JavaScript-compatible indentation, folding, and enter regexes', () => {
		const config = loadConfig();
		const patterns = [
			...Object.values(config.indentationRules ?? {}),
			...Object.values(config.folding?.markers ?? {}),
			...(config.onEnterRules ?? []).flatMap((rule) =>
				[rule.beforeText, rule.afterText].filter((pattern): pattern is string => Boolean(pattern)),
			),
		];

		for (const pattern of patterns) {
			expect(() => new RegExp(pattern)).not.toThrow();
		}
	});

	it('keeps static editor block indentation aligned with smart block openers', () => {
		const config = loadConfig();
		const increase = new RegExp(config.indentationRules?.increaseIndentPattern ?? '');
		const cases = [
			'sub foo()',
			'public function total() as long',
			'property get name() as string',
			'if ready then',
			'for i = 1 to 10',
			'for each item in collection',
			'do',
			'do while ready',
			'do until ready',
			'while ready',
			'with activesheet',
			'select case value',
			'type tpoint',
			'enum color',
			'#if vba7 then',
		];

		for (const line of cases) {
			expect(detectSmartBlockOpener(line), line).toBeDefined();
			expect(increase.test(`    ${line}`), line).toBe(true);
			expect(enterRuleMatches(config, `    ${line}`), line).toBe(true);
		}
	});

	it('does not indent incomplete block openers that smart enter rejects', () => {
		const config = loadConfig();
		const increase = new RegExp(config.indentationRules?.increaseIndentPattern ?? '');
		const cases = [
			'if ready then value = 1',
			'if then',
			'for',
			'for i = 1',
			'for each item in',
			'do while',
			'do until',
			'while',
			'with',
			'select case',
			'#if vba7',
			'declare sub sleep lib "kernel32" ()',
		];

		for (const line of cases) {
			expect(detectSmartBlockOpener(line), line).toBeUndefined();
			expect(increase.test(`    ${line}`), line).toBe(false);
			expect(enterRuleMatches(config, `    ${line}`), line).toBe(false);
		}
	});
});
