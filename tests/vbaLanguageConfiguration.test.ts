import { readFileSync } from 'fs';
import { describe, expect, it } from 'vitest';
import { detectSmartBlockOpener, VBA_SMART_BLOCK_SNIPPETS } from '../src/vbaLinter';

interface VbaLanguageConfiguration {
	indentationRules?: Record<string, string>;
	folding?: { markers?: Record<string, string> };
	onEnterRules?: Array<{ beforeText?: string; afterText?: string }>;
}

interface PackageConfiguration {
	contributes?: {
		configuration?: {
			properties?: Record<string, unknown>;
		};
	};
}

function loadConfig(): VbaLanguageConfiguration {
	return JSON.parse(
		readFileSync('language-configuration/vba-language-configuration.json', 'utf8'),
	) as VbaLanguageConfiguration;
}

function loadPackage(): PackageConfiguration {
	return JSON.parse(readFileSync('package.json', 'utf8')) as PackageConfiguration;
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
		const cases = Array.from(new Set(
			VBA_SMART_BLOCK_SNIPPETS
				.map((spec) => spec.smartEnterExample)
				.filter((line): line is string => Boolean(line)),
		));

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

	it('contributes the shared Smart Enter/snippet block layout setting', () => {
		const setting = loadPackage()
			.contributes
			?.configuration
			?.properties
			?.['xlide.editor.blockLayout'] as { enum?: string[]; default?: string } | undefined;

		expect(setting?.default).toBe('comfy');
		expect(setting?.enum).toEqual(['comfy', 'compact']);
	});
});
