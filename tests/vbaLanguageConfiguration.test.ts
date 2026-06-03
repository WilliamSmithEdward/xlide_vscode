import { readFileSync } from 'fs';
import { describe, expect, it } from 'vitest';
import { detectSmartBlockOpener, VBA_SMART_BLOCK_SNIPPETS } from '../src/vbaStructuralAnalysis';

interface VbaLanguageConfiguration {
	indentationRules?: Record<string, string>;
	folding?: { markers?: Record<string, string> };
	onEnterRules?: Array<{ beforeText?: string; afterText?: string }>;
}

interface PackageConfiguration {
	activationEvents?: string[];
	contributes?: {
		commands?: PackageCommand[];
		configuration?: {
			properties?: Record<string, PackageSetting>;
		};
		viewsContainers?: {
			activitybar?: PackageViewContainer[];
		};
		views?: Record<string, PackageView[]>;
		viewsWelcome?: PackageViewWelcome[];
		menus?: {
			'view/item/context'?: PackageMenuContribution[];
		};
	};
}

interface PackageSetting {
	default?: unknown;
	enum?: string[];
	scope?: string;
	additionalProperties?: {
		enum?: string[];
	};
}

interface PackageViewContainer {
	id?: string;
	title?: string;
	icon?: string;
}

interface PackageView {
	id?: string;
	name?: string;
	contextualTitle?: string;
	type?: string;
}

interface PackageViewWelcome {
	view?: string;
	contents?: string;
	when?: string;
}

interface PackageCommand {
	command?: string;
	title?: string;
	category?: string;
}

interface PackageMenuContribution {
	command?: string;
	when?: string;
	group?: string;
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
			?.['xlide.editor.blockLayout'];

		expect(setting?.default).toBe('comfy');
		expect(setting?.enum).toEqual(['comfy', 'compact']);
	});

	it('keeps analysis rule severity overrides guarded by the shared severity vocabulary', () => {
		const setting = loadPackage()
			.contributes
			?.configuration
			?.properties
			?.['xlide.analysis.ruleSeverityOverrides'];

		expect(setting?.default).toEqual({});
		expect(setting?.additionalProperties?.enum).toEqual(['off', 'warning']);
	});

	it('keeps contributed XLIDE settings machine-scoped', () => {
		const settings = loadPackage()
			.contributes
			?.configuration
			?.properties ?? {};
		const xlideSettings = Object.entries(settings)
			.filter(([key]) => key.startsWith('xlide.'))
			.sort(([a], [b]) => a.localeCompare(b));

		expect(xlideSettings.map(([key]) => key)).toEqual([
			'xlide.analysis.ruleSeverityOverrides',
			'xlide.analysis.untrackedRules',
			'xlide.analysis.visibleSeverities',
			'xlide.attachToRunningExcel',
			'xlide.diagnostics.enabled',
			'xlide.docs.enabled',
			'xlide.docs.metadataGlob',
			'xlide.editor.blockLayout',
			'xlide.pythonPath',
		]);

		for (const [key, setting] of xlideSettings) {
			expect(setting.scope, key).toBe('machine');
		}
	});

	it('contributes one dedicated XLIDE activity bar sidebar without replacing the explorer tree', () => {
		const contributes = loadPackage().contributes;
		const activityContainers = contributes?.viewsContainers?.activitybar ?? [];
		const xlideContainer = activityContainers.find((container) => container.id === 'xlide');

		expect(xlideContainer).toMatchObject({
			id: 'xlide',
			title: 'XLIDE',
			icon: 'assets/icons/xlide-activity.svg',
		});
		expect(contributes?.views?.xlide).toEqual([
			expect.objectContaining({
				id: 'xlide.sidebar',
				type: 'webview',
			}),
		]);
		expect(contributes?.views?.explorer?.map((view) => view.id)).toContain('xlide.explorer');
		expect(loadPackage().activationEvents).not.toEqual(expect.arrayContaining([
			'onView:xlide.sidebar',
			'onView:xlide.explorer',
		]));
		expect(contributes?.viewsWelcome?.map((entry) => entry.view)).not.toContain('xlide.sidebar');
		expect(contributes?.viewsWelcome?.filter((entry) => entry.view === 'xlide.explorer')).toEqual([
			expect.objectContaining({
				when: '!xlide.setupComplete',
				contents: expect.stringContaining('XLIDE setup is not complete.'),
			}),
			expect.objectContaining({
				when: 'xlide.setupComplete',
				contents: expect.stringContaining('No Excel workbooks'),
			}),
		]);
	});

	it('contributes the workbook settings command used by the XLIDE sidebar', () => {
		const commands = loadPackage()
			.contributes
			?.commands ?? [];
		const command = commands.find((entry) => entry.command === 'xlide.openWorkbookSettings');
		const globalCommand = commands.find((entry) => entry.command === 'xlide.openGlobalSettings');
		const downloadPythonCommand = commands.find((entry) => entry.command === 'xlide.downloadPython');

		expect(command).toMatchObject({
			command: 'xlide.openWorkbookSettings',
			title: 'Open Workbook Settings',
			category: 'XLIDE',
		});
		expect(globalCommand).toMatchObject({
			command: 'xlide.openGlobalSettings',
			title: 'Open Global Settings',
			category: 'XLIDE',
		});
		expect(downloadPythonCommand).toMatchObject({
			command: 'xlide.downloadPython',
			title: 'Download Python',
			category: 'XLIDE',
		});
	});

	it('keeps Validate VBA Project out of workbook tree context actions', () => {
		const workbookTreeCommands = loadPackage()
			.contributes
			?.menus
			?.['view/item/context']
			?.filter((entry) => entry.when === 'view == xlide.explorer && viewItem == xlsm')
			.map((entry) => entry.command) ?? [];

		expect(workbookTreeCommands).toContain('xlide.analyzeWorkbook');
		expect(workbookTreeCommands).not.toContain('xlide.validateWorkbook');
	});
});
