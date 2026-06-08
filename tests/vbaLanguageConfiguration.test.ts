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
		configurationDefaults?: Record<string, Record<string, unknown>>;
		languages?: PackageLanguage[];
		grammars?: PackageGrammar[];
		keybindings?: PackageKeybinding[];
		viewsContainers?: {
			activitybar?: PackageViewContainer[];
		};
		views?: Record<string, PackageView[]>;
		viewsWelcome?: PackageViewWelcome[];
		menus?: {
			'view/item/context'?: PackageMenuContribution[];
			'editor/context'?: PackageMenuContribution[];
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

interface PackageLanguage {
	id?: string;
	aliases?: string[];
	extensions?: string[];
	configuration?: string;
}

interface PackageGrammar {
	language?: string;
	scopeName?: string;
	path?: string;
}

interface PackageKeybinding {
	command?: string;
	key?: string;
	when?: string;
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

	it('scopes noisy editor defaults to XLIDE virtual VBA modules only', () => {
		const contributes = loadPackage().contributes;
		const xlideLanguage = contributes?.languages?.find((language) => language.id === 'xlide-vba');
		const standaloneLanguage = contributes?.languages?.find((language) => language.id === 'vba');

		expect(contributes?.configurationDefaults?.['[vba]']).toBeUndefined();
		expect(contributes?.configurationDefaults?.['[xlide-vba]']).toMatchObject({
			'editor.detectIndentation': false,
			'editor.tabSize': 4,
			'editor.minimap.enabled': true,
			'editor.minimap.renderCharacters': false,
			'editor.minimap.showMarkSectionHeaders': false,
			'editor.minimap.showRegionSectionHeaders': false,
			'editor.overviewRulerBorder': false,
			'editor.overviewRulerLanes': 3,
		});
		expect(standaloneLanguage?.extensions).toEqual(['.bas', '.cls', '.frm']);
		expect(xlideLanguage).toMatchObject({
			id: 'xlide-vba',
			configuration: './language-configuration/vba-language-configuration.json',
		});
		expect(xlideLanguage?.extensions).toBeUndefined();
		expect(contributes?.grammars).toEqual(expect.arrayContaining([
			expect.objectContaining({ language: 'vba', scopeName: 'source.vba' }),
			expect.objectContaining({ language: 'xlide-vba', scopeName: 'source.vba' }),
		]));
		expect(contributes?.keybindings?.find((entry) => entry.command === 'xlide.vba.smartBackspace')?.when)
			.toContain('xlide-vba');
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
		const runVbaTestsCommand = commands.find((entry) => entry.command === 'xlide.runVbaTests');

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
		expect(runVbaTestsCommand).toMatchObject({
			command: 'xlide.runVbaTests',
			title: 'Unit Tests',
			category: 'XLIDE',
		});
		expect(commands.map((entry) => entry.command)).not.toEqual(expect.arrayContaining([
			'xlide.runVbaTestsWithFilters',
			'xlide.runVbaTestsInCurrentModule',
			'xlide.runVbaTestAtCursor',
			'xlide.installVbaTestSupport',
		]));
	});

	it('keeps workbook tree tests centralized through the Unit Tests GUI', () => {
		const workbookTreeCommands = loadPackage()
			.contributes
			?.menus
			?.['view/item/context']
			?.filter((entry) => entry.when === 'view == xlide.explorer && viewItem == xlsm')
			.map((entry) => entry.command) ?? [];

		expect(workbookTreeCommands).toContain('xlide.analyzeWorkbook');
		expect(workbookTreeCommands).toContain('xlide.runVbaTests');
		expect(workbookTreeCommands).not.toContain('xlide.runVbaTestsWithFilters');
		expect(workbookTreeCommands).not.toContain('xlide.installVbaTestSupport');
		expect(workbookTreeCommands).not.toContain('xlide.validateWorkbook');

		const editorContextCommands = loadPackage()
			.contributes
			?.menus
			?.['editor/context']
			?.filter((entry) => entry.when === 'resourceScheme == xlide-vba')
			.map((entry) => entry.command) ?? [];
		expect(editorContextCommands).not.toContain('xlide.runVbaTestsInCurrentModule');
		expect(editorContextCommands).not.toContain('xlide.runVbaTestAtCursor');
	});
});
