import { readFileSync } from 'fs';
import { describe, expect, it } from 'vitest';
import { detectSmartBlockOpener } from '../src/vbaSmartEnter';
import { VBA_SMART_BLOCK_SNIPPETS } from '../src/vbaSmartBlockSnippets';
import { XLIDE_VBA_EDITOR_OVERRIDES } from '../src/xlideVbaEditorOverrides';

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
			commandPalette?: PackageMenuContribution[];
		} & Record<string, PackageMenuContribution[] | undefined>;
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

	it('indents after a continuation underscore, whitespace after it included (issue #83)', () => {
		const config = loadConfig();
		expect(enterRuleMatches(config, '    total = a + _')).toBe(true);
		expect(enterRuleMatches(config, '    total = a + _  ')).toBe(true);
		expect(enterRuleMatches(config, '    total = a_')).toBe(false);
	});

	it('outdents the one-word EndIf like End If (issue #88)', () => {
		const decrease = new RegExp(loadConfig().indentationRules?.decreaseIndentPattern ?? '');
		for (const line of ['    EndIf', '    endif', '    End If']) {
			expect(decrease.test(line), line).toBe(true);
		}
		expect(decrease.test('    EndIfDone = 1')).toBe(false);
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
		expect(contributes?.configurationDefaults?.['[xlide-vba]']).toEqual(Object.fromEntries(
			XLIDE_VBA_EDITOR_OVERRIDES.map(({ key, value }) => [`editor.${key}`, value]),
		));
		// Every file a VB6 project names as a module, so a UserControl or a
		// PropertyPage opened from disk colours and completes like a form.
		expect(standaloneLanguage?.extensions).toEqual(['.bas', '.cls', '.frm', '.ctl', '.pag', '.dsr']);
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

	it('registers the form designer as the DEFAULT editor for .form documents', () => {
		// The designer and the markup are one unit: opening a .form face
		// anywhere lands in the custom editor, and F5 works from inside it.
		const pkg = loadPackage() as PackageConfiguration & {
			contributes?: { customEditors?: Array<{ viewType?: string; selector?: Array<{ filenamePattern?: string }>; priority?: string }> };
		};
		const editor = pkg.contributes?.customEditors?.find((e) => e.viewType === 'xlideFormDesigner');
		expect(editor?.priority).toBe('default');
		expect(editor?.selector?.some((s) => s.filenamePattern === '*.form')).toBe(true);
		expect(pkg.contributes?.keybindings?.some((k) =>
			k.command === 'xlide.launchFormHost' && k.when?.includes('xlideFormDesigner'))).toBe(true);
		// The old webview-panel undo bindings are gone: undo is text undo now.
		expect(pkg.contributes?.commands?.some((c) => c.command === 'xlide.designerUndo')).toBe(false);
	});

	it('lets the smartTab keybinding yield to a visible inline suggestion', () => {
		// When AI ghost text (an inline suggestion) is showing, Tab must reach
		// VS Code's editor.action.inlineSuggest.commit to accept it, not be
		// hijacked by smartTab's indentation. The guard is `!inlineSuggestionVisible`.
		const smartTab = loadPackage().contributes?.keybindings?.find(
			(entry) => entry.command === 'xlide.vba.smartTab',
		);
		expect(smartTab?.key).toBe('tab');
		expect(smartTab?.when).toContain('!inlineSuggestionVisible');
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
			'xlide.agent.showWriteDiffs',
			'xlide.analysis.ignoreFilesOutsideTree',
			'xlide.analysis.ruleSeverityOverrides',
			'xlide.analysis.untrackedRules',
			'xlide.analysis.visibleSeverities',
			// The names the Office integration settings had while they only
			// served Excel: still contributed, marked deprecated, so a value a
			// user already stored keeps working and is not flagged as unknown.
			'xlide.attachToRunningExcel',
			'xlide.diagnostics.enabled',
			'xlide.docs.enabled',
			'xlide.docs.metadataGlob',
			'xlide.editor.blockLayout',
			'xlide.editor.continueCommentOnNewline',
			'xlide.editor.mirrorCommentSpacing',
			'xlide.excelIntegration.coordinationMode',
			'xlide.excelIntegration.reopenAfterClose',
			'xlide.excelIntegration.reopenMode',
			'xlide.excelIntegration.reopenReadOnlyAfterSave',
			'xlide.excelIntegration.trackOpenedWorkbooks',
			'xlide.explorer.autoExpandCollapse',
			'xlide.explorer.view',
			'xlide.formRun.injectShowMacro',
			'xlide.officeIntegration.attachToRunning',
			'xlide.officeIntegration.coordinationMode',
			'xlide.officeIntegration.reopenAfterClose',
			'xlide.officeIntegration.reopenMode',
			'xlide.officeIntegration.reopenReadOnlyAfterSave',
			'xlide.officeIntegration.trackOpenedFiles',
			'xlide.performance.trace',
		]);

		const deprecated = xlideSettings.filter(([, setting]) => setting.deprecationMessage !== undefined);
		expect(deprecated.map(([key]) => key)).toEqual([
			'xlide.attachToRunningExcel',
			'xlide.excelIntegration.coordinationMode',
			'xlide.excelIntegration.reopenAfterClose',
			'xlide.excelIntegration.reopenMode',
			'xlide.excelIntegration.reopenReadOnlyAfterSave',
			'xlide.excelIntegration.trackOpenedWorkbooks',
		]);
		for (const [key, setting] of deprecated) {
			expect(setting.deprecationMessage, key).toMatch(/^Renamed to xlide\.officeIntegration\./);
		}

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
		// The workbook engine runs in-process, so no setup gate stands between
		// the user and the tree: the only welcome view is the empty-workspace one,
		// and it is unconditional.
		expect(contributes?.viewsWelcome?.filter((entry) => entry.view === 'xlide.explorer')).toEqual([
			{
				view: 'xlide.explorer',
				contents: expect.stringContaining('No macro-enabled Office files'),
			},
		]);
	});

	it('contributes the project settings command used by the XLIDE sidebar', () => {
		const commands = loadPackage()
			.contributes
			?.commands ?? [];
		const command = commands.find((entry) => entry.command === 'xlide.openProjectSettings');
		const globalCommand = commands.find((entry) => entry.command === 'xlide.openGlobalSettings');
		const runVbaTestsCommand = commands.find((entry) => entry.command === 'xlide.runVbaTests');
		const performanceCommand = commands.find((entry) => entry.command === 'xlide.copyPerformanceSnapshot');

		expect(command).toMatchObject({
			command: 'xlide.openProjectSettings',
			title: 'Open Project Settings',
			category: 'XLIDE',
		});
		expect(globalCommand).toMatchObject({
			command: 'xlide.openGlobalSettings',
			title: 'Open Global Settings',
			category: 'XLIDE',
		});
		expect(runVbaTestsCommand).toMatchObject({
			command: 'xlide.runVbaTests',
			title: 'Unit Tests',
			category: 'XLIDE',
		});
		expect(performanceCommand).toMatchObject({
			command: 'xlide.copyPerformanceSnapshot',
			title: 'Copy Performance Snapshot',
			category: 'XLIDE',
		});
	});

	it('keeps project tree tests centralized through the Unit Tests GUI', () => {
		// Workbook nodes carry contextValue 'xlsm'; other containers use
		// 'macroDocument'/'macroReadOnly', so a workbook-tree command is any
		// entry whose when-clause matches the xlsm context value.
		const matchesXlsm = (when: string | undefined): boolean =>
			when !== undefined
			&& when.startsWith('view == xlide.explorer && viewItem')
			&& /viewItem (== xlsm$|=~ .*[(|]xlsm[)|])/.test(when);
		const projectTreeCommands = loadPackage()
			.contributes
			?.menus
			?.['view/item/context']
			?.filter((entry) => matchesXlsm(entry.when))
			.map((entry) => entry.command) ?? [];

		expect(projectTreeCommands).toContain('xlide.analyzeProject');
		expect(projectTreeCommands).toContain('xlide.runVbaTests');
		expect(projectTreeCommands).not.toContain('xlide.validateProject');
	});

	it('gives a module row no Open Module button, since clicking the row opens it', () => {
		const moduleRowButtons = loadPackage()
			.contributes
			?.menus
			?.['view/item/context']
			?.filter((entry) => entry.group?.startsWith('inline') && entry.when?.includes('viewItem =~ /^module-/'))
			.map((entry) => entry.command) ?? [];

		expect(moduleRowButtons).not.toContain('xlide.openModule');
	});

	it('offers Rename and Delete on a form row, and never on a document module', () => {
		// A form's designer now follows a rename and goes with a delete, and a
		// rename updates the code that uses the form. A document module's name
		// is its document's, and the engine refuses both.
		const offeredOn = (command: string, contextValue: string): boolean => (loadPackage()
			.contributes
			?.menus
			?.['view/item/context'] ?? [])
			.filter((entry) => entry.command === command)
			.some((entry) => {
				const pattern = /viewItem =~ \/(.+)\/$/.exec(entry.when ?? '')?.[1];
				return pattern !== undefined && new RegExp(pattern).test(contextValue);
			});
		for (const command of ['xlide.renameModule', 'xlide.deleteModule']) {
			for (const row of ['module-userform', 'module-userform-agent-pending', 'module-standard', 'module-class']) {
				expect(offeredOn(command, row), `${command} on ${row}`).toBe(true);
			}
			expect(offeredOn(command, 'module-document'), `${command} on module-document`).toBe(false);
		}
	});

	it('puts Review, Keep and Revert on the hover bar of a row an agent changed', () => {
		const pendingRowButtons = loadPackage()
			.contributes
			?.menus
			?.['view/item/context']
			?.filter((entry) => entry.group?.startsWith('inline') && entry.when?.includes('-agent-pending$/'))
			.sort((a, b) => (a.group ?? '').localeCompare(b.group ?? ''))
			.map((entry) => entry.command) ?? [];

		expect(pendingRowButtons).toEqual(['xlide.reviewAgentChange', 'xlide.keepAgentChange', 'xlide.revertAgentChange']);
	});

	it('keeps commands that need a tree row out of the Command Palette', () => {
		// Picked from the palette with no row, these returned at once and did
		// nothing. A row-only command added later has to decide the same way.
		const manifest = loadPackage().contributes;
		const menus = manifest?.menus ?? {};
		const palette = new Map((menus.commandPalette ?? []).map((entry) => [entry.command, entry.when]));
		const elsewhere = new Set(Object.entries(menus)
			.filter(([menu]) => menu !== 'view/item/context' && menu !== 'commandPalette')
			.flatMap(([, entries]) => (entries ?? []).map((entry) => entry.command)));
		const rowOnly = [...new Set((menus['view/item/context'] ?? []).map((entry) => entry.command))]
			.filter((command) => command !== undefined && !elsewhere.has(command));
		// These answer a palette pick themselves: the open form or project, or
		// a message saying what to select.
		const answersThePalette = new Set(['xlide.previewForm', 'xlide.analyzeProject', 'xlide.runVbaTests']);

		expect(rowOnly.filter((command) => !palette.has(command!) && !answersThePalette.has(command!))).toEqual([]);
		for (const command of [
			'xlide.openModule', 'xlide.findReferences', 'xlide.newModule', 'xlide.newClassModule',
			'xlide.newUserForm', 'xlide.newAccessForm', 'xlide.newAccessReport', 'xlide.openFormMarkup',
			'xlide.deleteModule', 'xlide.renameModule',
			'xlide.reviewAgentChange', 'xlide.keepAgentChange', 'xlide.revertAgentChange',
		]) {
			expect(palette.get(command), command).toBe('false');
		}
		// These fall back to the open module, so they show while one is open -
		// and only on a desktop, since the browser build does not register
		// them at all (see src/platformFeatures.ts).
		for (const command of [
			'xlide.openInOfficeApp', 'xlide.openInOfficeAppReadOnly',
			'xlide.exportModulesToFolder', 'xlide.importModulesFromFolder',
		]) {
			expect(palette.get(command), command).toBe('resourceScheme == xlide-vba && !xlide.isWeb');
		}
	});

	it('keeps every command the browser build cannot register out of its palette', () => {
		const menus = loadPackage().contributes?.menus ?? {};
		const palette = new Map((menus.commandPalette ?? []).map((entry) => [entry.command, entry.when]));

		// Registered by platformFeaturesNode and not by platformFeaturesWeb,
		// so in a browser the palette would otherwise offer a command that is
		// not there. A `when` of 'false' already hides it everywhere.
		for (const command of [
			'xlide.openInOfficeApp', 'xlide.openInOfficeAppReadOnly',
			'xlide.runMacroAtCursor', 'xlide.runVbaTests',
			'xlide.exportModulesToFolder', 'xlide.importModulesFromFolder',
			'xlide.exportCurrentModuleToFolder', 'xlide.exportSupportBundle',
			'xlide.compareModuleWithHead', 'xlide.compareModuleWithRevision',
			'xlide.compareProjectWithHead', 'xlide.restoreModuleFromHead',
			'xlide.moduleHistory', 'xlide.launchFormHost',
		]) {
			const when = palette.get(command);
			if (when === undefined) {
				continue; // not offered in the palette at all
			}
			expect(when === 'false' || when.includes('!xlide.isWeb'), `${command}: ${when}`).toBe(true);
		}
	});
});
