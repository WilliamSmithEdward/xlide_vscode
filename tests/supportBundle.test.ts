import { describe, expect, it } from 'vitest';
import {
	anonymizedWorkbookAnalysisReportFromResult,
	buildSupportBundle,
	defaultSupportBundleFileName,
	redactPath,
	supportBundleDisclosureText,
	supportDiagnosticsText,
	type SupportBundleInput,
} from '../src/supportBundle';

function baseInput(overrides: Partial<SupportBundleInput> = {}): SupportBundleInput {
	return {
		generatedAt: '2026-06-01T12:00:00.000Z',
		extension: {
			id: 'WilliamSmithE.xlide',
			name: 'XLIDE: VBA for VS Code',
			version: '1.0.9',
		},
		vscode: {
			version: '1.95.0',
			appName: 'Visual Studio Code',
		},
		runtime: {
			platform: 'win32',
			arch: 'x64',
			node: 'v20.0.0',
		},
		workspace: {
			folderCount: 1,
		},
		settings: [
			{ key: 'xlide.pythonPath', value: 'C:\\Tools\\Python\\python.exe', source: 'machine' },
			{ key: 'xlide.diagnostics.enabled', value: true, source: 'default' },
			{ key: 'xlide.docs.enabled', value: false, source: 'machine' },
			{ key: 'xlide.editor.blockLayout', value: 'comfy', source: 'default' },
		],
		workbook: {
			available: true,
			workbookPath: 'C:\\Users\\William\\Documents\\ClientWorkbook.xlsm',
			extension: '.xlsm',
			moduleCount: 3,
			moduleTypes: { standard: 2, class: 1 },
			activeModuleType: 'standard',
		},
		analysis: {
			available: true,
			moduleType: 'standard',
			errorCount: 1,
			warningCount: 2,
			suppressedCount: 3,
			byCode: { 'unknown-call': 1, 'option-explicit-missing': 2 },
		},
		commands: [],
		...overrides,
	};
}

describe('support bundle', () => {
	it('redacts paths while preserving useful extension shape', () => {
		expect(redactPath('C:\\Users\\William\\repo\\Book.xlsm')).toBe('<redacted>.xlsm');
		expect(redactPath('C:\\Users\\William\\repo')).toBe('<redacted>');
	});

	it('builds a non-source support snapshot with redacted settings and workbook path', () => {
		const bundle = buildSupportBundle(baseInput());

		expect(bundle.schemaVersion).toBe(1);
		expect(bundle.setup).toMatchObject({
			diagnosticsEnabled: true,
			docsEnabled: false,
			pythonPathConfigured: true,
			excelComStatus: 'available-on-windows-not-checked',
		});
		expect(bundle.settings.map((setting) => setting.key)).toEqual([
			'xlide.diagnostics.enabled',
			'xlide.docs.enabled',
			'xlide.editor.blockLayout',
			'xlide.pythonPath',
		]);
		expect(bundle.settings.find((setting) => setting.key === 'xlide.pythonPath')?.value).toBe(
			'<redacted>.exe',
		);
		expect(bundle.workbook.workbookPath).toBe('<redacted>.xlsm');
		expect(bundle.workbook.moduleTypes).toEqual({ standard: 2, class: 1 });
		expect(bundle.privacy).toEqual({
			workbookSourceIncluded: false,
			pathsRedacted: true,
			commandArgumentsIncluded: false,
			writeAuditIncluded: true,
			anonymizedAnalysisReportIncluded: false,
			selectedLogsIncluded: false,
			logPathsRedacted: true,
		});
		expect(bundle.anonymizedReports.workbookAnalysis).toEqual({
			included: false,
			unavailableReason: 'not-requested',
		});
		expect(bundle.selectedLogs).toEqual({
			included: false,
			entries: [],
		});
	});

	it('keeps only the latest command-log entries', () => {
		const commands = Array.from({ length: 30 }, (_, index) => ({
			timestamp: `2026-06-01T12:00:${String(index).padStart(2, '0')}.000Z`,
			command: `xlide.command${index}`,
			outcome: 'succeeded' as const,
			durationMs: index,
		}));

		const bundle = buildSupportBundle(baseInput({ commands }));

		expect(bundle.recentCommands).toHaveLength(25);
		expect(bundle.recentCommands[0].command).toBe('xlide.command5');
		expect(bundle.recentCommands[24].command).toBe('xlide.command29');
	});

	it('uses a deterministic support bundle filename', () => {
		expect(defaultSupportBundleFileName(new Date('2026-06-01T12:34:56.789Z'))).toBe(
			'xlide-support-2026-06-01T12-34-56-789Z.json',
		);
	});

	it('describes support-bundle contents before export without exposing source or paths', () => {
		const bundle = buildSupportBundle(baseInput());
		const disclosure = supportBundleDisclosureText(bundle);

		expect(disclosure).toContain('Included:');
		expect(disclosure).toContain('Not included:');
		expect(disclosure).toContain('Workbook VBA source or cell data');
		expect(disclosure).toContain('Anonymized workbook analysis report');
		expect(disclosure).toContain('<redacted>.xlsm');
		expect(disclosure).not.toContain('C:\\Users\\William');
	});

	it('formats copyable diagnostics from the same redacted support model', () => {
		const bundle = buildSupportBundle(baseInput({
			commands: [
				{
					timestamp: '2026-06-01T12:00:00.000Z',
					command: 'xlide.setup',
					outcome: 'failed',
					durationMs: 50,
					errorCategory: 'python-backend',
				},
			],
		}));
		const text = supportDiagnosticsText(bundle);

		expect(text).toContain('XLIDE Diagnostics');
		expect(text).toContain('xlide.setup | failed');
		expect(text).toContain('errorCategory=python-backend');
		expect(text).toContain('xlide.pythonPath (machine): <redacted>.exe');
		expect(text).toContain('Workbook source included: false');
		expect(text).not.toContain('C:\\Users\\William');
	});

	it('includes recent write audit entries with paths redacted', () => {
		const bundle = buildSupportBundle(baseInput({
			writeAudits: [
				{
					timestamp: '2026-06-01T12:00:00.000Z',
					command: 'xlide.exportModulesToFolder',
					operation: 'export-modules',
					outcome: 'succeeded',
					workbookPath: 'C:\\Users\\William\\Documents\\ClientWorkbook.xlsm',
					targetPath: 'C:\\Users\\William\\Documents\\repo',
					summary: 'Export modules: 2 changed',
				},
			],
		}));
		const text = supportDiagnosticsText(bundle);

		expect(bundle.recentWriteAudits[0].workbookPath).toBe('<redacted>.xlsm');
		expect(bundle.recentWriteAudits[0].targetPath).toBe('<redacted>');
		expect(text).toContain('xlide.exportModulesToFolder | export-modules | succeeded');
		expect(text).not.toContain('C:\\Users\\William');
	});

	it('includes selected redacted logs only when provided', () => {
		const bundle = buildSupportBundle(baseInput({
			selectedLogs: [
				{
					timestamp: '2026-06-01T12:00:00.000Z',
					line: 'Reading C:\\Users\\William\\Documents\\ClientWorkbook.xlsm',
				},
			],
		}));

		expect(bundle.selectedLogs.included).toBe(true);
		expect(bundle.selectedLogs.entries[0].line).toBe('Reading <redacted>.xlsm');
		expect(bundle.privacy.selectedLogsIncluded).toBe(true);
		expect(supportDiagnosticsText(bundle)).not.toContain('C:\\Users\\William');
	});

	it('creates anonymized workbook analysis reports without source paths or module names', () => {
		const report = anonymizedWorkbookAnalysisReportFromResult({
			filePath: 'C:\\Users\\William\\Documents\\ClientWorkbook.xlsm',
			moduleCount: 2,
			errorCount: 1,
			warningCount: 1,
			summary: {
				byCategory: { syntax: 1, semantic: 1 },
				byDiagnosticKind: { 'compile-error': 1, 'runtime-risk': 1 },
				vbeCompileEquivalentCount: 1,
				nonVbeCompileEquivalentCount: 1,
				suppressedCount: 3,
			},
			problems: [
				{
					moduleName: 'CustomerPricing',
					moduleType: 'standard',
					severity: 'error',
					code: 'missing-block-closer',
					category: 'syntax',
					diagnosticKind: 'compile-error',
					vbeCompileEquivalent: true,
				},
				{
					moduleName: 'SecretClientModel',
					moduleType: 'class',
					severity: 'warning',
					code: 'missing-return-assignment',
					category: 'semantic',
					diagnosticKind: 'runtime-risk',
					vbeCompileEquivalent: false,
				},
			],
		});
		const bundle = buildSupportBundle(baseInput({ anonymizedWorkbookAnalysisReport: report }));
		const json = JSON.stringify(bundle);

		expect(bundle.privacy.anonymizedAnalysisReportIncluded).toBe(true);
		expect(bundle.anonymizedReports.workbookAnalysis).toMatchObject({
			included: true,
			workbookExtension: '.xlsm',
			moduleCount: 2,
			problemCount: 2,
			suppressedCount: 3,
			byCode: {
				'missing-block-closer': 1,
				'missing-return-assignment': 1,
			},
		});
		expect(json).not.toContain('CustomerPricing');
		expect(json).not.toContain('SecretClientModel');
		expect(json).not.toContain('C:\\Users\\William');
	});
});
