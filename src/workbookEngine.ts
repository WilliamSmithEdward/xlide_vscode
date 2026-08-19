// Native workbook engine.
//
// Exposes the request/response surface the extension already speaks
// (`call(method, params)`) but answers every method in-process from
// src/vba/**, with no external runtime. Keeping the call shape means every
// existing caller - explorer, file system, analysis, tests, agent tools - works
// unchanged.

import * as vscode from 'vscode';
import { WorkbookEngineError } from './workbookEngineErrors';
import * as svc from './vba/workbookService';
import type { CellValue } from './vba/xlsx';

type Params = Record<string, unknown>;

function str(params: Params, key: string): string {
	const value = params[key];
	if (typeof value !== 'string' || value.length === 0) {
		throw new WorkbookEngineError(`Missing required '${key}' parameter.`, -32602);
	}
	return value;
}

function optionalBool(params: Params, key: string, fallback = false): boolean {
	const value = params[key];
	return typeof value === 'boolean' ? value : fallback;
}

function grid(params: Params, key: string): CellValue[][] {
	const value = params[key];
	if (!Array.isArray(value)) {
		throw new WorkbookEngineError(`Missing required '${key}' parameter.`, -32602);
	}
	return value.map((row) => (Array.isArray(row) ? row : [row]) as CellValue[]);
}

export class WorkbookEngine implements vscode.Disposable {
	constructor(
		private readonly _context: vscode.ExtensionContext,
		private readonly _out?: vscode.OutputChannel,
	) {}

	/**
	 * Path of the bundled blank file used to seed a new one, chosen by the
	 * target's extension. Each template is a file freshly authored by its own
	 * Office application, so the result opens without a repair prompt - and
	 * every format needs its own template: what makes a package an add-in, a
	 * template, or a Word/PowerPoint file at all is its content types, which
	 * renaming cannot change. Anything unrecognized seeds .xlsm, the default;
	 * formats XLIDE cannot author refuse with the reason.
	 */
	private templatePathFor(targetPath: string): string {
		const lower = targetPath.toLowerCase();
		const extension = /\.([a-z0-9]+)$/.exec(lower)?.[1] ?? '';
		if (['xls', 'doc', 'ppt'].includes(extension)) {
			throw new WorkbookEngineError(
				`Creating new legacy-format files (.${extension}) is not supported; create the modern macro format and use the Office app to save down.`,
				-32602,
			);
		}
		if (['accdb', 'accda', 'mdb'].includes(extension)) {
			throw new WorkbookEngineError(
				'Access databases cannot be created by XLIDE (Access files are read-only).',
				-32602,
			);
		}
		const templates: Record<string, string> = {
			xlsb: 'blank.xlsb',
			xlam: 'blank.xlam',
			docm: 'blank.docm',
			dotm: 'blank.dotm',
			pptm: 'blank.pptm',
			potm: 'blank.potm',
		};
		const template = templates[extension] ?? 'blank.xlsm';
		return vscode.Uri.joinPath(
			this._context.extensionUri, 'assets', 'templates', template,
		).fsPath;
	}

	async call<T = unknown>(
		method: string,
		params: unknown = {},
		token?: vscode.CancellationToken,
	): Promise<T> {
		if (token?.isCancellationRequested) {
			throw new vscode.CancellationError();
		}
		const p = (params ?? {}) as Params;
		try {
			return this.dispatch(method, p) as T;
		} catch (err) {
			if (err instanceof WorkbookEngineError || err instanceof vscode.CancellationError) {
				throw err;
			}
			const message = err instanceof Error ? err.message : String(err);
			this._out?.appendLine(`[workbook] ${method} failed: ${message}`);
			throw new WorkbookEngineError(message, -32000);
		}
	}

	private dispatch(method: string, p: Params): unknown {
		switch (method) {
			// --- VBA modules ---
			case 'listModules':
				return svc.listModules(str(p, 'path'));
			case 'readModules':
				return svc.readModules(str(p, 'path'), optionalBool(p, 'full'));
			case 'readModule':
				return svc.readModule(str(p, 'path'), str(p, 'module'), optionalBool(p, 'full'));
			case 'listSubs':
				return svc.listSubs(str(p, 'path'), str(p, 'module'));
			case 'writeModule':
				return svc.writeModule(
					str(p, 'path'),
					str(p, 'module'),
					typeof p.source === 'string' ? p.source : '',
					p.kind === 'class' ? 'class' : 'standard',
				);
			case 'renameModule':
				return svc.renameModule(str(p, 'path'), str(p, 'module'), str(p, 'newName'));
			case 'deleteModule':
				return svc.deleteModule(str(p, 'path'), str(p, 'module'));
			case 'readFormExport':
				return svc.readFormExport(str(p, 'path'), str(p, 'module'));
			case 'writeFormDesigner':
				return svc.writeFormDesigner(
					str(p, 'path'),
					str(p, 'module'),
					Buffer.from(String(p.frxBase64 ?? ''), 'base64'),
					typeof p.frmDesignerBlock === 'string' ? p.frmDesignerBlock : undefined,
				);

			// --- workbook structure ---
			case 'getProtectionInfo':
				return svc.getProtectionInfo(str(p, 'path'));
			case 'getModulesAndProtectionInfo':
				return svc.getModulesAndProtectionInfo(str(p, 'path'));
			case 'getWorkbookInfo':
				return svc.getWorkbookInfo(str(p, 'path'));
			case 'validateWorkbook':
				return svc.validateWorkbook(str(p, 'path'));
			case 'createWorkbook': {
				const target = str(p, 'path');
				return svc.createWorkbook(target, this.templatePathFor(target));
			}

			// --- sheets and cells ---
			case 'listSheets':
				return svc.listSheets(str(p, 'path'));
			case 'readCells':
				return svc.readCells(str(p, 'path'), str(p, 'sheet'), str(p, 'range'));
			case 'readFormulas':
				return svc.readFormulas(str(p, 'path'), str(p, 'sheet'), str(p, 'range'));
			case 'writeCells':
				return svc.writeCells(str(p, 'path'), str(p, 'sheet'), str(p, 'startCell'), grid(p, 'data'));

			default:
				throw new WorkbookEngineError(`Method not found: ${method}`, -32601);
		}
	}

	dispose(): void {
		// Nothing to tear down: every call runs in-process.
	}
}
