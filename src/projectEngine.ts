// Native project engine.
//
// Exposes the request/response surface the extension already speaks
// (`call(method, params)`) but answers every method in-process from
// src/vba/**, with no external runtime. Keeping the call shape means every
// existing caller - explorer, file system, analysis, tests, agent tools - works
// unchanged.

import * as vscode from 'vscode';
import { enginePriming } from './enginePriming';
import { ProjectEngineError } from './projectEngineErrors';
import * as svc from './vba/projectService';
import type { CellValue } from './vba/xlsx';
import type { ShapeEdit } from './vba/xlsxShapes';

type Params = Record<string, unknown>;

function str(params: Params, key: string): string {
	const value = params[key];
	if (typeof value !== 'string' || value.length === 0) {
		throw new ProjectEngineError(`Missing required '${key}' parameter.`, -32602);
	}
	return value;
}

/** A text parameter that may be empty: a blank document is still a document. */
function text(params: Params, key: string): string {
	const value = params[key];
	if (typeof value !== 'string') {
		throw new ProjectEngineError(`Missing required '${key}' parameter.`, -32602);
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
		throw new ProjectEngineError(`Missing required '${key}' parameter.`, -32602);
	}
	return value.map((row) => (Array.isArray(row) ? row : [row]) as CellValue[]);
}

/** An optional text parameter; an empty string is kept, since it clears a value. */
function optionalText(params: Params, key: string): string | undefined {
	const value = params[key];
	if (value === undefined || value === null) { return undefined; }
	if (typeof value !== 'string') {
		throw new ProjectEngineError(`The '${key}' parameter must be text.`, -32602);
	}
	return value;
}

function shapeEdit(params: Params): ShapeEdit {
	const action = str(params, 'action');
	if (action !== 'add' && action !== 'update' && action !== 'delete') {
		throw new ProjectEngineError(`The 'action' parameter must be add, update or delete, not '${action}'.`, -32602);
	}
	const edit: ShapeEdit = { action };
	for (const key of ['name', 'type', 'range', 'text', 'macro', 'linkedCell', 'inputRange', 'altText', 'newName'] as const) {
		const value = optionalText(params, key);
		if (value !== undefined) { (edit as unknown as Record<string, string>)[key] = value; }
	}
	// Where a slide or a document puts a shape: points, not cells.
	for (const key of ['left', 'top', 'width', 'height'] as const) {
		const value = params[key];
		if (value === undefined || value === null) { continue; }
		const points = typeof value === 'number' ? value : Number(value);
		if (!Number.isFinite(points)) {
			throw new ProjectEngineError(`The '${key}' parameter must be a number of points, not '${String(value)}'.`, -32602);
		}
		edit[key] = points;
	}
	return edit;
}

export class ProjectEngine implements vscode.Disposable {
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
	/**
	 * The template whose VBA project a file with none starts from: the blank
	 * file of its own format. A slideshow or an add-in has no blank file of
	 * its own, and takes a presentation's project, which is the same project;
	 * its package stays its own. A legacy file cannot take one at all.
	 */
	private addVbaTemplatePathFor(targetPath: string): string {
		const extension = /\.([a-z0-9]+)$/i.exec(targetPath)?.[1]?.toLowerCase() ?? '';
		const modern: Record<string, string> = { xls: 'xlsm', xlt: 'xltm', xla: 'xlam', doc: 'docm', dot: 'dotm', ppt: 'pptm', ppa: 'ppam' };
		if (modern[extension]) {
			throw new ProjectEngineError(
				`XLIDE adds a VBA project only to the macro-enabled Office Open XML formats, and a .${extension} file is the older binary format. `
				+ `Save it as .${modern[extension]} first, or write the first macro in the application.`,
				-32602,
			);
		}
		const sameProjectAs = extension === 'ppsm' || extension === 'ppam' ? targetPath.replace(/\.[^.]+$/, '.pptm') : targetPath;
		return this.templatePathFor(sameProjectAs);
	}

	private templatePathFor(targetPath: string): string {
		const lower = targetPath.toLowerCase();
		const extension = /\.([a-z0-9]+)$/.exec(lower)?.[1] ?? '';
		if (['xls', 'xlt', 'xla', 'doc', 'dot', 'ppt', 'ppa'].includes(extension)) {
			throw new ProjectEngineError(
				`Creating new legacy-format files (.${extension}) is not supported; create the modern macro format and use the Office app to save down.`,
				-32602,
			);
		}
		if (extension === 'ppam') {
			throw new ProjectEngineError(
				'PowerPoint add-ins are saved from a presentation; create a .pptm and use PowerPoint to save it as an add-in.',
				-32602,
			);
		}
		if (extension === 'ppsm') {
			// A slideshow's package content type differs from a presentation's,
			// and renaming a .pptm cannot change it - there is no authored
			// .ppsm template to seed from.
			throw new ProjectEngineError(
				'Creating .ppsm slideshows is not supported; create a .pptm and use PowerPoint to save it as a slideshow.',
				-32602,
			);
		}
		if (['xlsx', 'xltx', 'docx', 'dotx', 'pptx', 'potx', 'csv'].includes(extension)) {
			throw new ProjectEngineError(
				`.${extension} is not a macro-enabled format, so it cannot hold a VBA project; use .xlsm, .docm, or .pptm instead.`,
				-32602,
			);
		}
		const templates: Record<string, string> = {
			xlsb: 'blank.xlsb',
			xlam: 'blank.xlam',
			xltm: 'blank.xltm',
			docm: 'blank.docm',
			dotm: 'blank.dotm',
			pptm: 'blank.pptm',
			potm: 'blank.potm',
			// An Access database and an Access add-in are the same file with
			// different extensions; the blank one Access writes has an empty
			// VBA project, which is what a new one starts from.
			accdb: 'blank.accdb',
			accda: 'blank.accdb',
			mdb: 'blank.mdb',
			mda: 'blank.mdb',
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
		// The engine is synchronous and, in a browser, cannot reach a file on
		// its own: the container is loaded before dispatch and whatever the
		// call wrote is sent back after. Both are no-ops on a desktop, where
		// the engine uses node:fs directly. See src/enginePriming.ts.
		await enginePriming.prime(ProjectEngine.enginePathsOf(p));
		let result: T;
		try {
			result = this.dispatch(method, p) as T;
		} catch (err) {
			// A failed call leaves no finished container behind, so its
			// writes go no further.
			enginePriming.discard();
			if (err instanceof ProjectEngineError || err instanceof vscode.CancellationError) {
				throw err;
			}
			const message = err instanceof Error ? err.message : String(err);
			this._out?.appendLine(`[project] ${method} failed: ${message}`);
			throw new ProjectEngineError(message, -32000);
		}
		await enginePriming.flush();
		return result;
	}

	/**
	 * The files a call will read. Every engine method names its container in
	 * `path`; nothing else in the parameter set is a file path, so this is the
	 * whole list. A method that later reaches a second file - a VB6 project's
	 * referenced modules, say - has to name it here, or the browser build
	 * fails loudly on the un-primed path rather than reading it as empty.
	 */
	private static enginePathsOf(p: Params): string[] {
		return typeof p.path === 'string' && p.path.length > 0 ? [p.path] : [];
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
			case 'formDesignerOp':
				return svc.applyFormDesignerOp(str(p, 'path'), str(p, 'module'), p.op as never);
			case 'vb6FormDesignerOp':
				return svc.applyVb6FormDesignerOp(str(p, 'path'), text(p, 'text'), p.op as never,
					typeof p.pendingBytes === 'number' ? p.pendingBytes : 0);
			case 'vb6AppendSidecar':
				return svc.appendVb6Sidecar(str(p, 'path'), str(p, 'file'),
					typeof p.base === 'number' ? p.base : 0,
					Array.isArray(p.records) ? (p.records as string[]) : []);
			case 'readFormDesignerSnapshot':
				return svc.readFormDesignerSnapshot(str(p, 'path'), str(p, 'module'));
			case 'restoreFormDesignerSnapshot':
				return svc.restoreFormDesignerSnapshot(str(p, 'path'), str(p, 'module'), p.streams as Record<string, string>);
			case 'readFormPreview':
				return svc.readFormPreview(str(p, 'path'), str(p, 'module'),
					typeof p.selected === 'string' ? p.selected : undefined,
					typeof p.markup === 'string' ? p.markup : undefined,
					typeof p.identityPath === 'string' ? p.identityPath : undefined);
			case 'readVb6FormPreview':
				return svc.readVb6FormPreview(str(p, 'path'), text(p, 'text'),
					typeof p.selected === 'string' ? p.selected : undefined,
					typeof p.vbpPath === 'string' ? p.vbpPath : undefined,
					p.pending && typeof p.pending === 'object' ? (p.pending as never) : undefined);
			case 'readFormMarkup':
				return svc.readFormMarkup(str(p, 'path'), str(p, 'module'));
			case 'duplicateFormControls':
				return svc.duplicateFormControls(
					str(p, 'path'), str(p, 'module'),
					Array.isArray(p.names) ? (p.names as string[]) : [],
					typeof p.offsetPt === 'number' ? p.offsetPt : undefined);
			case 'removeFormControls':
				return svc.removeFormControls(
					str(p, 'path'), str(p, 'module'),
					Array.isArray(p.names) ? (p.names as string[]) : []);
			case 'applyFormMarkup':
				return svc.applyFormMarkup(str(p, 'path'), str(p, 'module'), str(p, 'markup'));
			case 'addForm':
				return svc.addFormModule(
					str(p, 'path'), str(p, 'module'),
					typeof p.source === 'string' ? p.source : '',
					p.kind === 'report' ? 'report' : p.kind === 'form' ? 'form' : undefined,
				);
			case 'readFormExport':
				return svc.readFormExport(str(p, 'path'), str(p, 'module'));
			case 'writeFormDesigner':
				return svc.writeFormDesigner(
					str(p, 'path'),
					str(p, 'module'),
					Buffer.from(String(p.frxBase64 ?? ''), 'base64'),
					typeof p.frmDesignerBlock === 'string' ? p.frmDesignerBlock : undefined,
				);

			// --- project structure ---
			case 'getProtectionInfo':
				return svc.getProtectionInfo(str(p, 'path'));
			case 'getModulesAndProtectionInfo':
				return svc.getModulesAndProtectionInfo(str(p, 'path'));
			case 'getProjectInfo':
				return svc.getProjectInfo(str(p, 'path'));
			// Whether the file holds a VBA project at all, as opposed to
			// holding one with nothing in it. Both list no modules, and only
			// the second can take a new one.
			case 'hasVbaProject':
				return { hasVbaProject: svc.hasVbaProject(str(p, 'path')) };
			case 'validateProject':
				return svc.validateProject(str(p, 'path'));
			case 'createProject': {
				const target = str(p, 'path');
				return svc.createProject(target, this.templatePathFor(target));
			}
			case 'addVbaProject': {
				const target = str(p, 'path');
				return svc.addVbaProject(target, this.addVbaTemplatePathFor(target));
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

			case 'listReferences':
				return { references: svc.listReferences(str(p, 'path')) };
			case 'addReference':
				return svc.addReference(str(p, 'path'), str(p, 'library'));
			case 'removeReference':
				return svc.removeReference(str(p, 'path'), str(p, 'library'));

			// --- shapes ---
			// `sheet` is the old name for `surface`, still accepted.
			case 'listShapes':
				return svc.listShapes(str(p, 'path'), optionalText(p, 'surface') || optionalText(p, 'sheet') || undefined);
			case 'editShape':
				return svc.editShape(
					str(p, 'path'),
					optionalText(p, 'surface') || optionalText(p, 'sheet') || '',
					shapeEdit(p),
				);

			default:
				throw new ProjectEngineError(`Method not found: ${method}`, -32601);
		}
	}

	dispose(): void {
		// Nothing to tear down: every call runs in-process.
	}
}
