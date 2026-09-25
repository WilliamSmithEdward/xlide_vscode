// Sheets and shapes in the explorer, where each host keeps them.
//
// An Excel workbook's sheets are a Sheets folder under the project, in tab
// order. A sheet with a module is its module row, named the way the VBE
// names it: `Sheet1 (Budget)`. A sheet with shapes and no module is a row of
// its own. The sheets with neither sit last, in a folder of their own,
// Sheets With No Modules or Shapes. A sheet's shapes are a Shapes folder
// under its row, drawn only when it has one or more. Word's shapes belong to
// the document, so they are a Shapes folder under ThisDocument, one row per
// story (the body, a header, a footer). A PowerPoint slide is no module at
// all, so the slides are a folder of their own under the project.
//
// The sheets come from the workbook's own sheet list (vba/workbookSheets.ts),
// which every Excel format answers, and the shapes from a listing read when
// a sheet row or a Shapes folder is drawn, which the OOXML formats answer.
// Excel gives a worksheet its module only once the workbook's VBA editor has
// been opened after the sheet was added, and a workbook saved before then
// carries sheets with no module - SheetsFixture has three.
//
// Both are kept per project until the file changes or the tree refreshes.
// After a change the sheet list, which is cheap, is read again, and the
// Sheets folder is drawn again only when a sheet was added, removed, renamed
// or given a module; the shapes are read again only for a row someone has
// open. A file that changes on every keystroke's auto-save thus costs the
// tree no drawing parse while nobody is looking at shapes.

import * as vscode from 'vscode';
import * as path from 'path';
import { containerHostForPath } from './macroContainerUi';
import type { ProjectEngine } from './projectEngine';
import type { XlideNode } from './projectExplorer';
import type { ShapeInfo, ShapeKind } from './vba/shapes';
import type { ShapeSurface } from './vba/projectService';
import { canRunMacro, type ShapeHost } from './vba/shapeCapabilities';
import type { WorkbookSheet } from './vba/workbookSheets';
import { projectIdentityKey } from './xlideFileSystem';

/**
 * The host whose shape tools read a file, from its extension; undefined for a
 * file they cannot. The binary formats (.xlsb, .xls, .doc, .ppt) keep shapes
 * in records the engine does not write, and an add-in presentation has no
 * slides.
 */
export function shapeHostForPath(filePath: string): ShapeHost | undefined {
	switch (/\.([a-z0-9]+)$/i.exec(filePath)?.[1]?.toLowerCase()) {
		case 'xlsm': case 'xltm': case 'xlam': return 'excel';
		case 'docm': case 'dotm': return 'word';
		case 'pptm': case 'potm': case 'ppsm': return 'powerpoint';
		default: return undefined;
	}
}

/** What each kind of shape is called on its row. */
const KIND_LABELS: Record<ShapeKind, string> = {
	shape: 'AutoShape', textBox: 'text box', line: 'line', picture: 'picture', chart: 'chart', group: 'group',
	button: 'button', checkBox: 'check box', optionButton: 'option button', dropDown: 'drop-down',
	listBox: 'list box', scrollBar: 'scroll bar', spinner: 'spinner', label: 'label', groupBox: 'group box',
	editBox: 'edit box', formControl: 'form control', activeX: 'ActiveX control', placeholder: 'placeholder',
	table: 'table', canvas: 'drawing canvas', other: 'shape',
};

const KIND_ICONS: Record<ShapeKind, string> = {
	shape: 'primitive-square', textBox: 'symbol-text', line: 'symbol-ruler', picture: 'file-media',
	chart: 'pie-chart', group: 'layers', button: 'run', checkBox: 'checklist', optionButton: 'record',
	dropDown: 'list-unordered', listBox: 'list-unordered', scrollBar: 'symbol-ruler', spinner: 'symbol-ruler',
	label: 'symbol-text', groupBox: 'symbol-structure', editBox: 'symbol-text', formControl: 'symbol-structure',
	activeX: 'window', placeholder: 'layout', table: 'table', canvas: 'layers', other: 'symbol-misc',
};

export function shapeKindLabel(kind: ShapeKind): string {
	return KIND_LABELS[kind] ?? kind;
}

/** Word's note stories hold one story per note, so nothing is added to one from here. */
export function takesNewShapes(host: ShapeHost, surface: string): boolean {
	return host !== 'word' || !['footnotes', 'endnotes', 'comments'].includes(surface.toLowerCase());
}

/** What the tree knows about the file behind a row, for the commands that act on one. */
export interface ShapeRowContext {
	host: ShapeHost;
	/** The surface as the engine names it: a sheet, a slide, a story. */
	surface: string;
	shape?: ShapeInfo;
	/** Whether the shape is inside a group or canvas. */
	inGroup?: boolean;
}

/** The label of the folder that leads a workbook's rows, and of the one that ends its sheets. */
export const SHEETS_FOLDER_LABEL = 'Sheets';
export const BARE_SHEETS_FOLDER_LABEL = 'Sheets With No Modules or Shapes';

type ShapeFolderKind = NonNullable<XlideNode['shapeFolder']>;

/** What a sheet row calls each kind, Excel's way. */
const SHEET_KIND_LABELS: Record<WorkbookSheet['kind'], string> = {
	worksheet: 'worksheet',
	chartsheet: 'chart sheet',
	dialogsheet: 'dialog sheet',
	macrosheet: 'macro sheet',
};

/**
 * The sheet and shape rows of the explorer. The explorer owns the tree and
 * asks this for the rows under a project, a module or one of the rows made
 * here, and for their tree items.
 */
export class ShapeRows {
	private readonly listings = new Map<string, ShapeSurface[]>();
	private readonly stale = new Set<string>();
	private readonly loads = new Map<string, Promise<ShapeSurface[]>>();
	private readonly failures = new Map<string, string>();
	/** Each workbook's sheets, by project, while the tree shows them. */
	private readonly catalogs = new Map<string, WorkbookSheet[]>();
	private readonly catalogLoads = new Map<string, Promise<WorkbookSheet[] | undefined>>();
	private readonly catalogFailures = new Set<string>();
	/** The reason last given for each project whose sheets could not be listed. */
	private readonly catalogFailureSaid = new Map<string, string>();
	/** The sheet list each Sheets folder was last drawn from, as one comparable value. */
	private readonly sheetsDrawn = new Map<string, string>();
	/** One node per folder and per surface, so a redraw finds the row VS Code has. */
	private readonly folders = new Map<string, XlideNode>();
	/** The rows whose shapes were drawn, and so are drawn again when the file changes. */
	private readonly opened = new Set<XlideNode>();
	private readonly parents = new WeakMap<XlideNode, XlideNode>();
	private readonly contexts = new WeakMap<XlideNode, ShapeRowContext>();
	/** The workbook sheet a sheet row stands for. */
	private readonly sheetOfRow = new WeakMap<XlideNode, WorkbookSheet>();
	/** The sheets a Sheets With No Modules or Shapes folder holds, decided when its parent was drawn. */
	private readonly bareSheets = new WeakMap<XlideNode, WorkbookSheet[]>();
	private generation = 0;

	constructor(
		private readonly bridge: ProjectEngine,
		private readonly fire: (node?: XlideNode) => void,
		private readonly out?: vscode.OutputChannel,
	) {}

	/** Forget every listing and row: the tree is being drawn again from the files. */
	clear(): void {
		this.generation++;
		this.listings.clear();
		this.stale.clear();
		this.loads.clear();
		this.failures.clear();
		this.catalogs.clear();
		this.catalogLoads.clear();
		this.catalogFailures.clear();
		this.sheetsDrawn.clear();
		this.folders.clear();
		this.opened.clear();
	}

	/**
	 * A file changed, by a shape edit or a save from outside: its listing is
	 * read again when next asked for, and the shape rows someone opened are
	 * redrawn, which asks. The sheet list, which is cheap, is read again
	 * now, and the Sheets folder redrawn when it differs from what the
	 * folder shows: a sheet added, removed, renamed or given a module.
	 * Nothing else is read, since a file that changes on every keystroke's
	 * auto-save should cost the tree nothing while its shapes are not on
	 * screen.
	 */
	refresh(filePath: string, options: { shapesChanged?: boolean } = {}): void {
		const project = projectIdentityKey(filePath);
		this.failures.delete(project);
		this.loads.delete(project);
		if (this.listings.has(project)) { this.stale.add(project); }
		this.catalogs.delete(project);
		this.catalogLoads.delete(project);
		this.catalogFailures.delete(project);
		for (const node of this.opened) {
			if (projectIdentityKey(node.filePath) === project) { this.fire(node); }
		}
		const sheets = this.folders.get(folderKey(filePath, 'sheets', undefined));
		const drawn = this.sheetsDrawn.get(project);
		if (!sheets || drawn === undefined) {
			return;
		}
		if (options.shapesChanged) {
			// XLIDE's own shape edit: a sheet may have its first shape now, or
			// its last one gone, which moves it between the folder's rows.
			this.fire(sheets);
			return;
		}
		void this.catalog(filePath).then((catalog) => {
			if (catalog && catalogSignature(catalog) !== drawn) { this.fire(sheets); }
		});
	}

	/** The surface and shape a row stands for, for the commands on it. */
	contextOf(node: XlideNode): ShapeRowContext | undefined {
		return this.contexts.get(node);
	}

	/** The row a row made here sits under, and the Sheets folder a sheet's module row sits under. */
	parentOf(node: XlideNode): XlideNode | undefined {
		return this.parents.get(node);
	}

	/**
	 * The Shapes folder that goes first under a module row: a worksheet's,
	 * when the sheet has one or more shapes to show; a Word document's
	 * always, since that is where a shape is added.
	 */
	async moduleFolder(module: XlideNode): Promise<XlideNode | undefined> {
		const host = shapeHostForPath(module.filePath);
		if (!module.moduleName || !host) {
			return undefined;
		}
		if (host === 'word') {
			if (module.documentType !== 'document') {
				return undefined;
			}
			const folder = this.folder(module.filePath, 'module', 'Shapes', { moduleName: module.moduleName });
			this.parents.set(folder, module);
			return folder;
		}
		if (host !== 'excel' || module.sheetName === undefined) {
			return undefined;
		}
		const sheet = sheetOfModule(await this.surfacesIfAny(module.filePath), module.moduleName);
		if (!sheet || sheet.shapes.length === 0) {
			return undefined;
		}
		const folder = this.folder(module.filePath, 'module', 'Shapes', { moduleName: module.moduleName });
		this.parents.set(folder, module);
		return folder;
	}

	/**
	 * The rows that lead a project's: a presentation's Slides folder, or a
	 * workbook's Sheets folder; and the module rows left for the project
	 * itself, since a sheet's module row is drawn under Sheets, named for
	 * its sheet. A workbook whose sheets cannot be listed keeps every module
	 * at the project, as before there was a Sheets folder.
	 */
	async projectRows(
		project: XlideNode,
		modules: readonly XlideNode[],
	): Promise<{ folders: XlideNode[]; modules: XlideNode[] }> {
		if (shapeHostForPath(project.filePath) === 'powerpoint') {
			const slides = this.folder(project.filePath, 'slides', 'Slides');
			this.parents.set(slides, project);
			return { folders: [slides], modules: [...modules] };
		}
		const catalog = await this.catalog(project.filePath);
		if (!catalog) {
			return { folders: [], modules: [...modules] };
		}
		const sheets = this.folder(project.filePath, 'sheets', SHEETS_FOLDER_LABEL);
		sheets.itemCount = catalog.length;
		this.parents.set(sheets, project);
		const byCodeName = new Map<string, WorkbookSheet>();
		for (const sheet of catalog) {
			if (sheet.codeName) { byCodeName.set(sheet.codeName.toLowerCase(), sheet); }
		}
		const rest: XlideNode[] = [];
		for (const module of modules) {
			const sheet = module.kind === 'module' && module.moduleName
				? byCodeName.get(module.moduleName.toLowerCase())
				: undefined;
			if (sheet) {
				this.placeModuleRow(module, sheet, sheets);
			} else {
				rest.push(module);
			}
		}
		return { folders: [sheets], modules: rest };
	}

	/** A sheet's module row: under the Sheets folder, named the way the VBE names it. */
	private placeModuleRow(module: XlideNode, sheet: WorkbookSheet, sheets: XlideNode): void {
		module.sheetName = sheet.name;
		module.label = `${module.moduleName} (${sheet.name})`;
		this.parents.set(module, sheets);
	}

	/** The rows under a row made here. */
	async children(node: XlideNode, modulesOf: () => Promise<readonly XlideNode[]>): Promise<XlideNode[]> {
		if (node.kind === 'shape') {
			const context = this.contexts.get(node);
			return context
				? (node.shape?.shapes ?? []).map((member) => this.shapeRow(node, context.host, context.surface, member, true))
				: [];
		}
		if (node.kind !== 'shapes' && node.kind !== 'surface') { return []; }
		if (node.shapeFolder === 'sheets') {
			return this.sheetRows(node, await modulesOf());
		}
		if (node.shapeFolder === 'bareSheets') {
			return (this.bareSheets.get(node) ?? []).map((sheet) => this.sheetRow(node, sheet, undefined));
		}
		const host = shapeHostForPath(node.filePath);
		if (!host) { return []; }
		this.opened.add(node);
		let surfaces: ShapeSurface[];
		try {
			surfaces = await this.surfaces(node.filePath);
		} catch (err) {
			return [this.infoRow(node, 'Shapes could not be read', err instanceof Error ? err.message : String(err))];
		}
		if (node.kind === 'surface') {
			const surface = surfaces.find((s) => s.surface === node.surface);
			if (host !== 'excel') {
				return this.shapeRowsOf(node, host, surface);
			}
			// A sheet's shapes sit in a Shapes folder under it, as a module's do.
			if (!surface || surface.shapes.length === 0) { return []; }
			const folder = this.folder(node.filePath, 'surface', 'Shapes', { surface: surface.surface });
			this.parents.set(folder, node);
			this.contexts.set(folder, { host, surface: surface.surface });
			return [folder];
		}
		switch (node.shapeFolder) {
			case 'module': {
				if (host === 'excel') {
					const sheet = sheetOfModule(surfaces, node.moduleName ?? '');
					if (sheet) { this.contexts.set(node, { host, surface: sheet.surface }); }
					return this.shapeRowsOf(node, host, sheet);
				}
				// Word: the body always, since it is where a shape is added,
				// and any other story that holds one.
				this.contexts.set(node, { host, surface: surfaces[0]?.surface ?? 'Document' });
				return surfaces
					.filter((s, index) => index === 0 || s.shapes.length > 0)
					.map((s) => this.surfaceRow(node, host, s));
			}
			case 'surface':
				return this.shapeRowsOf(node, host, surfaces.find((s) => s.surface === node.surface));
			case 'slides':
				return surfaces.map((s) => this.surfaceRow(node, host, s));
			default:
				return [];
		}
	}

	/**
	 * The surface a folder or surface row adds a shape to: the row's own, or
	 * for a worksheet module's Shapes folder the sheet the module stands for,
	 * read now when the folder has never been opened.
	 */
	async surfaceOf(node: XlideNode): Promise<ShapeRowContext | undefined> {
		const known = this.contexts.get(node);
		if (known) { return known; }
		const host = shapeHostForPath(node.filePath);
		if (!host || node.kind !== 'shapes' || node.shapeFolder !== 'module') { return undefined; }
		const surfaces = await this.surfaces(node.filePath);
		if (host === 'word') { return { host, surface: surfaces[0]?.surface ?? 'Document' }; }
		const sheet = sheetOfModule(surfaces, node.moduleName ?? '');
		return sheet ? { host, surface: sheet.surface } : undefined;
	}

	/** The tree item for a row made here. */
	treeItem(node: XlideNode): vscode.TreeItem {
		const context = this.contexts.get(node);
		switch (node.kind) {
			case 'shapes': {
				const item = new vscode.TreeItem(node.label, vscode.TreeItemCollapsibleState.Collapsed);
				item.id = `sf::${folderKey(node.filePath, node.shapeFolder ?? '', node.moduleName ?? node.surface)}`;
				item.iconPath = new vscode.ThemeIcon(node.shapeFolder === 'sheets' ? 'table' : 'layers');
				if (node.shapeFolder === 'slides') {
					item.contextValue = 'shapeFolder';
					item.tooltip = 'The slides of this presentation, and the shapes on each.';
				} else if (node.shapeFolder === 'sheets') {
					item.contextValue = 'sheetsFolder';
					item.description = node.itemCount === 1 ? '1 sheet' : `${node.itemCount ?? 0} sheets`;
					item.tooltip = 'The worksheets and chart sheets of this workbook, in tab order. A sheet with a module'
						+ ' is its module, named the way the VBA editor names it; a sheet with shapes lists them under'
						+ ' Shapes; the sheets with neither are in a folder at the end.';
				} else if (node.shapeFolder === 'bareSheets') {
					item.contextValue = 'shapeFolder';
					item.description = node.itemCount === 1 ? '1 sheet' : `${node.itemCount ?? 0} sheets`;
					item.tooltip = 'Sheets with no module in the VBA project and no shapes. Excel gives a sheet its'
						+ ' module once the VBA editor is opened after the sheet was added.';
				} else {
					item.contextValue = 'shapeFolder-add';
					item.tooltip = shapeHostForPath(node.filePath) === 'word'
						? 'The shapes of this document, by story. Word cannot run a macro from a shape.'
						: node.shapeFolder === 'surface'
							? `The shapes on ${node.surface}.`
							: `The shapes on the worksheet whose module is ${node.moduleName}.`;
				}
				return item;
			}
			case 'surface': {
				const sheet = this.sheetOfRow.get(node);
				if (sheet) {
					return this.sheetItem(node, sheet, context);
				}
				const item = new vscode.TreeItem(node.label, vscode.TreeItemCollapsibleState.Collapsed);
				item.id = `su::${projectIdentityKey(node.filePath)}::${node.surface}`;
				const host = context?.host;
				item.iconPath = new vscode.ThemeIcon(host === 'powerpoint' ? 'preview' : host === 'word' ? 'note' : 'table');
				item.description = node.itemCount === 1 ? '1 shape' : `${node.itemCount ?? 0} shapes`;
				item.contextValue = host && takesNewShapes(host, node.surface ?? '') ? 'shapeSurface-add' : 'shapeSurface';
				return item;
			}
			case 'shape':
				return this.shapeItem(node, context);
			default: {
				const item = new vscode.TreeItem(node.label, vscode.TreeItemCollapsibleState.None);
				item.id = `ns::${projectIdentityKey(node.filePath)}::${node.surface ?? node.moduleName ?? ''}::${node.label}`;
				item.iconPath = new vscode.ThemeIcon(node.errorMessage ? 'warning' : 'info');
				item.contextValue = 'noShapes';
				if (node.errorMessage) { item.tooltip = node.errorMessage; }
				return item;
			}
		}
	}

	/** Every surface of a file with its shapes, read once per project until it changes. */
	surfaces(filePath: string): Promise<ShapeSurface[]> {
		const project = projectIdentityKey(filePath);
		const cached = this.listings.get(project);
		if (cached && !this.stale.has(project)) { return Promise.resolve(cached); }
		const failed = this.failures.get(project);
		if (failed !== undefined) { return Promise.reject(new Error(failed)); }
		let load = this.loads.get(project);
		if (!load) {
			const generation = this.generation;
			const started = this.bridge.call<{ surfaces?: ShapeSurface[] }>('listShapes', { path: filePath }).then(
				(result) => {
					const surfaces = Array.isArray(result?.surfaces) ? result.surfaces : [];
					if (this.generation === generation && this.loads.get(project) === started) {
						this.listings.set(project, surfaces);
						this.stale.delete(project);
					}
					return surfaces;
				},
				(err: unknown) => {
					const message = err instanceof Error ? err.message : String(err);
					this.out?.appendLine(`[projectExplorer] The shapes of "${path.basename(filePath)}" could not be read: ${message}`);
					if (this.generation === generation && this.loads.get(project) === started) {
						this.failures.set(project, message);
					}
					throw err;
				},
			);
			load = started;
			const settled = (): void => { if (this.loads.get(project) === started) { this.loads.delete(project); } };
			started.then(settled, settled);
			this.loads.set(project, started);
		}
		return load;
	}

	/**
	 * The workbook's sheets in tab order, read once per project; undefined
	 * for a file that is not a workbook, or whose sheets could not be read,
	 * which is said once in the output channel.
	 */
	catalog(filePath: string): Promise<WorkbookSheet[] | undefined> {
		if (containerHostForPath(filePath) !== 'excel') { return Promise.resolve(undefined); }
		const project = projectIdentityKey(filePath);
		const cached = this.catalogs.get(project);
		if (cached) { return Promise.resolve(cached); }
		if (this.catalogFailures.has(project)) { return Promise.resolve(undefined); }
		let load = this.catalogLoads.get(project);
		if (!load) {
			const generation = this.generation;
			const started: Promise<WorkbookSheet[] | undefined> = this.bridge
				.call<{ sheets?: WorkbookSheet[] }>('listWorkbookSheets', { path: filePath })
				.then((result) => {
					if (!Array.isArray(result?.sheets)) {
						throw new Error('The sheet list did not come back as one.');
					}
					if (this.generation === generation && this.catalogLoads.get(project) === started) {
						this.catalogs.set(project, result.sheets);
					}
					return result.sheets;
				})
				.catch((err: unknown) => {
					// An answer that is not a sheet list fails the same way as
					// no answer: the workbook keeps its flat listing. Said once
					// per reason, not on every save of a file that never reads.
					const message = err instanceof Error ? err.message : String(err);
					if (this.catalogFailureSaid.get(project) !== message) {
						this.catalogFailureSaid.set(project, message);
						this.out?.appendLine(`[projectExplorer] The sheets of "${path.basename(filePath)}" could not be listed: ${message}`);
					}
					if (this.generation === generation && this.catalogLoads.get(project) === started) {
						this.catalogFailures.add(project);
					}
					return undefined;
				});
			load = started;
			const settled = (): void => { if (this.catalogLoads.get(project) === started) { this.catalogLoads.delete(project); } };
			started.then(settled, settled);
			this.catalogLoads.set(project, started);
		}
		return load;
	}

	/** The shapes listing, or none: for a format the shape tools do not read, and for a listing that failed. */
	private async surfacesIfAny(filePath: string): Promise<ShapeSurface[]> {
		if (!shapeHostForPath(filePath)) { return []; }
		try {
			return await this.surfaces(filePath);
		} catch {
			return [];
		}
	}

	/**
	 * The rows of the Sheets folder, in tab order: a sheet's module row, or
	 * the sheet itself when it has shapes and no module; and, last, the
	 * folder of sheets with neither. What the shapes listing cannot say - a
	 * format it does not read, a drawing that would not parse - counts as no
	 * shapes, so the sheet is still listed, in that folder.
	 */
	private async sheetRows(folder: XlideNode, modules: readonly XlideNode[]): Promise<XlideNode[]> {
		const catalog = (await this.catalog(folder.filePath)) ?? [];
		this.sheetsDrawn.set(projectIdentityKey(folder.filePath), catalogSignature(catalog));
		const surfaces = await this.surfacesIfAny(folder.filePath);
		const byName = new Map<string, XlideNode>();
		for (const module of modules) {
			if (module.kind === 'module' && module.moduleName) { byName.set(module.moduleName.toLowerCase(), module); }
		}
		const rows: XlideNode[] = [];
		const bare: WorkbookSheet[] = [];
		for (const sheet of catalog) {
			const module = sheet.codeName ? byName.get(sheet.codeName.toLowerCase()) : undefined;
			if (module) {
				// Named here as well: a Sheets folder drawn again on its own,
				// after a sheet was renamed, shows the new name.
				this.placeModuleRow(module, sheet, folder);
				rows.push(module);
				continue;
			}
			const surface = surfaces.find((s) => s.surface === sheet.name);
			if (surface && surface.shapes.length > 0) {
				rows.push(this.sheetRow(folder, sheet, surface));
			} else {
				bare.push(sheet);
			}
		}
		if (bare.length > 0) {
			const none = this.folder(folder.filePath, 'bareSheets', BARE_SHEETS_FOLDER_LABEL);
			none.itemCount = bare.length;
			this.parents.set(none, folder);
			this.bareSheets.set(none, bare);
			rows.push(none);
		}
		return rows;
	}

	private folder(
		filePath: string,
		kind: ShapeFolderKind,
		label: string,
		owner: { moduleName?: string; surface?: string } = {},
	): XlideNode {
		const key = folderKey(filePath, kind, owner.moduleName ?? owner.surface);
		let node = this.folders.get(key);
		if (!node) {
			node = {
				kind: 'shapes',
				label,
				filePath,
				shapeFolder: kind,
				...(owner.moduleName ? { moduleName: owner.moduleName } : {}),
				...(owner.surface ? { surface: owner.surface } : {}),
			};
			this.folders.set(key, node);
		}
		return node;
	}

	/** A slide's or story's row: one object per surface, so a redraw finds the row VS Code has. */
	private surfaceRow(parent: XlideNode, host: ShapeHost, surface: ShapeSurface): XlideNode {
		const node = this.surfaceNode(parent.filePath, surface.surface);
		node.itemCount = surface.shapes.length;
		this.parents.set(node, parent);
		this.contexts.set(node, { host, surface: surface.surface });
		return node;
	}

	/**
	 * A sheet's own row, for a sheet with no module: with its shapes, under
	 * Sheets; with none, under Sheets With No Modules or Shapes. A worksheet
	 * the shape tools can write takes a new shape from its row.
	 */
	private sheetRow(parent: XlideNode, sheet: WorkbookSheet, surface: ShapeSurface | undefined): XlideNode {
		const node = this.surfaceNode(parent.filePath, sheet.name);
		node.itemCount = surface?.shapes.length ?? 0;
		this.parents.set(node, parent);
		this.sheetOfRow.set(node, sheet);
		const host = shapeHostForPath(parent.filePath);
		if (host === 'excel' && sheet.kind === 'worksheet') {
			this.contexts.set(node, { host, surface: sheet.name });
		} else {
			this.contexts.delete(node);
		}
		return node;
	}

	private surfaceNode(filePath: string, surface: string): XlideNode {
		const key = `${projectIdentityKey(filePath)}::surface::${surface.toLowerCase()}`;
		let node = this.folders.get(key);
		if (!node) {
			node = { kind: 'surface', label: surface, filePath, surface };
			this.folders.set(key, node);
		}
		return node;
	}

	private shapeRowsOf(parent: XlideNode, host: ShapeHost, surface: ShapeSurface | undefined): XlideNode[] {
		if (!surface || surface.shapes.length === 0) {
			return [this.infoRow(parent, 'No shapes')];
		}
		return surface.shapes.map((shape) => this.shapeRow(parent, host, surface.surface, shape, false));
	}

	private shapeRow(parent: XlideNode, host: ShapeHost, surface: string, shape: ShapeInfo, inGroup: boolean): XlideNode {
		const node: XlideNode = {
			kind: 'shape',
			label: shape.name,
			filePath: parent.filePath,
			surface,
			shape,
			shapePath: [...(parent.kind === 'shape' ? parent.shapePath ?? [] : []), shape.name],
		};
		this.parents.set(node, parent);
		this.contexts.set(node, { host, surface, shape, inGroup });
		return node;
	}

	private infoRow(parent: XlideNode, label: string, errorMessage?: string): XlideNode {
		const node: XlideNode = {
			kind: 'noShapes',
			label,
			filePath: parent.filePath,
			...(parent.surface ? { surface: parent.surface } : {}),
			...(parent.moduleName ? { moduleName: parent.moduleName } : {}),
			...(errorMessage ? { errorMessage } : {}),
		};
		this.parents.set(node, parent);
		return node;
	}

	/** The row of a sheet with no module: its name, its kind, whether it is hidden, and what it holds. */
	private sheetItem(node: XlideNode, sheet: WorkbookSheet, context: ShapeRowContext | undefined): vscode.TreeItem {
		const shapes = node.itemCount ?? 0;
		const item = new vscode.TreeItem(
			node.label,
			shapes > 0 ? vscode.TreeItemCollapsibleState.Collapsed : vscode.TreeItemCollapsibleState.None,
		);
		item.id = `su::${projectIdentityKey(node.filePath)}::${node.surface}`;
		item.iconPath = new vscode.ThemeIcon(sheet.kind === 'chartsheet' ? 'graph' : 'table');
		const kind = SHEET_KIND_LABELS[sheet.kind];
		const notes = [
			sheet.kind === 'worksheet' ? '' : kind,
			sheet.state === 'hidden' ? 'hidden' : sheet.state === 'veryHidden' ? 'very hidden' : '',
		].filter(Boolean);
		if (notes.length > 0) { item.description = notes.join(', '); }
		item.contextValue = context ? 'shapeSurface-add' : 'shapeSurface';
		item.tooltip = `${sheet.name}: a ${kind} with no module in the VBA project`
			+ (shapes > 0 ? `, with ${shapes === 1 ? '1 shape' : `${shapes} shapes`}.` : '.');
		return item;
	}

	private shapeItem(node: XlideNode, context: ShapeRowContext | undefined): vscode.TreeItem {
		const shape = node.shape!;
		const members = shape.shapes?.length ?? 0;
		const item = new vscode.TreeItem(
			node.label,
			members > 0 ? vscode.TreeItemCollapsibleState.Collapsed : vscode.TreeItemCollapsibleState.None,
		);
		item.id = `sh::${projectIdentityKey(node.filePath)}::${node.surface}::${(node.shapePath ?? [node.label]).join('/')}`;
		const icon = shape.kind === 'shape' && shape.geometry === 'ellipse' ? 'circle-large-outline' : KIND_ICONS[shape.kind] ?? 'symbol-misc';
		item.iconPath = new vscode.ThemeIcon(icon);
		const parts = [shapeKindLabel(shape.kind)];
		if (shape.macro) { parts.push(`runs ${shape.macro}`); }
		if (shape.hidden) { parts.push('hidden'); }
		item.description = parts.join(', ');
		const host = context?.host;
		const flags = [
			shape.kind !== 'activeX' ? 'edit' : '',
			host && canRunMacro(host, shape.kind) ? 'link' : '',
			shape.macro ? 'macro' : '',
			!context?.inGroup && shape.kind !== 'activeX' ? 'delete' : '',
		].filter(Boolean);
		item.contextValue = ['shape', ...flags].join('-');
		const lines = [`${shape.name} (${shapeKindLabel(shape.kind)}) on ${node.surface}`];
		if (shape.macro) { lines.push(`Runs ${shape.macro} on a click.`); }
		const where = shape.range ?? (shape.width !== undefined
			? `${shape.width} x ${shape.height} pt${shape.left !== undefined ? ` at ${shape.left}, ${shape.top}` : ''}`
			: undefined);
		if (where) { lines.push(where); }
		if (shape.text) { lines.push(`"${shape.text.length > 80 ? `${shape.text.slice(0, 80)}...` : shape.text}"`); }
		if (shape.kind === 'activeX') {
			lines.push('An ActiveX control runs event procedures in its sheet\'s module. XLIDE lists it but does not edit it.');
		}
		item.tooltip = lines.join('\n');
		if (shape.kind !== 'activeX') {
			item.command = { command: 'xlide.editShape', title: 'Edit Shape', arguments: [node] };
		}
		return item;
	}
}

/** The worksheet a module stands for, by its code name. */
function sheetOfModule(surfaces: readonly ShapeSurface[], moduleName: string): ShapeSurface | undefined {
	const wanted = moduleName.toLowerCase();
	return surfaces.find((s) => s.codeName?.toLowerCase() === wanted);
}

/** A sheet list as one comparable value: what a redraw of the Sheets folder would change. */
function catalogSignature(catalog: readonly WorkbookSheet[]): string {
	return catalog.map((sheet) => `${sheet.name}|${sheet.codeName ?? ''}|${sheet.kind}|${sheet.state ?? ''}`).join('\n');
}

/** A folder's key in the row map, apart from the surface rows' keys, which a sheet's own row uses. */
function folderKey(filePath: string, kind: string, owner: string | undefined): string {
	return `${projectIdentityKey(filePath)}::folder::${kind}::${(owner ?? '').toLowerCase()}`;
}
