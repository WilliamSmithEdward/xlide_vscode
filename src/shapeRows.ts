// Shapes in the explorer, where each host keeps them.
//
// An Excel worksheet is a module in the VBA project, so a sheet's shapes are
// a Shapes folder under its module. Word's shapes belong to the document, so
// they are a Shapes folder under ThisDocument, one row per story (the body,
// a header, a footer). A PowerPoint slide is no module at all, so the slides
// are a folder of their own under the project.
//
// Excel gives a worksheet its module only once the workbook's VBA editor has
// been opened after the sheet was added, and a workbook saved before then
// carries sheets with no module - ShapesFixture has one. Those sheets have no
// module row to hang under, so they are a folder at the project level.
//
// The shapes are read from the file when a folder is opened, and the listing
// is kept per project until the file changes or the tree refreshes. After a
// change the old listing stays on screen until the new one lands, so a save
// does not make a folder blink out and back.

import * as vscode from 'vscode';
import * as path from 'path';
import type { ProjectEngine } from './projectEngine';
import type { XlideNode } from './projectExplorer';
import type { ShapeInfo, ShapeKind } from './vba/shapes';
import type { ShapeSurface } from './vba/projectService';
import { canRunMacro, type ShapeHost } from './vba/shapeCapabilities';
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

/**
 * The shape rows of the explorer. The explorer owns the tree and asks this
 * for the rows under its module and project rows and under the rows made
 * here; `fire` redraws a row the way the explorer's own emitter does.
 */
export class ShapeRows {
	private readonly listings = new Map<string, ShapeSurface[]>();
	private readonly stale = new Set<string>();
	private readonly loads = new Map<string, Promise<ShapeSurface[]>>();
	private readonly failures = new Map<string, string>();
	/** Folder rows by key, so each folder is one row for as long as the tree lasts. */
	private readonly folders = new Map<string, XlideNode>();
	/** The folder and surface rows whose shapes were asked for: the ones a change redraws. */
	private readonly opened = new Set<XlideNode>();
	/** Each workbook's sheets with no module, as last drawn, by project. */
	private readonly orphansDrawn = new Map<string, string>();
	private readonly parents = new WeakMap<XlideNode, XlideNode>();
	private readonly contexts = new WeakMap<XlideNode, ShapeRowContext>();
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
		this.folders.clear();
		this.opened.clear();
		this.orphansDrawn.clear();
	}

	/**
	 * A file changed, by a shape edit or a save from outside: its listing is
	 * read again when next asked for, and the shape rows someone opened are
	 * redrawn, which asks. Nothing is read here, since a file that changes on
	 * every keystroke's auto-save should cost the tree nothing while its
	 * shapes are not on screen.
	 */
	refresh(filePath: string): void {
		const project = projectIdentityKey(filePath);
		this.failures.delete(project);
		this.loads.delete(project);
		if (this.listings.has(project)) { this.stale.add(project); }
		for (const node of this.opened) {
			if (projectIdentityKey(node.filePath) === project) { this.fire(node); }
		}
	}

	/** The surface and shape a row stands for, for the commands on it. */
	contextOf(node: XlideNode): ShapeRowContext | undefined {
		return this.contexts.get(node);
	}

	/** The row a shape row, surface row or folder row sits under. */
	parentOf(node: XlideNode): XlideNode | undefined {
		return this.parents.get(node);
	}

	/** The Shapes folder that goes first under a module row, when its module has shapes to show. */
	moduleFolder(module: XlideNode): XlideNode | undefined {
		const host = shapeHostForPath(module.filePath);
		const kind = module.documentType;
		if (!module.moduleName || !((host === 'excel' && kind === 'worksheet') || (host === 'word' && kind === 'document'))) {
			return undefined;
		}
		const folder = this.folder(module.filePath, 'module', module.moduleName, 'Shapes');
		this.parents.set(folder, module);
		return folder;
	}

	/**
	 * The folders that go first under a project row: a presentation's slides,
	 * and a workbook's sheets that have no module. The workbook's are known
	 * only once its shapes are read, so the first drawing starts the read and
	 * the row is drawn again when it lands.
	 */
	projectFolders(project: XlideNode, moduleNames: readonly string[]): XlideNode[] {
		const host = shapeHostForPath(project.filePath);
		if (host === 'powerpoint') {
			const slides = this.folder(project.filePath, 'slides', undefined, 'Slides');
			this.parents.set(slides, project);
			return [slides];
		}
		if (host !== 'excel') { return []; }
		const key = projectIdentityKey(project.filePath);
		const listing = this.listings.get(key);
		if (!listing || this.stale.has(key)) {
			// Read in the background; the row is drawn again only when the
			// sheets it shows have changed.
			void this.surfaces(project.filePath).then(
				(surfaces) => {
					if (sheetNames(orphanSheets(surfaces, moduleNames)) !== (this.orphansDrawn.get(key) ?? '')) {
						this.fire(project);
					}
				},
				() => undefined,
			);
		}
		const orphans = listing ? orphanSheets(listing, moduleNames) : [];
		this.orphansDrawn.set(key, sheetNames(orphans));
		if (orphans.length === 0) { return []; }
		const sheets = this.folder(project.filePath, 'sheets', undefined, 'Sheets With No Module');
		sheets.itemCount = orphans.length;
		this.parents.set(sheets, project);
		return [sheets];
	}

	/** The rows under a row made here. */
	async children(node: XlideNode, moduleNames: () => Promise<readonly string[]>): Promise<XlideNode[]> {
		if (node.kind === 'shape') {
			const context = this.contexts.get(node);
			return context
				? (node.shape?.shapes ?? []).map((member) => this.shapeRow(node, context.host, context.surface, member, true))
				: [];
		}
		if (node.kind !== 'shapes' && node.kind !== 'surface') { return []; }
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
			return this.shapeRowsOf(node, host, surfaces.find((s) => s.surface === node.surface));
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
			case 'slides':
				return surfaces.map((s) => this.surfaceRow(node, host, s));
			case 'sheets':
				return orphanSheets(surfaces, await moduleNames()).map((s) => this.surfaceRow(node, host, s));
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
				item.id = `sf::${folderKey(node.filePath, node.shapeFolder ?? '', node.moduleName)}`;
				item.iconPath = new vscode.ThemeIcon('layers');
				if (node.shapeFolder === 'slides') {
					item.contextValue = 'shapeFolder';
					item.tooltip = 'The slides of this presentation, and the shapes on each.';
				} else if (node.shapeFolder === 'sheets') {
					item.contextValue = 'shapeFolder';
					item.description = node.itemCount === 1 ? '1 sheet' : `${node.itemCount ?? 0} sheets`;
					item.tooltip = 'Worksheets with no module in the VBA project. Excel gives a sheet its module once'
						+ ' the VBA editor is opened after the sheet was added; until then its shapes are listed here.';
				} else {
					item.contextValue = 'shapeFolder-add';
					item.tooltip = shapeHostForPath(node.filePath) === 'word'
						? 'The shapes of this document, by story. Word cannot run a macro from a shape.'
						: `The shapes on the worksheet whose module is ${node.moduleName}.`;
				}
				return item;
			}
			case 'surface': {
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

	private folder(filePath: string, kind: 'module' | 'slides' | 'sheets', moduleName: string | undefined, label: string): XlideNode {
		const key = folderKey(filePath, kind, moduleName);
		let node = this.folders.get(key);
		if (!node) {
			node = { kind: 'shapes', label, filePath, shapeFolder: kind, ...(moduleName ? { moduleName } : {}) };
			this.folders.set(key, node);
		}
		return node;
	}

	/** A slide's or story's row: one object per surface, so a redraw finds the row VS Code has. */
	private surfaceRow(parent: XlideNode, host: ShapeHost, surface: ShapeSurface): XlideNode {
		const key = `${projectIdentityKey(parent.filePath)}::surface::${surface.surface.toLowerCase()}`;
		let node = this.folders.get(key);
		if (!node) {
			node = { kind: 'surface', label: surface.surface, filePath: parent.filePath, surface: surface.surface };
			this.folders.set(key, node);
		}
		node.itemCount = surface.shapes.length;
		this.parents.set(node, parent);
		this.contexts.set(node, { host, surface: surface.surface });
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

/** A set of sheets as one comparable value. */
function sheetNames(sheets: readonly ShapeSurface[]): string {
	return sheets.map((sheet) => sheet.surface).join('\n');
}

/** The worksheets of a listing that no module in the project stands for. */
function orphanSheets(surfaces: readonly ShapeSurface[], moduleNames: readonly string[]): ShapeSurface[] {
	const modules = new Set(moduleNames.map((name) => name.toLowerCase()));
	return surfaces.filter((s) => !s.codeName || !modules.has(s.codeName.toLowerCase()));
}

function folderKey(filePath: string, kind: string, moduleName: string | undefined): string {
	return `${projectIdentityKey(filePath)}::${kind}::${(moduleName ?? '').toLowerCase()}`;
}
