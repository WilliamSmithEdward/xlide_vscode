// The shape editor: one shape's every property in a form, in an editor tab.
//
// The tab opens on what the file holds now, not on the tree's listing,
// which can be older than a save from outside. A Save sends only what
// changed (shapeEditorModel.ts) as one edit through the Office write
// coordination every container write goes through, then refreshes the
// tree's shape rows and closes the tab; Cancel closes it; Delete asks first.
// A failed save keeps the tab open with the reason, so nothing typed is lost.

import * as vscode from 'vscode';
import * as path from 'path';
import type { ProjectEngine } from './projectEngine';
import type { ProjectExplorer } from './projectExplorer';
import {
	shapeEditFromForm,
	shapeEditorModel,
	type ShapeEditorModel,
	type ShapeEditorTarget,
	type ShapeFormValues,
} from './shapeEditorModel';
import type { ShapeMacro } from './vba/projectService';
import { shapeKindLabel } from './shapeRows';
import { runWriteWithHostCoordination } from './officeWriteCoordinator';
import { randomNonce, scriptJson } from './webview/html';
import { webviewHeadHtml } from './webview/page';
import { renderWebviewTemplate } from './webview/templates';
import { WEBVIEW_BODY_CSS, WEBVIEW_PRIMARY_BUTTON_CSS, xlideAccentPaletteCss } from './webview/styles';
import { errorMessage } from './util/errors';
import { isProjectLockedError, projectIdentityKey, reportProjectLocked } from './xlideFileSystem';
import { formatChangeSummary, withWriteAudit } from './xlideWriteAudit';
import type { ShapeEdit, ShapeInfo } from './vba/shapes';

export interface ShapeEditorDeps {
	bridge: ProjectEngine;
	explorer: ProjectExplorer;
	out: vscode.OutputChannel;
	/** Opens the Sub a shape runs, as the tree's Go to Macro does. */
	goToMacro: (filePath: string, macro: string) => Promise<void>;
}

/** Open editor tabs, so asking for one already open brings it forward. */
const openEditors = new Map<string, vscode.WebviewPanel>();
/** Repeated clicks share the reads that happen before a tab is registered. */
const openingEditors = new Map<string, Promise<void>>();

function editorKey(filePath: string, surface: string, shapePath: readonly string[] | undefined): string {
	return JSON.stringify([projectIdentityKey(filePath), surface.toLowerCase(), shapePath?.map(name => name.toLowerCase()) ?? null]);
}

/** The Subs a shape in this file can run; none when they cannot be read. */
async function shapeMacros(deps: ShapeEditorDeps, filePath: string): Promise<ShapeMacro[]> {
	try {
		return (await deps.bridge.call<{ macros: ShapeMacro[] }>('shapeMacros', { path: filePath })).macros;
	} catch (err) {
		deps.out.appendLine(`[shapeEditor] The Subs of "${path.basename(filePath)}" could not be listed: ${errorMessage(err)}`);
		return [];
	}
}

/** Make one shape edit, coordinated with the Office application holding the file, and audit it. */
export async function writeShapeEdit(
	deps: Pick<ShapeEditorDeps, 'bridge' | 'explorer'>,
	filePath: string,
	surface: string,
	edit: ShapeEdit,
	command: string,
): Promise<string> {
	const title = edit.action === 'add' ? 'Add shape' : edit.action === 'delete' ? 'Delete shape' : 'Change shape';
	try {
		const { result } = await withWriteAudit({
			command,
			operation: 'edit-shape',
			projectPath: filePath,
			failedSummary: `${title}: 0 changed, 1 failed`,
		}, async () => {
			const result = await runWriteWithHostCoordination(filePath, () => deps.bridge.call<{ ok: true; name: string }>(
				'editShape',
				{ path: filePath, surface, ...edit },
			));
			const shape = `${surface}!${result.name}`;
			return {
				result,
				summary: formatChangeSummary({
					operation: title,
					...(edit.action === 'delete' ? { removed: [shape] } : { changed: [shape] }),
				}),
			};
		});
		return result.name;
	} finally {
		// Whatever happened, the rows show the file as it is now.
		deps.explorer.refreshShapes(filePath, { shapesChanged: true });
	}
}

/**
 * Open the editor on a shape, or on a new one when `target.shape` is
 * missing. The shape is read again from the file first, so the tab shows
 * what is there now.
 */
export async function openShapeEditor(
	deps: ShapeEditorDeps,
	context: vscode.ExtensionContext,
	filePath: string,
	target: Omit<ShapeEditorTarget, 'fileName'> & { shapePath?: string[] },
): Promise<void> {
	const key = editorKey(filePath, target.surface, target.shape ? target.shapePath ?? [target.shape.name] : undefined);
	const open = openEditors.get(key);
	if (open) {
		open.reveal();
		return;
	}
	const pending = openingEditors.get(key);
	if (pending) {
		await pending;
		openEditors.get(key)?.reveal();
		return;
	}
	const started = createShapeEditor(deps, context, filePath, target, key);
	openingEditors.set(key, started);
	try {
		await started;
	} finally {
		if (openingEditors.get(key) === started) { openingEditors.delete(key); }
	}
}

async function createShapeEditor(
	deps: ShapeEditorDeps,
	context: vscode.ExtensionContext,
	filePath: string,
	target: Omit<ShapeEditorTarget, 'fileName'> & { shapePath?: string[] },
	key: string,
): Promise<void> {
	let fresh: ShapeEditorTarget = { ...target, fileName: path.basename(filePath) };
	if (target.shape) {
		const listing = await deps.bridge.call<{ surfaces: Array<{ surface: string; shapes: ShapeInfo[] }> }>(
			'listShapes', { path: filePath, surface: target.surface },
		);
		const shapes = listing.surfaces[0]?.shapes ?? [];
		const found = findByPath(shapes, target.shapePath ?? [target.shape.name]);
		if (!found) {
			void vscode.window.showWarningMessage(`"${target.shape.name}" is no longer on ${target.surface} in ${path.basename(filePath)}.`);
			deps.explorer.refreshShapes(filePath);
			return;
		}
		fresh = { ...fresh, shape: found, stackCount: shapes.filter((s) => s.zOrder !== undefined).length };
	}
	const model = shapeEditorModel(fresh, await shapeMacros(deps, filePath), fresh.shape ? shapeKindLabel(fresh.shape.kind) : undefined);
	const panel = vscode.window.createWebviewPanel(
		'xlide.shapeEditor',
		model.mode === 'add' ? `New Shape - ${model.surface}` : `${model.shape!.name} - Shape`,
		vscode.ViewColumn.Active,
		{
			enableScripts: true,
			// The form holds what the user typed; hidden, it must not reset.
			retainContextWhenHidden: true,
			localResourceRoots: [context.extensionUri],
		},
	);
	openEditors.set(key, panel);
	panel.webview.html = renderShapeEditorHtml(model);
	let writePending = false, disposed = false;
	const messages = panel.webview.onDidReceiveMessage(async (message: { type?: string; values?: ShapeFormValues; macro?: string }) => {
		if (disposed) { return; }
		const writing = message.type === 'save' || message.type === 'delete';
		if (writing && writePending) { return; }
		if (writing) { writePending = true; }
		try {
			switch (message.type) {
				case 'cancel':
					panel.dispose();
					return;
				case 'goToMacro':
					if (message.macro) { await deps.goToMacro(filePath, message.macro); }
					return;
				case 'delete':
					await deleteFromEditor(deps, panel, filePath, model, () => !disposed);
					return;
				case 'save':
					if (message.values) { await saveFromEditor(deps, panel, filePath, model, message.values); }
					return;
				default:
			}
		} catch (err) {
			if (isProjectLockedError(errorMessage(err))) {
				reportProjectLocked(filePath, 'write', err);
			}
			await panel.webview.postMessage({ type: 'error', error: errorMessage(err) });
		} finally {
			if (writing) { writePending = false; }
		}
	});
	panel.onDidDispose(() => {
		disposed = true;
		messages.dispose();
		if (openEditors.get(key) === panel) { openEditors.delete(key); }
	});
}

async function saveFromEditor(
	deps: ShapeEditorDeps,
	panel: vscode.WebviewPanel,
	filePath: string,
	model: ShapeEditorModel,
	values: ShapeFormValues,
): Promise<void> {
	const { edit, surface, errors } = shapeEditFromForm(model, values);
	if (Object.keys(errors).length > 0) {
		await panel.webview.postMessage({ type: 'invalid', errors });
		return;
	}
	if (!edit) {
		// Nothing changed: Save is as good as Cancel.
		panel.dispose();
		return;
	}
	const name = await writeShapeEdit(deps, filePath, surface, edit, 'xlide.editShape');
	panel.dispose();
	vscode.window.setStatusBarMessage(
		`XLIDE: ${edit.action === 'add' ? 'Added' : 'Saved'} shape "${name}" on ${surface}.`,
		6000,
	);
}

async function deleteFromEditor(deps: ShapeEditorDeps, panel: vscode.WebviewPanel, filePath: string, model: ShapeEditorModel, isOpen: () => boolean): Promise<void> {
	const name = model.shape?.name;
	if (!name) { return; }
	const choice = await vscode.window.showWarningMessage(
		`Delete shape "${name}" from ${model.surface} in "${path.basename(filePath)}"?`,
		{ modal: true },
		'Delete',
	);
	if (!isOpen()) { return; }
	if (choice !== 'Delete') {
		await panel.webview.postMessage({ type: 'idle' });
		return;
	}
	await writeShapeEdit(deps, filePath, model.surface, { action: 'delete', name }, 'xlide.deleteShape');
	panel.dispose();
	vscode.window.setStatusBarMessage(`XLIDE: Deleted shape "${name}" from ${model.surface}.`, 6000);
}

/** A shape found by its name after the names of the groups it is in. */
function findByPath(shapes: readonly ShapeInfo[], names: readonly string[]): ShapeInfo | undefined {
	let level: readonly ShapeInfo[] = shapes;
	let found: ShapeInfo | undefined;
	for (const name of names) {
		found = level.find((s) => s.name.toLowerCase() === name.toLowerCase());
		if (!found) { return undefined; }
		level = found.shapes ?? [];
	}
	return found;
}

/** The editor's page for a model. */
export function renderShapeEditorHtml(model: ShapeEditorModel): string {
	const nonce = randomNonce();
	return renderWebviewTemplate('assets/webview/shapeEditor.html', {
		head: webviewHeadHtml(nonce, model.mode === 'add' ? 'New Shape' : `${model.shape!.name} - Shape`),
		nonce,
		css: renderWebviewTemplate('assets/webview/shapeEditor.css', {
			accentCss: xlideAccentPaletteCss(),
			bodyCss: WEBVIEW_BODY_CSS,
			buttonCss: WEBVIEW_PRIMARY_BUTTON_CSS,
		}),
		js: renderWebviewTemplate('assets/webview/shapeEditor.js', {
			modelJson: scriptJson(model),
		}),
	});
}
