import * as vscode from 'vscode';
import * as path from 'path';
import type { XlideNode } from '../projectExplorer';
import { registerXlideCommand } from '../xlideCommandRegistration';
import { openShapeEditor, writeShapeEdit, type ShapeEditorDeps } from '../shapeEditor';
import type { ShapeMacro } from '../vba/projectService';
import { takesNewShapes } from '../shapeRows';
import { errorMessage } from '../util/errors';
import { isProjectLockedError, reportProjectLocked } from '../xlideFileSystem';
import type { CommandDeps } from './shared';

/**
 * The shape rows' commands: open a shape in the editor or add one, delete
 * one, and link, unlink or go to the Sub a click on it runs.
 */
export function registerShapeCommands(deps: CommandDeps): vscode.Disposable[] {
	const { bridge, explorer, context, out } = deps;
	const editor: ShapeEditorDeps = {
		bridge,
		explorer,
		out,
		goToMacro: (filePath, macro) => goToMacro(deps, filePath, macro),
	};

	/** Run a shape command, saying why it failed the way module writes do. */
	const guarded = (what: string, run: (node: XlideNode) => Promise<void>) => async (node: XlideNode): Promise<void> => {
		if (!node) { return; }
		try {
			await run(node);
		} catch (err) {
			if (isProjectLockedError(errorMessage(err))) {
				reportProjectLocked(node.filePath, 'write', err);
			} else {
				void vscode.window.showErrorMessage(`XLIDE: ${what} failed: ${errorMessage(err)}`);
			}
		}
	};

	return [
		registerXlideCommand('xlide.editShape', guarded('Opening the shape', async (node) => {
			const row = explorer.shapeContextOf(node);
			if (!row?.shape) { return; }
			await openShapeEditor(editor, context, node.filePath, {
				host: row.host,
				surface: row.surface,
				shape: row.shape,
				inGroup: row.inGroup === true,
				...(node.shapePath ? { shapePath: node.shapePath } : {}),
			});
		})),

		registerXlideCommand('xlide.addShape', guarded('Adding a shape', async (node) => {
			const where = await explorer.shapeSurfaceOf(node);
			if (!where) {
				void vscode.window.showWarningMessage(
					`XLIDE: No worksheet in "${path.basename(node.filePath)}" has ${node.moduleName ?? 'this module'} as its module.`,
				);
				return;
			}
			// Word's Shapes folder adds to any story that takes a shape, the
			// body first; a story's own row adds to that story.
			let surfaces: string[] | undefined;
			if (where.host === 'word' && node.kind === 'shapes') {
				const listing = await bridge.call<{ surfaces: Array<{ surface: string }> }>('listShapes', { path: node.filePath });
				surfaces = listing.surfaces.map((s) => s.surface).filter((surface) => takesNewShapes('word', surface));
			}
			await openShapeEditor(editor, context, node.filePath, {
				host: where.host,
				surface: where.surface,
				...(surfaces ? { surfaces } : {}),
			});
		})),

		registerXlideCommand('xlide.deleteShape', guarded('Deleting the shape', async (node) => {
			const row = explorer.shapeContextOf(node);
			if (!row?.shape) { return; }
			const choice = await vscode.window.showWarningMessage(
				`Delete shape "${row.shape.name}" from ${row.surface} in "${path.basename(node.filePath)}"?`,
				{ modal: true },
				'Delete',
			);
			if (choice !== 'Delete') { return; }
			await writeShapeEdit(editor, node.filePath, row.surface, { action: 'delete', name: row.shape.name }, 'xlide.deleteShape');
			vscode.window.setStatusBarMessage(`XLIDE: Deleted shape "${row.shape.name}" from ${row.surface}.`, 6000);
		})),

		registerXlideCommand('xlide.linkShapeMacro', guarded('Linking the shape to a Sub', async (node) => {
			const row = explorer.shapeContextOf(node);
			if (!row?.shape) { return; }
			const { macros } = await bridge.call<{ macros: ShapeMacro[] }>('shapeMacros', { path: node.filePath });
			if (macros.length === 0) {
				void vscode.window.showInformationMessage(
					'XLIDE: The project has no Public Sub without required parameters for a shape to run. Write one first.',
				);
				return;
			}
			const current = row.shape.macro?.toLowerCase();
			const pick = await vscode.window.showQuickPick(
				macros.map((m) => ({
					label: m.macro,
					description: m.macro === m.proc ? m.module : undefined,
					detail: m.macro.toLowerCase() === current ? 'Runs now' : undefined,
					macro: m.macro,
				})),
				{ title: `Link "${row.shape.name}" to a Sub`, placeHolder: 'The Sub a click on the shape runs' },
			);
			if (!pick) { return; }
			await writeShapeEdit(editor, node.filePath, row.surface, { action: 'update', name: row.shape.name, macro: pick.macro }, 'xlide.linkShapeMacro');
			vscode.window.setStatusBarMessage(`XLIDE: "${row.shape.name}" runs ${pick.macro}.`, 6000);
		})),

		registerXlideCommand('xlide.unlinkShapeMacro', guarded('Unlinking the shape', async (node) => {
			const row = explorer.shapeContextOf(node);
			if (!row?.shape?.macro) { return; }
			await writeShapeEdit(editor, node.filePath, row.surface, { action: 'update', name: row.shape.name, macro: '' }, 'xlide.unlinkShapeMacro');
			vscode.window.setStatusBarMessage(`XLIDE: "${row.shape.name}" runs no macro now.`, 6000);
		})),

		registerXlideCommand('xlide.goToShapeMacro', guarded('Opening the Sub', async (node) => {
			const macro = explorer.shapeContextOf(node)?.shape?.macro;
			if (macro) { await goToMacro(deps, node.filePath, macro); }
		})),
	];
}

/**
 * Open the Sub a shape runs at its declaration. A macro named alone lives
 * in a standard module, found the way the host finds it; one named
 * Module.Proc says where it is. A workbook prefix Excel may keep on it
 * ([0]!, 'Book.xlsm'!) names this file.
 */
async function goToMacro(deps: CommandDeps, filePath: string, macro: string): Promise<void> {
	const bare = macro.trim().replace(/^(?:\[0\]!|'[^']*'!|[^'!\s]+!)/, '');
	const dot = bare.lastIndexOf('.');
	const proc = dot >= 0 ? bare.slice(dot + 1) : bare;
	const qualifier = dot >= 0 ? bare.slice(0, dot) : undefined;
	const modules = await deps.bridge.call<Array<{ name: string }>>('listModules', { path: filePath });
	let moduleName = qualifier ? modules.find((m) => m.name.toLowerCase() === qualifier.toLowerCase())?.name : undefined;
	if (!qualifier) {
		const { macros } = await deps.bridge.call<{ macros: ShapeMacro[] }>('shapeMacros', { path: filePath });
		moduleName = macros.find((m) => m.proc.toLowerCase() === proc.toLowerCase())?.module;
	}
	if (!moduleName) {
		void vscode.window.showWarningMessage(`XLIDE: ${macro} is not a Public Sub in "${path.basename(filePath)}".`);
		return;
	}
	const subs = await deps.bridge.call<Array<{ name: string; kind: string; line: number }>>('listSubs', { path: filePath, module: moduleName });
	const sub = subs.find((s) => s.name.toLowerCase() === proc.toLowerCase());
	await vscode.commands.executeCommand('xlide.openModule', {
		kind: 'sub',
		label: `${sub?.kind ?? 'Sub'} ${sub?.name ?? proc}`,
		filePath,
		moduleName,
		...(sub ? { line: sub.line } : {}),
	} satisfies XlideNode);
}
