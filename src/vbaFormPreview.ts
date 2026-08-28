// The form designer panel: the parsed designer model rendered in a webview,
// with the canvas gestures the VBE's own designer has - drag to move, resize
// handles, a toolbox that places new controls, grid and neighbor snapping,
// arrow-key nudges, Delete.
//
// The webview owns only the GESTURE. Every mutation posts back here, applies
// through the engine's designer ops (the same primitives the markup diff
// uses, on the same authoring path live Excel verified), and the panel
// re-renders from the workbook - so what the canvas shows is always what the
// bytes say, never an optimistic guess.

import * as vscode from 'vscode';
import * as path from 'path';
import type { WorkbookEngine } from './workbookEngine';
import { decodeModuleUri, encodeFormMarkupUri, encodeModuleUri, XLIDE_SCHEME, XlideFileSystemProvider } from './xlideFileSystem';
import { runWriteWithExcelCoordination } from './excelWorkbookCoordinator';
import { recordXlideWriteAudit } from './xlideWriteAudit';
import { errorMessage } from './util/errors';

const panels = new Map<string, vscode.WebviewPanel>();

function panelKey(xlsmPath: string, moduleName: string): string {
	return `${xlsmPath.toLowerCase()}::${moduleName.toLowerCase()}`;
}

type DesignerMessage =
	| { type: 'geometry'; name: string; left?: number; top?: number; width?: number; height?: number }
	| { type: 'add'; container: string; controlKind: string; left: number; top: number }
	| { type: 'remove'; name: string }
	| { type: 'reparent'; name: string; container: string; left: number; top: number }
	| { type: 'setProp'; name: string; prop: string; value: string }
	| { type: 'openHandler'; name: string; event: string }
	| { type: 'formResize'; width: number; height: number };

export function registerFormPreview(
	context: vscode.ExtensionContext,
	bridge: Pick<WorkbookEngine, 'call'>,
	fsProvider?: Pick<XlideFileSystemProvider, 'notifyFileChanged'>,
): void {
	const render = async (
		panel: vscode.WebviewPanel,
		xlsmPath: string,
		moduleName: string,
		selected?: string,
	): Promise<void> => {
		try {
			const { html } = await bridge.call<{ html: string }>(
				'readFormPreview',
				{ path: xlsmPath, module: moduleName, selected },
			);
			panel.webview.html = html;
		} catch (err) {
			panel.webview.html = `<!DOCTYPE html><html><body style="font-family:sans-serif;padding:16px">`
				+ `<p>XLIDE could not render this form.</p><pre>${errorMessage(err)
					.replace(/&/g, '&amp;').replace(/</g, '&lt;')}</pre></body></html>`;
		}
	};

	// The VBE's double-click: the control's default event handler in the
	// code face - navigate to it, or append the stub the VBE would write and
	// land the cursor inside it. Not a workbook write: the stub is a document
	// edit the user saves like any typed code.
	const openEventHandler = async (
		xlsmPath: string,
		moduleName: string,
		controlName: string,
		eventName: string,
	): Promise<void> => {
		try {
			const handler = `${controlName === '' ? 'UserForm' : controlName}_${eventName}`;
			const escaped = handler.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
			const signature = new RegExp(`^[ \\t]*(?:Private\\s+|Public\\s+|Friend\\s+)?Sub\\s+${escaped}\\s*\\(`, 'im');
			const uri = encodeModuleUri(xlsmPath, moduleName);
			const doc = await vscode.workspace.openTextDocument(uri);
			if (!signature.test(doc.getText())) {
				const eol = doc.eol === vscode.EndOfLine.CRLF ? '\r\n' : '\n';
				const text = doc.getText();
				const lead = text.length === 0 ? '' : text.endsWith(eol) ? eol : eol + eol;
				const edit = new vscode.WorkspaceEdit();
				edit.insert(uri, doc.positionAt(text.length), `${lead}Private Sub ${handler}()${eol}${eol}End Sub${eol}`);
				await vscode.workspace.applyEdit(edit);
			}
			const match = signature.exec(doc.getText());
			const editor = await vscode.window.showTextDocument(doc, {
				viewColumn: vscode.ViewColumn.One,
				preserveFocus: false,
			});
			if (match) {
				const line = Math.min(doc.positionAt(match.index).line + 1, doc.lineCount - 1);
				const at = new vscode.Position(line, 0);
				editor.selection = new vscode.Selection(at, at);
				editor.revealRange(new vscode.Range(at, at), vscode.TextEditorRevealType.InCenterIfOutsideViewport);
			}
		} catch (err) {
			void vscode.window.showErrorMessage(`XLIDE: ${errorMessage(err)}`);
		}
	};

	const applyGesture = async (
		panel: vscode.WebviewPanel,
		xlsmPath: string,
		moduleName: string,
		message: DesignerMessage,
	): Promise<void> => {
		const op = message.type === 'geometry'
			? { kind: 'geometry' as const, name: message.name, left: message.left, top: message.top, width: message.width, height: message.height }
			: message.type === 'add'
				? { kind: 'add' as const, container: message.container, controlKind: message.controlKind, left: message.left, top: message.top }
				: message.type === 'reparent'
					? { kind: 'reparent' as const, name: message.name, container: message.container, left: message.left, top: message.top }
					: message.type === 'setProp'
						? { kind: 'setProp' as const, name: message.name, prop: message.prop, value: message.value }
						: message.type === 'formResize'
						? { kind: 'formSize' as const, width: message.width, height: message.height }
						: { kind: 'remove' as const, name: message.name };
		try {
			const result = await runWriteWithExcelCoordination(xlsmPath, () =>
				bridge.call<{ ok: boolean; signatureDropped: boolean; newName?: string }>(
					'formDesignerOp',
					{ path: xlsmPath, module: moduleName, op },
				));
			recordXlideWriteAudit({
				timestamp: new Date().toISOString(),
				command: 'xlide.previewForm',
				operation: `designer-${message.type}`,
				outcome: 'succeeded',
				workbookPath: xlsmPath,
				moduleName,
				summary: message.type === 'add'
					? `Designer: added ${op.kind === 'add' ? op.controlKind : ''} ${result.newName ?? ''}`
					: message.type === 'reparent'
						? `Designer: moved ${message.name} into ${message.container || 'the form'}`
						: message.type === 'setProp'
							? `Designer: set ${message.prop} of ${message.name || 'the form'}`
							: message.type === 'formResize'
							? `Designer: form resized to ${message.width}x${message.height}pt`
							: `Designer: ${message.type} ${'name' in message ? message.name : ''}`,
			});
			// The markup document is another face of the same form; a canvas
			// gesture must reach an open one the way an agent edit reaches code.
			fsProvider?.notifyFileChanged(encodeFormMarkupUri(xlsmPath, moduleName));
			const keepSelected = message.type === 'remove'
				? undefined
				: message.type === 'formResize'
					? ''
					: message.type === 'setProp'
						? (result.newName ?? message.name)
						: message.type === 'add' ? result.newName : message.name;
			await render(panel, xlsmPath, moduleName, keepSelected);
		} catch (err) {
			void vscode.window.showErrorMessage(`XLIDE: ${errorMessage(err)}`);
			await render(panel, xlsmPath, moduleName);
		}
	};

	context.subscriptions.push(
		vscode.commands.registerCommand('xlide.previewForm', async (node?: { kind?: string; filePath?: string; moduleName?: string; moduleType?: string }) => {
			let xlsmPath = node?.filePath;
			let moduleName = node?.moduleName;
			if ((!xlsmPath || !moduleName) && vscode.window.activeTextEditor) {
				// Invoked from the palette with a form's document focused.
				const uri = vscode.window.activeTextEditor.document.uri;
				if (uri.scheme === XLIDE_SCHEME) {
					const decoded = decodeModuleUri(uri);
					xlsmPath = decoded.xlsmPath;
					moduleName = decoded.moduleName;
				}
			}
			if (!xlsmPath || !moduleName) {
				void vscode.window.showInformationMessage(
					'XLIDE: Preview Form needs a UserForm - pick one in the explorer or focus its document.',
				);
				return;
			}
			// The vbide's designer shows the document under the form; here the
			// markup opens in an editor group below the canvas, unless some
			// group already shows it.
			const showMarkupBelow = async (panel: vscode.WebviewPanel): Promise<void> => {
				try {
					const markupUri = encodeFormMarkupUri(xlsmPath!, moduleName!);
					const shown = vscode.window.visibleTextEditors
						.some((e) => e.document.uri.toString() === markupUri.toString());
					if (shown) { return; }
					panel.reveal(undefined, false);
					await vscode.commands.executeCommand('workbench.action.newGroupBelow');
					await vscode.window.showTextDocument(markupUri, { preserveFocus: false });
				} catch {
					// The split is a nicety; the designer stands alone.
				}
			};
			const key = panelKey(xlsmPath, moduleName);
			const existing = panels.get(key);
			if (existing) {
				existing.reveal(vscode.ViewColumn.Beside, true);
				await render(existing, xlsmPath, moduleName);
				await showMarkupBelow(existing);
				return;
			}
			const panel = vscode.window.createWebviewPanel(
				'xlideFormPreview',
				`${moduleName} (${path.basename(xlsmPath)})`,
				{ viewColumn: vscode.ViewColumn.Beside, preserveFocus: true },
				{ enableScripts: true, retainContextWhenHidden: true },
			);
			panels.set(key, panel);
			panel.onDidDispose(() => panels.delete(key), undefined, context.subscriptions);
			panel.webview.onDidReceiveMessage(
				(message: DesignerMessage) => void (message.type === 'openHandler'
					? openEventHandler(xlsmPath!, moduleName!, message.name, message.event)
					: applyGesture(panel, xlsmPath!, moduleName!, message)),
				undefined,
				context.subscriptions,
			);
			await render(panel, xlsmPath, moduleName);
			await showMarkupBelow(panel);
		}),

		// A saved markup document refreshes its form's open canvas, so the two
		// faces track each other in both directions.
		vscode.workspace.onDidSaveTextDocument((doc) => {
			if (doc.uri.scheme !== XLIDE_SCHEME) { return; }
			try {
				const { xlsmPath, moduleName, face } = decodeModuleUri(doc.uri);
				if (face !== 'form') { return; }
				const panel = panels.get(panelKey(xlsmPath, moduleName));
				if (panel) { void render(panel, xlsmPath, moduleName); }
			} catch {
				// Not a module URI this preview knows; nothing to refresh.
			}
		}),
	);
}
