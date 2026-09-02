// The VB6 form designer: the MSForms designer's canvas over a form's own
// file. The document IS the .frm (or .ctl, .pag), so there is no scratch
// copy and no apply. The canvas renders from the document's text; a gesture
// asks the engine for the header rewritten and lands as one edit at the top
// of the document, undoable like any edit; a save is the file's own save.
// The pane below the canvas shows the header as written, and an edit there
// replaces the header and leaves the code alone.

import * as vscode from 'vscode';
import type { ProjectEngine } from './projectEngine';
import { moduleLocationOfDocument } from './vbaDocumentLocation';
import { vb6FormHandlerPrefix, vb6HeaderEndOf } from './vba/vb6/frmDesignerOps';
import { errorMessage } from './util/errors';

export const VB6_FORM_DESIGNER_VIEW_TYPE = 'xlideVb6FormDesigner';

type DesignerMessage =
	| { type: 'geometry'; name: string; left?: number; top?: number; width?: number; height?: number }
	| { type: 'geometryBatch'; anchor?: string; items: { name: string; left?: number; top?: number; width?: number; height?: number }[] }
	| { type: 'add'; container: string; controlKind: string; left: number; top: number }
	| { type: 'remove'; name: string }
	| { type: 'reparent'; name: string; container: string; left: number; top: number }
	| { type: 'setProp'; name: string; prop: string; value: string }
	| { type: 'openHandler'; name: string; event: string }
	| { type: 'markupEdit'; text: string }
	| { type: 'docUndo' }
	| { type: 'docRedo' }
	| { type: 'docSave' }
	| { type: 'formResize'; width: number; height: number }
	| { type: 'paste'; names: string[] }
	| { type: 'removeMany'; names: string[] }
	| { type: 'zOrder'; name: string; toFront: boolean }
	| { type: 'tabOrder'; container: string; names: string[] };

interface OpResult {
	text: string;
	newName?: string;
	newNames?: string[];
}

export function registerVb6FormDesigner(
	context: vscode.ExtensionContext,
	bridge: Pick<ProjectEngine, 'call'>,
): void {
	class Vb6FormDesignerProvider implements vscode.CustomTextEditorProvider {
		async resolveCustomTextEditor(
			document: vscode.TextDocument,
			panel: vscode.WebviewPanel,
		): Promise<void> {
			const modulePath = document.uri.fsPath;
			const vbpPath = moduleLocationOfDocument(document)?.projectPath;
			panel.webview.options = { enableScripts: true };

			let disposed = false;
			/** Text this provider just placed in the document (a gesture echo). */
			let suppressEcho: string | null = null;
			/** The document text the webview last rendered, to skip repeats. */
			let lastRenderedText: string | null = null;
			/** All work runs through one lane: edits never interleave. */
			let queue: Promise<void> = Promise.resolve();

			const enqueue = (work: () => Promise<void>): void => {
				queue = queue.then(work).catch((err) => {
					if (!disposed) {
						void vscode.window.showErrorMessage(`XLIDE: ${errorMessage(err)}`);
					}
				});
			};

			const render = async (selected?: string): Promise<void> => {
				if (disposed) { return; }
				const text = document.getText();
				try {
					const { html } = await bridge.call<{ html: string }>('readVb6FormPreview', {
						path: modulePath,
						text,
						selected,
						vbpPath,
					});
					panel.webview.html = html;
					lastRenderedText = text;
				} catch (err) {
					lastRenderedText = null;
					if (panel.webview.html) {
						// The canvas keeps the last good form; the strip names the problem.
						void panel.webview.postMessage({ type: 'markupError', message: errorMessage(err) });
						return;
					}
					panel.webview.html = `<!DOCTYPE html><html><body style="font-family:sans-serif;padding:16px">`
						+ `<p>XLIDE could not render this form.</p><pre>${errorMessage(err)
							.replace(/&/g, '&amp;').replace(/</g, '&lt;')}</pre></body></html>`;
				}
			};

			/**
			 * Replaces the document's text as one edit, touching only the span
			 * that changed - the header, for a gesture - so a text editor open
			 * on the same file keeps its place in the code.
			 */
			const setDocumentText = async (text: string, suppress: boolean): Promise<void> => {
				const current = document.getText();
				if (current === text) { return; }
				let start = 0;
				while (start < current.length && start < text.length && current[start] === text[start]) { start += 1; }
				let endOld = current.length;
				let endNew = text.length;
				while (endOld > start && endNew > start && current[endOld - 1] === text[endNew - 1]) {
					endOld -= 1;
					endNew -= 1;
				}
				suppressEcho = suppress ? text : null;
				const edit = new vscode.WorkspaceEdit();
				edit.replace(
					document.uri,
					new vscode.Range(document.positionAt(start), document.positionAt(endOld)),
					text.slice(start, endNew),
				);
				const applied = await vscode.workspace.applyEdit(edit);
				if (!applied) { suppressEcho = null; }
			};

			const applyOp = async (op: Record<string, unknown>, selectAfter: (result: OpResult) => string | undefined): Promise<void> => {
				try {
					const result = await bridge.call<OpResult>('vb6FormDesignerOp', {
						path: modulePath,
						text: document.getText(),
						op,
					});
					await setDocumentText(result.text, true);
					await render(selectAfter(result));
				} catch (err) {
					void vscode.window.showErrorMessage(`XLIDE: ${errorMessage(err)}`);
					await render();
				}
			};

			const applyGesture = (message: DesignerMessage): Promise<void> => {
				switch (message.type) {
					case 'geometry':
						return applyOp({ kind: 'geometry', name: message.name, left: message.left, top: message.top, width: message.width, height: message.height }, () => message.name);
					case 'geometryBatch':
						return applyOp({ kind: 'geometryBatch', items: message.items }, () => message.anchor);
					case 'add':
						return applyOp({ kind: 'add', container: message.container, controlKind: message.controlKind, left: message.left, top: message.top }, (r) => r.newName);
					case 'remove':
						return applyOp({ kind: 'remove', name: message.name }, () => undefined);
					case 'reparent':
						return applyOp({ kind: 'reparent', name: message.name, container: message.container, left: message.left, top: message.top }, () => message.name);
					case 'setProp':
						return applyOp({ kind: 'setProp', name: message.name, prop: message.prop, value: message.value }, (r) => r.newName ?? message.name);
					case 'formResize':
						return applyOp({ kind: 'formSize', width: message.width, height: message.height }, () => '');
					case 'zOrder':
						return applyOp({ kind: 'zOrder', name: message.name, toFront: message.toFront }, () => message.name);
					case 'tabOrder':
						return applyOp({ kind: 'tabOrder', container: message.container, names: message.names }, () => undefined);
					case 'paste':
						return applyOp({ kind: 'duplicate', names: message.names }, (r) => r.newNames?.[0]);
					case 'removeMany':
						return applyOp({ kind: 'removeMany', names: message.names }, () => undefined);
					default:
						return Promise.resolve();
				}
			};

			/** The pane's text is the header; the code below it stays. */
			const applyMarkupEdit = async (headerText: string): Promise<void> => {
				const text = document.getText();
				const eol = document.eol === vscode.EndOfLine.CRLF ? '\r\n' : '\n';
				const header = headerText.endsWith('\n') ? headerText : headerText + eol;
				await setDocumentText(header + text.slice(vb6HeaderEndOf(text)), false);
			};

			/**
			 * The VBE's double-click: the control's default event handler in
			 * the code below the header - navigate to it, or append the stub
			 * the VBE would write and land the cursor inside it. A control
			 * array's handler takes the element's Index.
			 */
			const openHandler = async (controlName: string, eventName: string): Promise<void> => {
				if (!eventName) {
					void vscode.window.showInformationMessage(`XLIDE: ${controlName || 'this control'} has no events.`);
					return;
				}
				try {
					const text = document.getText();
					const owner = controlName === '' ? vb6FormHandlerPrefix(text) : controlName.replace(/\(\d+\)$/, '');
					const isArray = /\(\d+\)$/.test(controlName);
					const handler = `${owner}_${eventName}`;
					const escaped = handler.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
					const signature = new RegExp(`^[ \\t]*(?:Private\\s+|Public\\s+|Friend\\s+)?Sub\\s+${escaped}\\s*\\(`, 'im');
					if (!signature.test(text)) {
						const eol = document.eol === vscode.EndOfLine.CRLF ? '\r\n' : '\n';
						const lead = text.length === 0 ? '' : text.endsWith(eol) ? eol : eol + eol;
						const params = isArray ? 'Index As Integer' : '';
						const edit = new vscode.WorkspaceEdit();
						edit.insert(document.uri, document.positionAt(text.length), `${lead}Private Sub ${handler}(${params})${eol}${eol}End Sub${eol}`);
						await vscode.workspace.applyEdit(edit);
					}
					const match = signature.exec(document.getText());
					const editor = await vscode.window.showTextDocument(document, {
						viewColumn: vscode.ViewColumn.One,
						preserveFocus: false,
					});
					if (match) {
						const line = Math.min(document.positionAt(match.index).line + 1, document.lineCount - 1);
						const at = new vscode.Position(line, 0);
						editor.selection = new vscode.Selection(at, at);
						editor.revealRange(new vscode.Range(at, at), vscode.TextEditorRevealType.InCenterIfOutsideViewport);
					}
				} catch (err) {
					void vscode.window.showErrorMessage(`XLIDE: ${errorMessage(err)}`);
				}
			};

			// Text changes from any face - the pane, undo, a text editor on the
			// same file - repaint the canvas after a breath. A gesture's own
			// echo is already rendered.
			let changeTimer: ReturnType<typeof setTimeout> | undefined;
			const changeListener = vscode.workspace.onDidChangeTextDocument((e) => {
				if (e.document !== document || e.contentChanges.length === 0) { return; }
				if (suppressEcho !== null && document.getText() === suppressEcho) {
					suppressEcho = null;
					return;
				}
				suppressEcho = null;
				if (changeTimer) { clearTimeout(changeTimer); }
				changeTimer = setTimeout(() => enqueue(async () => {
					if (document.getText() === lastRenderedText) { return; }
					await render();
				}), 250);
			});

			const messageListener = panel.webview.onDidReceiveMessage((message: DesignerMessage) => {
				switch (message.type) {
					case 'openHandler':
						void openHandler(message.name, message.event);
						break;
					case 'markupEdit':
						enqueue(() => applyMarkupEdit(message.text));
						break;
					case 'docUndo':
						enqueue(async () => { await vscode.commands.executeCommand('undo'); });
						break;
					case 'docRedo':
						enqueue(async () => { await vscode.commands.executeCommand('redo'); });
						break;
					case 'docSave':
						enqueue(async () => { await vscode.workspace.save(document.uri); });
						break;
					default:
						enqueue(() => applyGesture(message));
						break;
				}
			});

			panel.onDidDispose(() => {
				disposed = true;
				if (changeTimer) { clearTimeout(changeTimer); }
				changeListener.dispose();
				messageListener.dispose();
			});

			enqueue(() => render());
		}
	}

	context.subscriptions.push(
		vscode.window.registerCustomEditorProvider(
			VB6_FORM_DESIGNER_VIEW_TYPE,
			new Vb6FormDesignerProvider(),
			{
				webviewOptions: { retainContextWhenHidden: true },
				supportsMultipleEditorsPerDocument: false,
			},
		),
	);
}
