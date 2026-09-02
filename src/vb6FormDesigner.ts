// The VB6 form designer: the MSForms designer's canvas over a form's own
// file. The document IS the .frm (or .ctl, .pag), so there is no scratch
// copy and no apply. The canvas renders from the document's text; a gesture
// asks the engine for the header rewritten and lands as one edit over the
// header alone - the span the parser bounds - so the code below is never
// inside an edit, and text undo is the undo. A save is the file's own save,
// and the sidecar records a gesture placed (a multi-line Text) go into the
// .frx right before it, so the file and its sidecar move together. The pane
// below the canvas shows the header as written, and an edit there replaces
// the header and leaves the code alone.

import * as vscode from 'vscode';
import type { ProjectEngine } from './projectEngine';
import { moduleLocationOfDocument } from './vbaDocumentLocation';
import { vb6FormHandlerPrefix, vb6HeaderEndOf, vb6PendingRecordsToWrite } from './vba/vb6/frmDesignerOps';
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

interface SidecarRecord {
	file: string;
	base: number;
	offset: number;
	record: string;
}

interface OpResult {
	text: string;
	headerEnd: number;
	oldHeaderEnd: number;
	newName?: string;
	newNames?: string[];
	sidecar?: SidecarRecord;
}

/**
 * Sidecar records a document's designer placed that the file has not taken
 * yet, keyed by document URI: appended to the sidecar right before the
 * document saves, dropped when the document closes unsaved. Kept outside
 * the panel because the document can outlive it.
 */
interface PendingSidecar {
	file: string;
	base: number;
	records: string[];
	/** Each record's offset, as the header spells it in its reference. */
	offsets: number[];
	bytes: number;
}

const pendingSidecars = new Map<string, PendingSidecar>();

export function registerVb6FormDesigner(
	context: vscode.ExtensionContext,
	bridge: Pick<ProjectEngine, 'call'>,
): void {
	context.subscriptions.push(
		vscode.workspace.onWillSaveTextDocument((e) => {
			const key = e.document.uri.toString();
			const pending = pendingSidecars.get(key);
			if (!pending || pending.records.length === 0) { return; }
			const records = vb6PendingRecordsToWrite(pending.file, pending.records, pending.offsets, e.document.getText());
			if (records.length === 0) {
				pendingSidecars.delete(key);
				return;
			}
			e.waitUntil((async (): Promise<vscode.TextEdit[]> => {
				try {
					await bridge.call('vb6AppendSidecar', {
						path: e.document.uri.fsPath,
						file: pending.file,
						base: pending.base,
						records,
					});
					pendingSidecars.delete(key);
				} catch (err) {
					// The document still saves; the values stay pending for the
					// next save, and the message says what to do about them.
					void vscode.window.showErrorMessage(`XLIDE: ${errorMessage(err)}`);
				}
				return [];
			})());
		}),
		vscode.workspace.onDidCloseTextDocument((doc) => {
			pendingSidecars.delete(doc.uri.toString());
		}),
	);

	class Vb6FormDesignerProvider implements vscode.CustomTextEditorProvider {
		async resolveCustomTextEditor(
			document: vscode.TextDocument,
			panel: vscode.WebviewPanel,
		): Promise<void> {
			const modulePath = document.uri.fsPath;
			const key = document.uri.toString();
			const vbpPath = moduleLocationOfDocument(document)?.projectPath;
			panel.webview.options = { enableScripts: true };

			let disposed = false;
			/** Text this provider just placed in the document (a gesture echo). */
			let suppressEcho: string | null = null;
			/** The document text the webview last rendered, to skip repeats. */
			let lastRenderedText: string | null = null;
			/** The header text the pane placed last, which need not parse while it is being typed. */
			let paneHeader: string | null = null;
			/** All work runs through one lane: edits never interleave. */
			let queue: Promise<void> = Promise.resolve();

			const enqueue = (work: () => Promise<void>): void => {
				queue = queue.then(work).catch((err) => {
					if (!disposed) {
						void vscode.window.showErrorMessage(`XLIDE: ${errorMessage(err)}`);
					}
				});
			};

			const pendingOf = (): PendingSidecar | undefined => pendingSidecars.get(key);

			const render = async (selected?: string): Promise<void> => {
				if (disposed) { return; }
				const text = document.getText();
				const pending = pendingOf();
				try {
					const { html } = await bridge.call<{ html: string }>('readVb6FormPreview', {
						path: modulePath,
						text,
						selected,
						vbpPath,
						pending: pending ? { file: pending.file, base: pending.base, records: pending.records } : undefined,
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
			 * Replaces the document's header - the span [0, end) - with
			 * `header` as one edit. The code below the header is never in
			 * the edit, so a text editor open on the same file keeps what it
			 * typed there meanwhile.
			 */
			const replaceHeader = async (end: number, header: string, suppress: boolean): Promise<void> => {
				const current = document.getText();
				if (current.slice(0, end) === header) { return; }
				suppressEcho = suppress ? header + current.slice(end) : null;
				const edit = new vscode.WorkspaceEdit();
				edit.replace(document.uri, new vscode.Range(document.positionAt(0), document.positionAt(end)), header);
				const applied = await vscode.workspace.applyEdit(edit);
				if (!applied) { suppressEcho = null; }
			};

			/** A record a gesture placed joins the document's pending sidecar. */
			const takePending = (sidecar: SidecarRecord): void => {
				const bytes = Buffer.from(sidecar.record, 'base64').length;
				const pending = pendingOf();
				if (!pending) {
					pendingSidecars.set(key, { file: sidecar.file, base: sidecar.base, records: [sidecar.record], offsets: [sidecar.offset], bytes });
					return;
				}
				pending.records.push(sidecar.record);
				pending.offsets.push(sidecar.offset);
				pending.bytes += bytes;
			};

			const applyOp = async (op: Record<string, unknown>, selectAfter: (result: OpResult) => string | undefined): Promise<void> => {
				try {
					const text = document.getText();
					const result = await bridge.call<OpResult>('vb6FormDesignerOp', {
						path: modulePath,
						text,
						op,
						pendingBytes: pendingOf()?.bytes ?? 0,
					});
					// The gesture rewrote the header of the text it was given. The
					// document may have moved on below that header meanwhile; only
					// the header span is replaced, and only while it is still the
					// header the gesture saw.
					if (!document.getText().startsWith(text.slice(0, result.oldHeaderEnd))) {
						await render();
						return;
					}
					if (result.sidecar) { takePending(result.sidecar); }
					await replaceHeader(result.oldHeaderEnd, result.text.slice(0, result.headerEnd), true);
					paneHeader = null;
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

			/**
			 * The pane's text is the header; the code below it stays. The
			 * text arrives with the browser's line endings and takes the
			 * document's, so a CRLF form never gets an LF header.
			 */
			const applyMarkupEdit = async (headerText: string): Promise<void> => {
				const text = document.getText();
				const eol = document.eol === vscode.EndOfLine.CRLF ? '\r\n' : '\n';
				let header = headerText.replace(/\r?\n/g, eol);
				if (!header.endsWith(eol)) { header += eol; }
				const end = vb6HeaderEndOf(text, paneHeader ?? undefined);
				if (end === undefined) {
					void panel.webview.postMessage({
						type: 'markupError',
						message: 'The document no longer opens with a form header this text can replace; edit the file in the text editor.',
					});
					return;
				}
				paneHeader = header;
				await replaceHeader(end, header, false);
			};

			/**
			 * The VBE's double-click: the control's default event handler in
			 * the code below the header - navigate to it, or append the stub
			 * the VBE would write and land the cursor inside it. A control
			 * array's handler takes the element's Index. Runs on the queue,
			 * so its edit never crosses a gesture's.
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
						enqueue(() => openHandler(message.name, message.event));
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
