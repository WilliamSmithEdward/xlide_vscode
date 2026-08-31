// The form designer: ONE custom editor over the .form document - the canvas,
// the properties pane, and the markup text live in a single tab, the way the
// vbide draws them. The document is the truth the user edits; a SCRATCH COPY
// of the workbook holds that truth applied, so the canvas renders real bytes
// without ever writing the user's file. A gesture mutates the scratch through
// the engine's designer ops, prints back to markup, and lands in the document
// as an ordinary text edit - so the tab carries the dirty dot, Ctrl+Z is text
// undo, and SAVE is the only workbook write (the .form provider's apply).

import * as vscode from 'vscode';
import * as path from 'path';
import * as fs from 'fs';
import * as crypto from 'crypto';
import type { WorkbookEngine } from './workbookEngine';
import { decodeModuleUri, encodeFormMarkupUri, encodeModuleUri, XLIDE_SCHEME } from './xlideFileSystem';
import {
	closeWorkbookInExcel,
	markWorkbookOpenedByXlide,
	resolveExcelCoordinationSettings,
	runWriteWithExcelCoordination,
	shouldAttemptClose,
	withWorkbookReopenSuppressed,
} from './excelWorkbookCoordinator';
import { ExcelMacroError, runWorkbookMacroReadOnly } from './excelLauncher';
import {
	composeLauncherSource,
	launcherSubExists,
	launcherSubName,
	LAUNCHER_MODULE,
} from './vbaFormLauncher';
import { xlideAttachToRunningExcelFromConfig } from './globalSettings';
import { errorMessage } from './util/errors';

export const FORM_DESIGNER_VIEW_TYPE = 'xlideFormDesigner';

/** The workbook behind the most recently focused designer, for F5. */
let lastFocusedDesignerWorkbook: string | undefined;
/** Its form module, so F5 knows which form a Show launcher should open. */
let lastFocusedDesignerModule: string | undefined;
/** A launch (its consent modal included) is running; a second F5 waits. */
let launchInFlight = false;

/**
 * Persists what F5 is about to run: the form's own document (the designer
 * writes its gestures there) and the focused XLIDE document, when either is
 * dirty. False means a save was refused - the caller must NOT run on, because
 * the workbook still holds the previous version of the form.
 */
async function savePendingLaunchEdits(
	filePath: string,
	formModule: string | undefined,
): Promise<boolean> {
	const pending: vscode.TextDocument[] = [];
	const active = vscode.window.activeTextEditor?.document;
	if (active && active.uri.scheme === XLIDE_SCHEME && active.isDirty) {
		pending.push(active);
	}
	if (formModule) {
		// The designer's document, whether or not it is the focused editor -
		// F5 from the canvas has no active text editor at all.
		const formUri = encodeFormMarkupUri(filePath, formModule).toString();
		for (const doc of vscode.workspace.textDocuments) {
			if (doc.uri.toString() === formUri && doc.isDirty && !pending.includes(doc)) {
				pending.push(doc);
			}
		}
	}
	for (const doc of pending) {
		try {
			if (!(await doc.save())) { return false; }
		} catch {
			return false;
		}
	}
	return true;
}

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

type GestureMessage = Extract<DesignerMessage,
	{ type: 'geometry' } | { type: 'geometryBatch' } | { type: 'add' } | { type: 'remove' }
	| { type: 'reparent' } | { type: 'setProp' } | { type: 'formResize' }
	| { type: 'zOrder' } | { type: 'tabOrder' }>;

export function registerFormPreview(
	context: vscode.ExtensionContext,
	bridge: Pick<WorkbookEngine, 'call'>,
): void {
	const scratchDir = path.join(context.globalStorageUri.fsPath, 'designer-scratch');
	try {
		fs.mkdirSync(scratchDir, { recursive: true });
		// Scratches from editors that never got to clean up (a crash, a kill).
		const weekAgo = Date.now() - 7 * 24 * 60 * 60 * 1000;
		for (const name of fs.readdirSync(scratchDir)) {
			const file = path.join(scratchDir, name);
			try {
				if (fs.statSync(file).mtimeMs < weekAgo) { fs.unlinkSync(file); }
			} catch { /* someone else's scratch; leave it */ }
		}
	} catch { /* the first render will surface a real failure */ }

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

	class FormDesignerProvider implements vscode.CustomTextEditorProvider {
		async resolveCustomTextEditor(
			document: vscode.TextDocument,
			panel: vscode.WebviewPanel,
		): Promise<void> {
			const { xlsmPath, moduleName } = decodeModuleUri(document.uri);
			panel.webview.options = { enableScripts: true };
			lastFocusedDesignerWorkbook = xlsmPath;
			lastFocusedDesignerModule = moduleName;

			// The pid keeps two windows on the same form from sharing a scratch;
			// a crash's orphan falls to the age sweep above.
			const scratchPath = path.join(
				scratchDir,
				crypto.createHash('sha1')
					.update(`${xlsmPath.toLowerCase()}::${moduleName.toLowerCase()}`)
					.digest('hex').slice(0, 20) + `-${process.pid}${path.extname(xlsmPath)}`,
			);
			/** Real workbook mtime the scratch was last copied from. */
			let baselineMtime = -1;
			/** Document text currently applied to the scratch; null = bare copy. */
			let appliedText: string | null = null;
			/** Text this provider just placed in the document (a gesture echo). */
			let suppressEcho: string | null = null;
			/** All work runs through one lane: edits never interleave. */
			let queue: Promise<void> = Promise.resolve();
			let disposed = false;

			const enqueue = (work: () => Promise<void>): void => {
				queue = queue.then(work).catch((err) => {
					if (!disposed) {
						void vscode.window.showErrorMessage(`XLIDE: ${errorMessage(err)}`);
					}
				});
			};

			/** Copies the real workbook under the scratch when it moved. */
			const ensureScratch = (): void => {
				const mtime = Math.floor(fs.statSync(xlsmPath).mtimeMs);
				if (mtime !== baselineMtime || !fs.existsSync(scratchPath)) {
					fs.copyFileSync(xlsmPath, scratchPath);
					baselineMtime = mtime;
					appliedText = null;
				}
			};

			/**
			 * Makes the scratch say exactly what the document says: a fresh
			 * baseline copy plus one whole-document apply. Rebuilding from the
			 * baseline is what makes text UNDO honest - an attribute line that
			 * disappears returns to the saved value, not to the edit before.
			 */
			const syncScratchToText = async (text: string): Promise<void> => {
				ensureScratch();
				if (appliedText === text) { return; }
				if (appliedText !== null) {
					fs.copyFileSync(xlsmPath, scratchPath);
					appliedText = null;
				}
				await bridge.call('applyFormMarkup', { path: scratchPath, module: moduleName, markup: text });
				appliedText = text;
			};

			/** The document text the webview last rendered, to skip repeats. */
			let lastRenderedText: string | null = null;
			const render = async (selected?: string): Promise<void> => {
				if (disposed) { return; }
				try {
					const markup = document.getText();
					const { html } = await bridge.call<{ html: string }>('readFormPreview', {
						path: scratchPath,
						module: moduleName,
						selected,
						markup,
						identityPath: xlsmPath,
					});
					panel.webview.html = html;
					lastRenderedText = markup;
				} catch (err) {
					lastRenderedText = null;
					panel.webview.html = `<!DOCTYPE html><html><body style="font-family:sans-serif;padding:16px">`
						+ `<p>XLIDE could not render this form.</p><pre>${errorMessage(err)
							.replace(/&/g, '&amp;').replace(/</g, '&lt;')}</pre></body></html>`;
				}
			};

			/** A text change lands on the canvas, or its error on the strip. */
			const syncAndRender = async (): Promise<void> => {
				if (disposed) { return; }
				try {
					await syncScratchToText(document.getText());
				} catch (err) {
					void panel.webview.postMessage({ type: 'markupError', message: errorMessage(err) });
					return;
				}
				void panel.webview.postMessage({ type: 'markupOk' });
				// A pending typing debounce can fire right after a gesture
				// already rendered this very text; repeating the render only
				// flickers.
				if (document.getText() === lastRenderedText) { return; }
				await render();
			};

			/**
			 * Replaces the whole document with `text` as one undo step. A
			 * gesture SUPPRESSES its own echo (its scratch is already applied
			 * and rendered); a markup-pane edit must NOT - the change listener
			 * is exactly what applies it and repaints the canvas.
			 */
			const setDocumentText = async (text: string, suppress: boolean): Promise<void> => {
				if (document.getText() === text) { return; }
				suppressEcho = suppress ? text : null;
				const edit = new vscode.WorkspaceEdit();
				edit.replace(
					document.uri,
					new vscode.Range(new vscode.Position(0, 0), document.positionAt(document.getText().length)),
					text,
				);
				const applied = await vscode.workspace.applyEdit(edit);
				if (!applied) { suppressEcho = null; }
			};

			/**
			 * The selection gestures the engine answers by rewriting the whole
			 * MARKUP rather than by moving sites: paste and multi-delete. Both
			 * land as ONE document edit, so a three-control delete is a single
			 * Ctrl+Z, and both name their result so the canvas can select it.
			 */
			const applyMarkupTransform = async <T>(
				verb: string,
				method: string,
				names: string[],
				selectAfter: (result: T) => string | undefined,
			): Promise<void> => {
				try {
					await syncScratchToText(document.getText());
					const result = await bridge.call<T>(
						method, { path: scratchPath, module: moduleName, names },
					);
					const { markup } = await bridge.call<{ markup: string }>(
						'readFormMarkup',
						{ path: scratchPath, module: moduleName },
					);
					appliedText = markup;
					await setDocumentText(markup, true);
					await render(selectAfter(result));
				} catch (err) {
					void vscode.window.showErrorMessage(`XLIDE: could not ${verb}: ${errorMessage(err)}`);
					await render();
				}
			};

			/** Paste duplicates by NAME, and lands selected on the first copy. */
			const applyPaste = (names: string[]): Promise<void> =>
				applyMarkupTransform<{ newNames: string[] }>(
					'paste', 'duplicateFormControls', names, (r) => r.newNames[0]);

			/** Deleting a multi-selection: every name goes in one edit. */
			const applyRemoveMany = (names: string[]): Promise<void> =>
				applyMarkupTransform<{ removed: string[] }>(
					'delete', 'removeFormControls', names, () => undefined);

			const applyGesture = async (message: GestureMessage): Promise<void> => {
				const op = message.type === 'geometry'
					? { kind: 'geometry' as const, name: message.name, left: message.left, top: message.top, width: message.width, height: message.height }
					: message.type === 'geometryBatch'
						? { kind: 'geometryBatch' as const, items: message.items }
						: message.type === 'add'
							? { kind: 'add' as const, container: message.container, controlKind: message.controlKind, left: message.left, top: message.top }
							: message.type === 'reparent'
								? { kind: 'reparent' as const, name: message.name, container: message.container, left: message.left, top: message.top }
								: message.type === 'setProp'
									? { kind: 'setProp' as const, name: message.name, prop: message.prop, value: message.value }
									: message.type === 'formResize'
										? { kind: 'formSize' as const, width: message.width, height: message.height }
										: message.type === 'zOrder'
											? { kind: 'zOrder' as const, name: message.name, toFront: message.toFront }
											: message.type === 'tabOrder'
												? { kind: 'tabOrder' as const, container: message.container, names: message.names }
												: { kind: 'remove' as const, name: message.name };
				try {
					await syncScratchToText(document.getText());
					const result = await bridge.call<{ ok: boolean; newName?: string }>(
						'formDesignerOp',
						{ path: scratchPath, module: moduleName, op },
					);
					const { markup } = await bridge.call<{ markup: string }>(
						'readFormMarkup',
						{ path: scratchPath, module: moduleName },
					);
					appliedText = markup;
					await setDocumentText(markup, true);
					const keepSelected = message.type === 'remove'
						? undefined
						: message.type === 'geometryBatch'
							? message.anchor
							: message.type === 'formResize'
								? ''
								: message.type === 'tabOrder'
									? undefined
									: message.type === 'setProp'
										? (result.newName ?? message.name)
										: message.type === 'add' ? result.newName : message.name;
					await render(keepSelected);
				} catch (err) {
					void vscode.window.showErrorMessage(`XLIDE: ${errorMessage(err)}`);
					await render();
				}
			};

			// Text changes from ANY face - the markup pane, undo/redo, a
			// reopened text editor - re-apply to the scratch after a breath.
			// A gesture's own echo is already applied and already rendered.
			let changeTimer: ReturnType<typeof setTimeout> | undefined;
			const changeListener = vscode.workspace.onDidChangeTextDocument((e) => {
				if (e.document !== document || e.contentChanges.length === 0) { return; }
				if (suppressEcho !== null && document.getText() === suppressEcho) {
					suppressEcho = null;
					return;
				}
				suppressEcho = null;
				if (changeTimer) { clearTimeout(changeTimer); }
				changeTimer = setTimeout(() => enqueue(syncAndRender), 250);
			});

			// A save writes the real workbook through the .form provider; the
			// scratch already says the same, so only the baseline stamp moves.
			const saveListener = vscode.workspace.onDidSaveTextDocument((doc) => {
				if (doc !== document) { return; }
				try {
					baselineMtime = Math.floor(fs.statSync(xlsmPath).mtimeMs);
				} catch { /* the next gesture recopies */ }
			});

			const viewStateListener = panel.onDidChangeViewState((e) => {
				if (e.webviewPanel.active) {
					lastFocusedDesignerWorkbook = xlsmPath;
					lastFocusedDesignerModule = moduleName;
				}
			});

			const messageListener = panel.webview.onDidReceiveMessage((message: DesignerMessage) => {
				switch (message.type) {
					case 'openHandler':
						void openEventHandler(xlsmPath, moduleName, message.name, message.event);
						break;
					case 'paste':
						enqueue(() => applyPaste(message.names));
						break;
					case 'removeMany':
						enqueue(() => applyRemoveMany(message.names));
						break;
					case 'markupEdit':
						enqueue(() => setDocumentText(message.text, false));
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
				saveListener.dispose();
				viewStateListener.dispose();
				messageListener.dispose();
				try { fs.unlinkSync(scratchPath); } catch { /* already gone */ }
			});

			// First light: the document may already be dirty (a restored
			// backup), so the scratch takes the document's word from the
			// start. A CLEAN document just came from the workbook itself, so
			// its text IS the baseline's print - trusting that skips a whole
			// parse-and-diff on every designer open.
			if (!document.isDirty) {
				try {
					ensureScratch();
					appliedText = document.getText();
				} catch { /* the first sync will surface the real failure */ }
			}
			enqueue(syncAndRender);
		}
	}

	context.subscriptions.push(
		vscode.window.registerCustomEditorProvider(
			FORM_DESIGNER_VIEW_TYPE,
			new FormDesignerProvider(),
			{
				webviewOptions: { retainContextWhenHidden: true },
				supportsMultipleEditorsPerDocument: false,
			},
		),

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
			await vscode.commands.executeCommand(
				'vscode.openWith',
				encodeFormMarkupUri(xlsmPath, moduleName),
				FORM_DESIGNER_VIEW_TYPE,
			);
		}),

		// F5 from the designer or a markup text editor launches the
		// workbook's host application - the closest thing to the VBE's Run
		// while the engine stays COMless. The active markup editor names the
		// workbook; a focused designer remembered its own.
		//
		// ONE launch at a time: the consent modal and the Excel run are both
		// long, and a second F5 arriving meanwhile must not stack a second
		// dialog on top of the first (measured when the canvas posted the
		// launch as well as the keybinding firing it).
		vscode.commands.registerCommand('xlide.launchFormHost', async () => {
			if (launchInFlight) { return; }
			launchInFlight = true;
			try {
				await launchFormHost();
			} finally {
				launchInFlight = false;
			}
		}),
	);

	async function launchFormHost(): Promise<void> {
		{
			const active = vscode.window.activeTextEditor;
			let filePath: string | undefined;
			let formModule: string | undefined;
			if (active && active.document.uri.scheme === XLIDE_SCHEME) {
				const decoded = decodeModuleUri(active.document.uri);
				filePath = decoded.xlsmPath;
				if (decoded.face === 'form') { formModule = decoded.moduleName; }
			}
			if (!filePath) {
				filePath = lastFocusedDesignerWorkbook;
				formModule = lastFocusedDesignerModule;
			}
			if (!filePath) {
				void vscode.window.showInformationMessage('XLIDE: F5 found no form workbook - focus a designer or a markup document.');
				return;
			}
			const wbPath = filePath;
			const excel = /\.(xlsm|xlsb|xlam|xls)$/i.test(wbPath);

			// F5 runs WHAT YOU SEE. The designer holds its gestures and markup
			// edits as pending document changes, so without this Excel would
			// faithfully show the LAST SAVED form and the change would look
			// like it had failed. (The Run-Macro command persists a dirty code
			// module for the same reason.) Suppressed, because the save's own
			// reopen would otherwise race the macro host's.
			const saved = excel
				? await withWorkbookReopenSuppressed(wbPath, () => savePendingLaunchEdits(wbPath, formModule))
				: await savePendingLaunchEdits(wbPath, formModule);
			if (!saved) {
				// The provider already surfaced WHY (a markup parse error names
				// its line); running on regardless would show a stale form.
				void vscode.window.showErrorMessage(
					'XLIDE: F5 did not launch - the pending changes could not be saved.',
				);
				return;
			}

			// The VBE's F5 SHOWS the form. Excel can only run a macro, so
			// with consent XLIDE injects a launcher and runs it; the choice
			// persists in xlide.formRun.injectShowMacro.
			//
			// ONE SUB PER FORM, all in module XlideRun: F5 on a second form
			// ADDS its sub beside the first rather than rewriting it, so the
			// launchers accumulate and each form keeps its own entry point.
			if (excel && formModule) {
				const config = vscode.workspace.getConfiguration('xlide');
				const subName = launcherSubName(formModule);
				const macro = `${LAUNCHER_MODULE}.${subName}`;
				// What the workbook already carries decides whether anything
				// is being injected at all.
				let launcherSource: string | undefined;
				try {
					launcherSource = (await bridge.call<{ source: string }>(
						'readModule', { path: wbPath, module: LAUNCHER_MODULE },
					)).source;
				} catch {
					launcherSource = undefined; // no launcher module yet
				}
				const subExists = launcherSubExists(launcherSource, subName);

				let mode = config.get<string>('formRun.injectShowMacro') ?? 'ask';
				// THIS form's launcher is already installed: nothing goes into
				// the workbook, so there is nothing to consent to - just run
				// it. An explicit Never still means never.
				if (mode === 'ask' && subExists) { mode = 'once'; }
				if (mode === 'ask') {
					const pick = await vscode.window.showInformationMessage(
						`Show ${formModule} on launch?`,
						{
							modal: true,
							detail: `XLIDE can add a small launcher macro (${subName}, in module XlideRun) to ${path.basename(wbPath)} and run it, so F5 behaves like the VBE's. `
								+ 'Each form gets its own sub; they stay in the workbook and are safe to delete. '
								+ '"Always" remembers this in the xlide.formRun.injectShowMacro setting.',
						},
						'Yes', 'Always', 'No',
					);
					if (pick === 'Always') {
						await config.update('formRun.injectShowMacro', 'always', vscode.ConfigurationTarget.Global);
						mode = 'always';
					} else if (pick === 'Yes') {
						mode = 'once';
					} else {
						mode = 'never-once';
					}
				}
				if (mode === 'always' || mode === 'once') {
					const attachToRunning = xlideAttachToRunningExcelFromConfig(config).value;
					const quiet = (): void => { /* the launcher logs on its own channel */ };
					const source = composeLauncherSource(launcherSource, formModule);
					try {
						// Suppression spans the write AND the run, as the Run-Macro
						// command does: the write's own background read-only refresh
						// would otherwise open the workbook in one Excel while the
						// macro host spawns another - two Excels for one F5.
						await withWorkbookReopenSuppressed(wbPath, async () => {
							// An installed launcher is run as it stands: no write, so
							// a repeat F5 never touches the workbook at all (and any
							// hand edit to the sub is honored rather than clobbered).
							if (!subExists) {
								await runWriteWithExcelCoordination(wbPath, () =>
									bridge.call('writeModule', { path: wbPath, module: LAUNCHER_MODULE, source }));
							}
							vscode.window.setStatusBarMessage(`XLIDE: showing ${formModule} in Excel...`, 8000);
							try {
								await runWorkbookMacroReadOnly(wbPath, macro, { attachToRunning }, quiet);
							} catch (err) {
								// Open for editing in Excel: honor the coordination
								// policy - close and retry - rather than asking the
								// user to close it by hand. block mode still rethrows.
								const settings = resolveExcelCoordinationSettings();
								if (err instanceof ExcelMacroError && err.code === 'REOPEN_BLOCKED'
									&& settings.mode !== 'block' && shouldAttemptClose(settings, wbPath)) {
									await closeWorkbookInExcel(wbPath, { force: settings.mode === 'closeForce' }, quiet);
									markWorkbookOpenedByXlide(wbPath);
									await runWorkbookMacroReadOnly(wbPath, macro, { attachToRunning }, quiet);
								} else {
									// RUN_FAILED means the host already reopened the
									// workbook before the macro raised; keep it tracked
									// so a later closeTracked save frees the lock.
									if (err instanceof ExcelMacroError && err.code === 'RUN_FAILED') {
										markWorkbookOpenedByXlide(wbPath);
									}
									throw err;
								}
							}
							markWorkbookOpenedByXlide(wbPath);
						});
					} catch (err) {
						void vscode.window.showErrorMessage(`XLIDE: could not show the form: ${errorMessage(err)}`);
					}
					return;
				}
			}

			vscode.window.setStatusBarMessage(`XLIDE: opening ${path.basename(wbPath)} in its host application...`, 5000);
			await vscode.commands.executeCommand(
				excel ? 'xlide.openWorkbook' : 'xlide.openInOfficeApp',
				{ kind: 'xlsm', label: path.basename(wbPath), filePath: wbPath },
			);
		}
	}
}
