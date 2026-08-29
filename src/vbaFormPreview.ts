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

/** The workbook behind the most recently focused designer panel, for F5. */
let lastFocusedDesignerWorkbook: string | undefined;
/** The panel key of the most recently focused designer, for undo/redo. */
let lastFocusedDesignerKey: string | undefined;

interface DesignerHistory {
	undo: Array<Record<string, string>>;
	redo: Array<Record<string, string>>;
}
const histories = new Map<string, DesignerHistory>();
const keyParts = new Map<string, [string, string]>();
const HISTORY_LIMIT = 50;

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
	| { type: 'splitCollapse'; which: 'self' | 'below' }
	| { type: 'splitStateQuery' }
	| { type: 'launchHost' }
	| { type: 'splitDrag'; phase: 'start' | 'end' }
	| { type: 'splitDrag'; phase: 'move'; delta: number }
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
		message: Exclude<DesignerMessage, { type: 'openHandler' } | { type: 'splitCollapse' } | { type: 'splitDrag' } | { type: 'splitStateQuery' } | { type: 'launchHost' }>,
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
			const key = panelKey(xlsmPath, moduleName);
			try {
				const snapshot = await bridge.call<{ streams: Record<string, string> }>(
					'readFormDesignerSnapshot',
					{ path: xlsmPath, module: moduleName },
				);
				const history = histories.get(key) ?? { undo: [], redo: [] };
				history.undo.push(snapshot.streams);
				if (history.undo.length > HISTORY_LIMIT) { history.undo.shift(); }
				history.redo = [];
				histories.set(key, history);
			} catch {
				// A gesture without its snapshot still applies; only its undo
				// step is lost.
			}
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

	// Ctrl+Z / Ctrl+Y while the designer is focused: the gesture history is
	// a stack of byte-true designer snapshots, restored whole - so undo puts
	// back exactly what the gesture changed, wherever in the form it landed.
	const stepHistory = async (direction: 'undo' | 'redo'): Promise<void> => {
		const key = lastFocusedDesignerKey;
		const panel = key ? panels.get(key) : undefined;
		const history = key ? histories.get(key) : undefined;
		if (!key || !panel || !history) { return; }
		const from = direction === 'undo' ? history.undo : history.redo;
		const to = direction === 'undo' ? history.redo : history.undo;
		const snapshot = from.pop();
		if (!snapshot) { return; }
		const [xlsmPath, moduleName] = keyParts.get(key) ?? [];
		if (!xlsmPath || !moduleName) { return; }
		try {
			const current = await bridge.call<{ streams: Record<string, string> }>(
				'readFormDesignerSnapshot',
				{ path: xlsmPath, module: moduleName },
			);
			to.push(current.streams);
			if (to.length > HISTORY_LIMIT) { to.shift(); }
			await runWriteWithExcelCoordination(xlsmPath, () =>
				bridge.call('restoreFormDesignerSnapshot', { path: xlsmPath, module: moduleName, streams: snapshot }));
			fsProvider?.notifyFileChanged(encodeFormMarkupUri(xlsmPath, moduleName));
			await render(panel, xlsmPath, moduleName);
		} catch (err) {
			from.push(snapshot);
			void vscode.window.showErrorMessage(`XLIDE: ${errorMessage(err)}`);
		}
	};

	// The split between the designer and the markup below, resized with
	// LAYOUT-TREE pixel math (vscode.get/setEditorLayout): the drag is
	// continuous, its direction is the mouse's by construction, and nothing
	// outside the designer+markup column can move - the two earlier
	// mechanisms failed all three ways (the workbench maximize swallowed the
	// window; focus-dependent height steps were coarse and could grab the
	// wrong group).
	interface LayoutNode { groups?: LayoutNode[]; size?: number }
	interface SplitGrab {
		layout: { orientation: number; groups: LayoutNode[] };
		leaf: LayoutNode;
		sibling: LayoutNode;
		leafBase: number;
		siblingBase: number;
		/** Latest unapplied delta; applies coalesce to it. */
		pending: number | null;
		applying: boolean;
		lastApplied: number;
	}
	const splitGrabs = new Map<string, SplitGrab>();
	/** Per-panel collapse toggle: the sizes a collapse replaced, to restore. */
	const splitToggles = new Map<string, { saved?: { leaf: number; sibling: number }; collapsed: 'self' | 'below' | null }>();

	const grabSplit = async (panel: vscode.WebviewPanel): Promise<SplitGrab | undefined> => {
		const layout = await vscode.commands.executeCommand<{ orientation: number; groups: LayoutNode[] }>('vscode.getEditorLayout');
		if (!layout?.groups) { return undefined; }
		// Leaves in depth-first order correspond to tabGroups.all order.
		const leaves: Array<{ node: LayoutNode; parent: LayoutNode[]; index: number; vertical: boolean }> = [];
		const walk = (nodes: LayoutNode[], vertical: boolean): void => {
			for (const node of nodes) {
				if (node.groups?.length) { walk(node.groups, !vertical); }
				else { leaves.push({ node, parent: nodes, index: nodes.indexOf(node), vertical }); }
			}
		};
		// orientation 0 lays the root row out horizontally, so root children
		// stack vertically only when orientation is 1.
		walk(layout.groups, layout.orientation === 1);
		const groups = vscode.window.tabGroups.all;
		const at = groups.findIndex((group) => group.tabs.some((tab) =>
			tab.input instanceof vscode.TabInputWebview
			&& tab.input.viewType.includes('xlideFormPreview')
			&& tab.label === panel.title));
		const leaf = at >= 0 ? leaves[at] : undefined;
		if (!leaf || !leaf.vertical) { return undefined; }
		const sibling = leaf.parent[leaf.index + 1] ?? leaf.parent[leaf.index - 1];
		if (!sibling || sibling === leaf.node) { return undefined; }
		if (typeof leaf.node.size !== 'number' || typeof sibling.size !== 'number') { return undefined; }
		return {
			layout,
			leaf: leaf.node,
			sibling,
			leafBase: leaf.node.size,
			siblingBase: sibling.size,
			pending: null,
			applying: false,
			lastApplied: leaf.node.size,
		};
	};

	// Applies COALESCE: deltas stream faster than setEditorLayout runs, and
	// un-serialized applies interleave out of order - the stutter - while
	// the sheer rate thrashed the grid - the flicker. One apply loop per
	// grab drains only the LATEST delta, on integer sizes, skipping no-ops.
	const moveSplit = async (key: string, delta: number): Promise<void> => {
		const grab = splitGrabs.get(key);
		if (!grab) { return; }
		grab.pending = delta;
		if (grab.applying) { return; }
		grab.applying = true;
		try {
			while (grab.pending !== null && splitGrabs.get(key) === grab) {
				const next = grab.pending;
				grab.pending = null;
				const total = grab.leafBase + grab.siblingBase;
				const MIN = 40;
				const leafSize = Math.round(Math.min(total - MIN, Math.max(MIN, grab.leafBase + next)));
				if (leafSize === grab.lastApplied) { continue; }
				grab.leaf.size = leafSize;
				grab.sibling.size = total - leafSize;
				try {
					await vscode.commands.executeCommand('vscode.setEditorLayout', grab.layout);
					grab.lastApplied = leafSize;
				} catch {
					// The layout changed under the drag; the next grab starts fresh.
					splitGrabs.delete(key);
					break;
				}
			}
		} finally {
			grab.applying = false;
		}
	};

	const wirePanel = (panel: vscode.WebviewPanel, xlsmPath: string, moduleName: string): void => {
		const key = panelKey(xlsmPath, moduleName);
		panels.set(key, panel);
		keyParts.set(key, [xlsmPath, moduleName]);
		lastFocusedDesignerWorkbook = xlsmPath;
		lastFocusedDesignerKey = key;
		panel.onDidChangeViewState(
			(e) => {
				if (e.webviewPanel.active) {
					lastFocusedDesignerWorkbook = xlsmPath;
					lastFocusedDesignerKey = key;
				}
			},
			undefined,
			context.subscriptions,
		);
		panel.onDidDispose(() => panels.delete(key), undefined, context.subscriptions);
		panel.webview.onDidReceiveMessage(
			(message: DesignerMessage) => {
				if (message.type === 'openHandler') {
					void openEventHandler(xlsmPath, moduleName, message.name, message.event);
				} else if (message.type === 'splitDrag') {
					void (async () => {
						if (message.phase === 'start') {
							const grab = await grabSplit(panel);
							if (grab) { splitGrabs.set(key, grab); }
							// A hand on the sash means the collapse is over.
							const toggle = splitToggles.get(key);
							if (toggle && toggle.collapsed !== null) {
								toggle.collapsed = null;
								void panel.webview.postMessage({ type: 'splitState', collapsed: null });
							}
						} else if (message.phase === 'move') {
							await moveSplit(key, message.delta);
						} else {
							splitGrabs.delete(key);
						}
					})();
				} else if (message.type === 'splitCollapse') {
					// The arrows TOGGLE: the first press saves the split and
					// pushes it to the clamp; pressing the flipped arrow puts
					// the saved split back - the vbide behavior.
					void (async () => {
						const state = splitToggles.get(key) ?? { collapsed: null };
						const grab = await grabSplit(panel);
						if (!grab) { return; }
						splitGrabs.set(key, grab);
						const total = grab.leafBase + grab.siblingBase;
						if (state.collapsed === message.which && state.saved) {
							await moveSplit(key, state.saved.leaf - grab.leafBase);
							state.collapsed = null;
						} else {
							if (state.collapsed === null) {
								state.saved = { leaf: grab.leafBase, sibling: grab.siblingBase };
							}
							await moveSplit(key, message.which === 'self' ? -total : total);
							state.collapsed = message.which;
						}
						splitToggles.set(key, state);
						splitGrabs.delete(key);
						void panel.webview.postMessage({ type: 'splitState', collapsed: state.collapsed });
					})();
				} else if (message.type === 'launchHost') {
					void vscode.commands.executeCommand('xlide.launchFormHost');
				} else if (message.type === 'splitStateQuery') {
					void panel.webview.postMessage({
						type: 'splitState',
						collapsed: splitToggles.get(key)?.collapsed ?? null,
					});
				} else {
					void applyGesture(panel, xlsmPath, moduleName, message);
				}
			},
			undefined,
			context.subscriptions,
		);
	};

	// A window reload restores text editors natively; the designer needs
	// this serializer to come back beside them. Its identity rides in the
	// webview state the canvas stamps on every load.
	if (typeof vscode.window.registerWebviewPanelSerializer === 'function') {
		context.subscriptions.push(vscode.window.registerWebviewPanelSerializer('xlideFormPreview', {
			deserializeWebviewPanel: async (
				panel: vscode.WebviewPanel,
				state?: { wb?: string; mod?: string },
			): Promise<void> => {
				if (!state?.wb || !state?.mod) {
					panel.dispose();
					return;
				}
				wirePanel(panel, state.wb, state.mod);
				await render(panel, state.wb, state.mod);
			},
		}));
	}

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
			wirePanel(panel, xlsmPath, moduleName);
			await render(panel, xlsmPath, moduleName);
			await showMarkupBelow(panel);
		}),

		vscode.commands.registerCommand('xlide.designerUndo', () => stepHistory('undo')),
		vscode.commands.registerCommand('xlide.designerRedo', () => stepHistory('redo')),

		// F5 from the designer or the markup launches the workbook's host
		// application - the closest thing to the VBE's Run while the engine
		// stays COMless. The active markup editor names the workbook; a
		// focused designer panel remembered its own.
		vscode.commands.registerCommand('xlide.launchFormHost', async () => {
			const active = vscode.window.activeTextEditor;
			const fromEditor = active && active.document.uri.scheme === XLIDE_SCHEME
				? decodeModuleUri(active.document.uri).xlsmPath
				: undefined;
			const filePath = fromEditor ?? lastFocusedDesignerWorkbook;
			if (!filePath) {
				void vscode.window.showInformationMessage('XLIDE: F5 found no form workbook - focus a designer or a markup document.');
				return;
			}
			vscode.window.setStatusBarMessage(`XLIDE: opening ${path.basename(filePath)} in its host application...`, 5000);
			const excel = /\.(xlsm|xlsb|xlam|xls)$/i.test(filePath);
			await vscode.commands.executeCommand(
				excel ? 'xlide.openWorkbook' : 'xlide.openInOfficeApp',
				{ kind: 'xlsm', label: path.basename(filePath), filePath },
			);
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
