// Registration wiring for the VBA completion stack.
//
// Constructs the per-window editor project-context service and wires the
// completion provider (vbaCompletionProvider.ts), hover/signature provider
// (vbaHoverSignatureProvider.ts), and canonical-case controller
// (vbaCanonicalCaseController.ts) into VS Code, plus the keyword-snippet
// leave-detection state machine and the workbook-save cache invalidation.

import * as vscode from 'vscode';
import {
	XLIDE_SCHEME,
	decodeModuleUri,
	isVbaDocument,
} from './xlideFileSystem';
import { DocRegistry } from './analyzer';
import { VbaProjectIndexService } from './vbaProjectIndexService';
import { VbaCanonicalCaseController } from './vbaCanonicalCaseController';
import { VbaEditorProjectContextService } from './vbaEditorProjectContext';
import { VbaHoverSignatureProvider } from './vbaHoverSignatureProvider';
import {
	KEYWORD_SNIPPET_ACCEPTED_COMMAND,
	VbaMemberCompletionProvider,
} from './vbaCompletionProvider';

const KEYBOARD_NAV_TEXT_CHANGE_GRACE_MS = 150;

const ACTIVE_MEMBER_COMPLETION_PROVIDERS = new Set<VbaMemberCompletionProvider>();

export function invalidateVbaMemberCompletionCache(xlsmPath?: string): void {
	for (const provider of ACTIVE_MEMBER_COMPLETION_PROVIDERS) {
		provider.invalidate(xlsmPath);
	}
}

/** Registers the host-context member completion provider. */
export function registerVbaMemberCompletion(
	context: vscode.ExtensionContext,
	projectIndexService: VbaProjectIndexService,
	selector: vscode.DocumentSelector,
	docs?: DocRegistry,
): void {
	const projectContext = new VbaEditorProjectContextService(projectIndexService);
	const provider = new VbaMemberCompletionProvider(projectContext);
	const hoverSignature = new VbaHoverSignatureProvider(projectContext, docs);
	const canonicalCase = new VbaCanonicalCaseController(projectContext);
	ACTIVE_MEMBER_COMPLETION_PROVIDERS.add(provider);
	context.subscriptions.push({
		dispose: () => ACTIVE_MEMBER_COMPLETION_PROVIDERS.delete(provider),
	});
	let activeKeywordSnippet:
		| { editor: vscode.TextEditor; documentKey: string; textChangeSerialAtAccept: number }
		| undefined;
	let textChangeSerial = 0;
	const lastTextChange = new Map<string, { at: number; serial: number }>();
	const markTextChange = (document: vscode.TextDocument): void => {
		if (isVbaDocument(document)) {
			textChangeSerial += 1;
			lastTextChange.set(document.uri.toString(), {
				at: Date.now(),
				serial: textChangeSerial,
			});
		}
	};
	const maybeLeaveKeywordSnippet = (event: vscode.TextEditorSelectionChangeEvent): void => {
		if (!activeKeywordSnippet || event.textEditor !== activeKeywordSnippet.editor) {
			return;
		}
		if (!isVbaDocument(event.textEditor.document)) {
			activeKeywordSnippet = undefined;
			return;
		}
		if (
			event.kind !== vscode.TextEditorSelectionChangeKind.Mouse &&
			event.kind !== vscode.TextEditorSelectionChangeKind.Keyboard
		) {
			return;
		}
		if (event.kind === vscode.TextEditorSelectionChangeKind.Keyboard) {
			const changed = lastTextChange.get(activeKeywordSnippet.documentKey);
			if (
				changed &&
				changed.serial > activeKeywordSnippet.textChangeSerialAtAccept &&
				Date.now() - changed.at <= KEYBOARD_NAV_TEXT_CHANGE_GRACE_MS
			) {
				return;
			}
		}
		activeKeywordSnippet = undefined;
		void vscode.commands.executeCommand('leaveSnippet');
	};

	context.subscriptions.push(
		vscode.commands.registerCommand(KEYWORD_SNIPPET_ACCEPTED_COMMAND, () => {
			const editor = vscode.window.activeTextEditor;
			if (!editor || !isVbaDocument(editor.document)) {
				return;
			}
			activeKeywordSnippet = {
				editor,
				documentKey: editor.document.uri.toString(),
				textChangeSerialAtAccept: textChangeSerial,
			};
		}),
		vscode.languages.registerCompletionItemProvider(
			selector,
			provider,
			'.',
			' ',
			'#',
			'@',
		),
		vscode.workspace.onDidChangeTextDocument((event) => {
			markTextChange(event.document);
			canonicalCase.handleTextDocumentChange(event);
		}),
		vscode.window.onDidChangeTextEditorSelection((event) => {
			maybeLeaveKeywordSnippet(event);
			canonicalCase.handleSelectionChange(event);
		}),
		vscode.window.onDidChangeActiveTextEditor((editor) => {
			if (activeKeywordSnippet && editor !== activeKeywordSnippet.editor) {
				activeKeywordSnippet = undefined;
			}
			canonicalCase.handleActiveEditorChange(editor);
		}),
		vscode.window.onDidChangeWindowState((state) => {
			canonicalCase.handleWindowStateChange(state);
		}),
		vscode.languages.registerHoverProvider(selector, hoverSignature),
		vscode.languages.registerSignatureHelpProvider(
			selector,
			hoverSignature,
			'(',
			',',
			' ',
		),
		vscode.workspace.onDidSaveTextDocument((doc) => {
			if (doc.uri.scheme !== XLIDE_SCHEME) {
				return;
			}
			try {
				const { xlsmPath } = decodeModuleUri(doc.uri);
				provider.invalidate(xlsmPath);
			} catch {
				// Ignore URIs we cannot decode.
			}
		}),
		vscode.workspace.onDidCloseTextDocument((doc) => canonicalCase.handleDocumentClose(doc)),
	);
}

