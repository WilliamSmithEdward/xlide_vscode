// Registration wiring for the VBA completion stack.
//
// Constructs the per-window editor project-context service and wires the
// completion provider + keyword-snippet tracker (vbaCompletionProvider.ts),
// hover/signature provider (vbaHoverSignatureProvider.ts), and canonical-case
// controller (vbaCanonicalCaseController.ts) into VS Code, plus the
// workbook-save cache invalidation.

import * as vscode from 'vscode';
import { moduleLocationOfDocument } from './vbaDocumentLocation';
import { DocRegistry } from './analyzer';
import { VbaProjectIndexService } from './vbaProjectIndexService';
import { VbaCanonicalCaseController } from './vbaCanonicalCaseController';
import { VbaEditorProjectContextService } from './vbaEditorProjectContext';
import { VbaHoverSignatureProvider } from './vbaHoverSignatureProvider';
import {
	KEYWORD_SNIPPET_ACCEPTED_COMMAND,
	VbaKeywordSnippetTracker,
	VbaMemberCompletionProvider,
} from './vbaCompletionProvider';

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
	const keywordSnippets = new VbaKeywordSnippetTracker();
	ACTIVE_MEMBER_COMPLETION_PROVIDERS.add(provider);
	context.subscriptions.push({
		dispose: () => ACTIVE_MEMBER_COMPLETION_PROVIDERS.delete(provider),
	});

	context.subscriptions.push(
		vscode.commands.registerCommand(
			KEYWORD_SNIPPET_ACCEPTED_COMMAND,
			() => keywordSnippets.handleSnippetAccepted(),
		),
		vscode.languages.registerCompletionItemProvider(
			selector,
			provider,
			'.',
			' ',
			'#',
			'@',
		),
		vscode.workspace.onDidChangeTextDocument((event) => {
			keywordSnippets.handleTextDocumentChange(event);
			canonicalCase.handleTextDocumentChange(event);
			// Drop the workbook's derived editor-context cache when ANY of its
			// modules is edited (even unsaved), so completion/hover for one module
			// does not serve stale cross-module symbols from a sibling module's
			// snapshot within the cache TTL. Editing the active module already
			// cache-misses via its own version bump, so the extra cost falls only
			// on the sibling-edit case this fixes.
			const location = moduleLocationOfDocument(event.document);
			if (location) {
				provider.invalidate(location.xlsmPath);
			}
		}),
		vscode.window.onDidChangeTextEditorSelection((event) => {
			keywordSnippets.handleSelectionChange(event);
			canonicalCase.handleSelectionChange(event);
		}),
		vscode.window.onDidChangeActiveTextEditor((editor) => {
			keywordSnippets.handleActiveEditorChange(editor);
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
			const location = moduleLocationOfDocument(doc);
			if (location) {
				provider.invalidate(location.xlsmPath);
			}
		}),
		vscode.workspace.onDidCloseTextDocument((doc) => { keywordSnippets.handleDocumentClose(doc); canonicalCase.handleDocumentClose(doc); }),
	);
}

