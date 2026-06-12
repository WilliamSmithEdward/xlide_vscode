// Registration wiring for the VBA completion stack.
//
// Constructs the per-window editor project-context service and wires the
// completion provider + keyword-snippet tracker (vbaCompletionProvider.ts),
// hover/signature provider (vbaHoverSignatureProvider.ts), and canonical-case
// controller (vbaCanonicalCaseController.ts) into VS Code, plus the
// workbook-save cache invalidation.

import * as vscode from 'vscode';
import {
	XLIDE_SCHEME,
	decodeModuleUri,
} from './xlideFileSystem';
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

