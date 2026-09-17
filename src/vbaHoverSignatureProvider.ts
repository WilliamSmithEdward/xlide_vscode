// Hover and signature-help provider for VBA, resolving against the shared
// editor project-context service with budgeted full-project fallbacks.
//
// Extracted verbatim from vbaMemberCompletion.ts (audit #27).

import * as vscode from 'vscode';
import { XLIDE_SCHEME } from './xlideFileSystem';
import {
	DocRegistry,
	HoverContext,
	resolveHover,
	resolveSignatureHelp,
	SignatureHelpContext,
} from './analyzer';
import { startPerformanceTrace } from './performanceTrace';
import {
	VbaEditorProjectContextService,
	toMemberCompletionContext,
	type EditorProjectContext,
} from './vbaEditorProjectContext';

const HOVER_PROJECT_CONTEXT_BUDGET_MS = 120;
const SIGNATURE_HELP_PROJECT_CONTEXT_BUDGET_MS = 150;

export class VbaHoverSignatureProvider
	implements vscode.HoverProvider, vscode.SignatureHelpProvider
{
	constructor(
		private readonly _projectContext: VbaEditorProjectContextService,
		private readonly _docs?: DocRegistry,
	) {}

	async provideHover(
		document: vscode.TextDocument,
		position: vscode.Position,
		token?: vscode.CancellationToken,
	): Promise<vscode.Hover | undefined> {
		const trace = startPerformanceTrace('hover', document.uri.scheme);
		try {
			return await this._provideHover(document, position, token);
		} finally {
			trace.end(token?.isCancellationRequested ? 'canceled' : 'ok', document.uri.scheme);
		}
	}

	private async _provideHover(
		document: vscode.TextDocument,
		position: vscode.Position,
		token?: vscode.CancellationToken,
	): Promise<vscode.Hover | undefined> {
		if (token?.isCancellationRequested) {
			return undefined;
		}
		const source = document.getText();
		const offset = document.offsetAt(position);
		const info = await this._resolveWithProjectContext(
			document,
			source,
			HOVER_PROJECT_CONTEXT_BUDGET_MS,
			token,
			(ctx) => resolveHover(source, offset, this._hoverContext(ctx)),
		);
		if (!info) {
			return undefined;
		}
		return this._toHover(info, document);
	}

	async provideSignatureHelp(
		document: vscode.TextDocument,
		position: vscode.Position,
		token?: vscode.CancellationToken,
	): Promise<vscode.SignatureHelp | undefined> {
		const trace = startPerformanceTrace('signatureHelp', document.uri.scheme);
		try {
			return await this._provideSignatureHelp(document, position, token);
		} finally {
			trace.end(token?.isCancellationRequested ? 'canceled' : 'ok', document.uri.scheme);
		}
	}

	private async _provideSignatureHelp(
		document: vscode.TextDocument,
		position: vscode.Position,
		token?: vscode.CancellationToken,
	): Promise<vscode.SignatureHelp | undefined> {
		if (token?.isCancellationRequested) {
			return undefined;
		}
		const source = document.getText();
		const offset = document.offsetAt(position);
		const info = await this._resolveWithProjectContext(
			document,
			source,
			SIGNATURE_HELP_PROJECT_CONTEXT_BUDGET_MS,
			token,
			(ctx) => resolveSignatureHelp(source, offset, this._signatureHelpContext(ctx, source)),
		);
		if (!info) {
			return undefined;
		}
		const sig = new vscode.SignatureInformation(info.label);
		sig.parameters = info.parameters.map((p) => {
			const pi = new vscode.ParameterInformation(p.label);
			if (p.documentation) {
				pi.documentation = new vscode.MarkdownString(p.documentation);
			}
			return pi;
		});
		if (info.documentation) {
			sig.documentation = new vscode.MarkdownString(info.documentation);
		}
		const help = new vscode.SignatureHelp();
		help.signatures = [sig];
		help.activeSignature = 0;
		help.activeParameter = info.activeParameter;
		return help;
	}

	private _toHover(
		info: NonNullable<ReturnType<typeof resolveHover>>,
		document: vscode.TextDocument,
	): vscode.Hover {
		const md = new vscode.MarkdownString();
		md.appendCodeblock(info.signature, 'vba');
		if (info.documentation) {
			md.appendMarkdown('\n\n');
			md.appendMarkdown(info.documentation);
		}
		if (info.details.length > 0) {
			md.appendMarkdown('\n\n');
			md.appendMarkdown(info.details.join('  \n'));
		}
		const range = new vscode.Range(
			document.positionAt(info.span.start),
			document.positionAt(info.span.end),
		);
		return new vscode.Hover(md, range);
	}

	/**
	 * Resolves against the cheapest project context that answers: the cached
	 * one when there is one, else the module's own text, then a project
	 * context built within the budget. Undefined once the request is stale.
	 */
	private async _resolveWithProjectContext<T>(
		document: vscode.TextDocument,
		source: string,
		budgetMs: number,
		token: vscode.CancellationToken | undefined,
		resolve: (ctx: EditorProjectContext) => T | undefined,
	): Promise<T | undefined> {
		const requestVersion = document.version;
		const cached = this._projectContext.cachedEditorProjectContext(document);
		const fastCtx = cached ?? this._projectContext.cheapEditorProjectContext(document);
		if (!cached && document.uri.scheme === XLIDE_SCHEME) {
			this._projectContext.warmEditorProjectContext(document, source);
		}
		let info = resolve(fastCtx);
		if (!info && !cached) {
			info = resolve(this._projectContext.localEditorProjectContext(document, source));
		}
		if (!info && !cached && document.uri.scheme === XLIDE_SCHEME) {
			const projectCtx = await this._projectContext.buildEditorProjectContextWithin(document, source, budgetMs);
			if (token?.isCancellationRequested || document.version !== requestVersion) {
				return undefined;
			}
			if (projectCtx) {
				info = resolve(projectCtx);
			}
		}
		return info;
	}

	private _hoverContext(ctx: EditorProjectContext): HoverContext {
		return {
			...toMemberCompletionContext(ctx),
			moduleName: ctx.moduleName,
			moduleKind: ctx.moduleKind,
			projectTypes: ctx.projectTypes,
			projectProcedures: ctx.projectProcedures,
			docRegistry: this._docs,
		};
	}

	private _signatureHelpContext(
		ctx: EditorProjectContext,
		source: string,
	): SignatureHelpContext {
		return {
			...toMemberCompletionContext(ctx),
			moduleName: ctx.moduleName,
			moduleSource: source,
			projectProcedures: ctx.projectProcedures,
			docRegistry: this._docs,
		};
	}
}
