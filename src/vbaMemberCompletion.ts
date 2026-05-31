// Host-context member completion provider.
//
// Registers a CompletionItemProvider triggered on '.' that resolves the type of
// the receiver expression (ThisWorkbook, Application, ActiveSheet, a worksheet
// code name, Me, or a typed local/module variable) and offers the verified
// Excel object-model members for that type. The same provider also offers
// type-name completions in a declaration type position (after `As` / `As New`)
// and a HoverProvider that describes the identifier under the cursor. See the
// Host-Context Member Completion addendum and Phases 6-7 in
// docs/xlide_vba_language_service_roadmap.md.

import * as vscode from 'vscode';
import { PythonBridge } from './pythonBridge';
import { XLIDE_SCHEME, decodeModuleUri } from './xlideFileSystem';
import {
	buildModuleSymbols,
	DocRegistry,
	getHostType,
	HoverContext,
	IdentifierCompletion,
	IdentifierCompletionContext,
	MemberCompletion,
	MemberCompletionContext,
	ModuleSymbolKind,
	ProjectTypeName,
	resolveHover,
	resolveIdentifierCompletions,
	resolveMemberCompletions,
	resolveSignatureHelp,
	resolveTypeCompletions,
	SignatureHelpContext,
	TypeCompletion,
	TypeCompletionContext,
} from './analyzer';

const WORKBOOK = 'Excel.Workbook';
const WORKSHEET = 'Excel.Worksheet';
const MODULE_CACHE_TTL_MS = 5000;

interface ModuleEntry {
	name: string;
	type: string;
}

interface CachedModules {
	entries: ModuleEntry[];
	loadedAt: number;
}

interface CachedProjectTypes {
	types: ProjectTypeName[];
	loadedAt: number;
}

/** Maps a document module to the host type that `Me` denotes inside it. */
function meTypeFor(entry: ModuleEntry | undefined): string | undefined {
	if (!entry || entry.type !== 'document') {
		return undefined;
	}
	return entry.name.toLowerCase() === 'thisworkbook' ? WORKBOOK : WORKSHEET;
}

/** Builds the lowercased code-name -> host type map for a workbook project. */
function codeNamesFor(entries: ModuleEntry[]): Record<string, string> {
	const out: Record<string, string> = {};
	for (const entry of entries) {
		if (entry.type !== 'document') {
			continue;
		}
		out[entry.name.toLowerCase()] =
			entry.name.toLowerCase() === 'thisworkbook' ? WORKBOOK : WORKSHEET;
	}
	return out;
}

class VbaMemberCompletionProvider
	implements
		vscode.CompletionItemProvider,
		vscode.HoverProvider,
		vscode.SignatureHelpProvider
{
	private readonly _cache = new Map<string, CachedModules>();
	private readonly _typeCache = new Map<string, CachedProjectTypes>();

	constructor(
		private readonly _bridge: PythonBridge,
		private readonly _docs?: DocRegistry,
	) {}

	/** Drop cached module lists for a workbook (e.g. after a project change). */
	invalidate(xlsmPath?: string): void {
		if (xlsmPath === undefined) {
			this._cache.clear();
			this._typeCache.clear();
		} else {
			const key = this._key(xlsmPath);
			this._cache.delete(key);
			this._typeCache.delete(key);
		}
	}

	async provideCompletionItems(
		document: vscode.TextDocument,
		position: vscode.Position,
	): Promise<vscode.CompletionItem[]> {
		const source = document.getText();
		const offset = document.offsetAt(position);

		const memberCtx = await this._buildContext(document);
		const members = resolveMemberCompletions(source, offset, memberCtx);
		if (members.length > 0) {
			return members.map((mem) => this._toItem(mem));
		}

		const typeCtx = await this._buildTypeContext(document, source);
		const types = resolveTypeCompletions(source, offset, typeCtx);
		if (types.length > 0) {
			return types.map((t) => this._toTypeItem(t));
		}

		const identCtx = await this._buildIdentifierContext(document);
		const idents = resolveIdentifierCompletions(source, offset, identCtx);
		return idents.map((id) => this._toIdentItem(id));
	}

	async provideHover(
		document: vscode.TextDocument,
		position: vscode.Position,
	): Promise<vscode.Hover | undefined> {
		const source = document.getText();
		const offset = document.offsetAt(position);
		const info = resolveHover(source, offset, await this._buildHoverContext(document));
		if (!info) {
			return undefined;
		}
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

	async provideSignatureHelp(
		document: vscode.TextDocument,
		position: vscode.Position,
	): Promise<vscode.SignatureHelp | undefined> {
		const source = document.getText();
		const offset = document.offsetAt(position);
		const ctx: SignatureHelpContext = {
			...(await this._buildContext(document)),
			moduleSource: source,
			docRegistry: this._docs,
		};
		const info = resolveSignatureHelp(source, offset, ctx);
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

	private async _buildHoverContext(
		document: vscode.TextDocument,
	): Promise<HoverContext> {
		if (document.uri.scheme !== XLIDE_SCHEME) {
			return { docRegistry: this._docs };
		}
		try {
			const decoded = decodeModuleUri(document.uri);
			const entries = await this._loadModules(decoded.xlsmPath);
			const current = entries?.find(
				(e) => e.name.toLowerCase() === decoded.moduleName.toLowerCase(),
			);
			return {
				codeNames: codeNamesFor(entries ?? []),
				meType: meTypeFor(current),
				moduleName: decoded.moduleName,
				moduleKind: current ? this._moduleKind(current.type) : 'standard',
				docRegistry: this._docs,
			};
		} catch {
			return { docRegistry: this._docs };
		}
	}

	private async _buildContext(
		document: vscode.TextDocument,
	): Promise<MemberCompletionContext> {
		if (document.uri.scheme !== XLIDE_SCHEME) {
			return {};
		}
		let decoded: { xlsmPath: string; moduleName: string };
		try {
			decoded = decodeModuleUri(document.uri);
		} catch {
			return {};
		}
		const entries = await this._loadModules(decoded.xlsmPath);
		if (!entries) {
			return {};
		}
		const current = entries.find(
			(e) => e.name.toLowerCase() === decoded.moduleName.toLowerCase(),
		);
		return {
			codeNames: codeNamesFor(entries),
			meType: meTypeFor(current),
		};
	}

	private async _loadModules(xlsmPath: string): Promise<ModuleEntry[] | undefined> {
		const key = this._key(xlsmPath);
		const cached = this._cache.get(key);
		if (cached && Date.now() - cached.loadedAt < MODULE_CACHE_TTL_MS) {
			return cached.entries;
		}
		try {
			const entries = await this._bridge.call<ModuleEntry[]>('listModules', {
				path: xlsmPath,
			});
			this._cache.set(key, { entries, loadedAt: Date.now() });
			return entries;
		} catch {
			return cached?.entries;
		}
	}

	/**
	 * Builds the type-completion context: user-defined types/enums declared in the
	 * current module plus class/UserForm module names from the workbook project.
	 */
	private async _buildTypeContext(
		document: vscode.TextDocument,
		source: string,
	): Promise<TypeCompletionContext> {
		const projectTypes: ProjectTypeName[] = [];
		const seen = new Set<string>();
		const push = (name: string, kind: ProjectTypeName['kind']): void => {
			const key = name.toLowerCase();
			if (name && !seen.has(key)) {
				seen.add(key);
				projectTypes.push({ name, kind });
			}
		};

		let currentName = 'Module';
		let currentKind: ModuleSymbolKind = 'standard';
		if (document.uri.scheme === XLIDE_SCHEME) {
			try {
				const decoded = decodeModuleUri(document.uri);
				currentName = decoded.moduleName;
				const entries = await this._loadModules(decoded.xlsmPath);
				for (const entry of entries ?? []) {
					if (entry.type === 'class' || entry.type === 'userform') {
						push(entry.name, 'class');
					}
					if (entry.name.toLowerCase() === currentName.toLowerCase()) {
						currentKind = this._moduleKind(entry.type);
					}
				}
				// Public Type/Enum declared in OTHER modules of the project.
				const crossModule = await this._loadCrossModuleTypes(
					decoded.xlsmPath,
					entries ?? [],
					currentName.toLowerCase(),
				);
				for (const t of crossModule) {
					push(t.name, t.kind);
				}
			} catch {
				// Fall back to a primitives + host-types only context.
			}
		}

		// User-defined Type/Enum declarations in the module currently being edited.
		// Read live from the editor so unsaved declarations are offered too.
		try {
			const mod = buildModuleSymbols(currentName, currentKind, source);
			for (const child of mod.root.children ?? []) {
				if (child.kind === 'type') {
					push(child.name, 'userType');
				} else if (child.kind === 'enum') {
					push(child.name, 'enum');
				}
			}
		} catch {
			// Ignore parse failures; primitives + host types still apply.
		}

		return { projectTypes };
	}

	/**
	 * Collects Public (non-`Private`) `Type` and `Enum` names declared across the
	 * workbook's modules, excluding the current module (handled live). Reads each
	 * module's saved source via the bridge and caches the result per workbook.
	 */
	private async _loadCrossModuleTypes(
		xlsmPath: string,
		entries: ModuleEntry[],
		currentLower: string,
	): Promise<ProjectTypeName[]> {
		const key = this._key(xlsmPath);
		const cached = this._typeCache.get(key);
		if (cached && Date.now() - cached.loadedAt < MODULE_CACHE_TTL_MS) {
			return cached.types;
		}

		const types: ProjectTypeName[] = [];
		for (const entry of entries) {
			if (entry.name.toLowerCase() === currentLower) {
				continue; // Current module is resolved from the live document.
			}
			let source: string;
			try {
				const res = await this._bridge.call<{ source: string }>('readModule', {
					path: xlsmPath,
					module: entry.name,
				});
				source = res.source;
			} catch {
				continue; // Skip modules we cannot read.
			}
			try {
				const mod = buildModuleSymbols(
					entry.name,
					this._moduleKind(entry.type),
					source,
				);
				for (const child of mod.root.children ?? []) {
					if (child.visibility === 'Private') {
						continue;
					}
					if (child.kind === 'type') {
						types.push({ name: child.name, kind: 'userType' });
					} else if (child.kind === 'enum') {
						types.push({ name: child.name, kind: 'enum' });
					}
				}
			} catch {
				// Ignore parse failures for individual modules.
			}
		}

		this._typeCache.set(key, { types, loadedAt: Date.now() });
		return types;
	}

	/**
	 * Builds the identifier-completion context: worksheet/document code names of
	 * the workbook project plus the module currently being edited.
	 */
	private async _buildIdentifierContext(
		document: vscode.TextDocument,
	): Promise<IdentifierCompletionContext> {
		if (document.uri.scheme !== XLIDE_SCHEME) {
			return {};
		}
		try {
			const decoded = decodeModuleUri(document.uri);
			const entries = await this._loadModules(decoded.xlsmPath);
			const codeNames: string[] = [];
			let moduleKind: ModuleSymbolKind = 'standard';
			for (const entry of entries ?? []) {
				if (entry.type === 'document') {
					codeNames.push(entry.name);
				}
				if (entry.name.toLowerCase() === decoded.moduleName.toLowerCase()) {
					moduleKind = this._moduleKind(entry.type);
				}
			}
			return { codeNames, moduleName: decoded.moduleName, moduleKind };
		} catch {
			return {};
		}
	}

	/** Maps a host module-type string to a {@link ModuleSymbolKind}. */
	private _moduleKind(type: string): ModuleSymbolKind {
		switch (type) {
			case 'class':
				return 'class';
			case 'document':
				return 'document';
			case 'userform':
				return 'userform';
			default:
				return 'standard';
		}
	}

	private _toItem(mem: MemberCompletion): vscode.CompletionItem {
		const item = new vscode.CompletionItem(
			mem.name,
			mem.kind === 'method'
				? vscode.CompletionItemKind.Method
				: vscode.CompletionItemKind.Property,
		);
		const ownerName = getHostType(mem.owner)?.displayName ?? mem.owner;
		const kindLabel = mem.kind === 'method' ? 'method' : 'property';
		item.detail = `${ownerName} ${kindLabel}`;
		if (mem.returns) {
			const returnName = getHostType(mem.returns)?.displayName ?? mem.returns;
			item.detail += ` -> ${returnName}`;
		}
		return item;
	}

	private _toTypeItem(t: TypeCompletion): vscode.CompletionItem {
		const item = new vscode.CompletionItem(t.name, this._typeItemKind(t));
		item.detail = t.detail;
		return item;
	}

	private _typeItemKind(t: TypeCompletion): vscode.CompletionItemKind {
		switch (t.kind) {
			case 'enum':
				return vscode.CompletionItemKind.Enum;
			case 'host':
			case 'class':
				return vscode.CompletionItemKind.Class;
			default:
				return vscode.CompletionItemKind.Struct;
		}
	}

	private _toIdentItem(id: IdentifierCompletion): vscode.CompletionItem {
		const item = new vscode.CompletionItem(id.name, this._identItemKind(id));
		item.detail = id.detail;
		return item;
	}

	private _identItemKind(id: IdentifierCompletion): vscode.CompletionItemKind {
		switch (id.kind) {
			case 'procedure':
				return vscode.CompletionItemKind.Method;
			case 'runtime':
				return vscode.CompletionItemKind.Function;
			case 'constant':
				return vscode.CompletionItemKind.Constant;
			case 'enum':
				return vscode.CompletionItemKind.Enum;
			case 'enumMember':
				return vscode.CompletionItemKind.EnumMember;
			case 'type':
				return vscode.CompletionItemKind.Struct;
			case 'global':
			case 'codeName':
				return vscode.CompletionItemKind.Variable;
			default:
				return vscode.CompletionItemKind.Variable;
		}
	}

	private _key(xlsmPath: string): string {
		return process.platform === 'win32' ? xlsmPath.toLowerCase() : xlsmPath;
	}
}

/** Registers the host-context member completion provider. */
export function registerVbaMemberCompletion(
	context: vscode.ExtensionContext,
	bridge: PythonBridge,
	selector: vscode.DocumentSelector,
	docs?: DocRegistry,
): void {
	const provider = new VbaMemberCompletionProvider(bridge, docs);
	context.subscriptions.push(
		vscode.languages.registerCompletionItemProvider(
			selector,
			provider,
			'.',
			' ',
		),
		vscode.languages.registerHoverProvider(selector, provider),
		vscode.languages.registerSignatureHelpProvider(
			selector,
			provider,
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
	);
}
