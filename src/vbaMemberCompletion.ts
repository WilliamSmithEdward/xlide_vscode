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
	EventHandlerCompletion,
	EventHandlerCompletionContext,
	EventHandlerDocumentType,
	eventHandlerDocumentTypeForContext,
	getHostType,
	HoverContext,
	IdentifierCompletion,
	IdentifierCompletionContext,
	MemberCompletion,
	MemberCompletionContext,
	ModuleSymbolKind,
	ProjectIndex,
	ProjectTypeName,
	VbaProcedureSignature,
	callableCompletionShouldInsertParens,
	resolveCanonicalCaseEdit,
	resolveEventHandlerCompletions,
	resolveHover,
	resolveIdentifierCompletions,
	resolveMemberCompletions,
	resolveSignatureHelp,
	resolveTypeCompletions,
	SignatureHelpContext,
	TypeCompletion,
	TypeCompletionContext,
	VbaProjectClassMembers,
} from './analyzer';

const WORKBOOK = 'Excel.Workbook';
const WORKSHEET = 'Excel.Worksheet';
const CHART = 'Excel.Chart';
const MODULE_CACHE_TTL_MS = 5000;

interface ModuleEntry {
	name: string;
	type: string;
	documentType?: EventHandlerDocumentType;
}

interface CachedModules {
	entries: ModuleEntry[];
	loadedAt: number;
}

interface CachedProjectTypes {
	types: ProjectTypeName[];
	loadedAt: number;
}

interface CachedProjectProcedures {
	procedures: VbaProcedureSignature[];
	loadedAt: number;
}

/** Maps a document module to the host type that `Me` denotes inside it. */
function meTypeFor(entry: ModuleEntry | undefined): string | undefined {
	if (!entry || entry.type !== 'document') {
		return undefined;
	}
	switch (documentTypeFor(entry)) {
		case 'workbook':
			return WORKBOOK;
		case 'chart':
			return CHART;
		default:
			return WORKSHEET;
	}
}

/** Maps `Me` to the source-backed current object module when applicable. */
function meProjectTypeFor(entry: ModuleEntry | undefined): string | undefined {
	if (!entry || !['class', 'document', 'userform'].includes(entry.type)) {
		return undefined;
	}
	return entry.name;
}

/** Builds the lowercased code-name -> host type map for a workbook project. */
function codeNamesFor(entries: ModuleEntry[]): Record<string, string> {
	const out: Record<string, string> = {};
	for (const entry of entries) {
		if (entry.type !== 'document') {
			continue;
		}
		out[entry.name.toLowerCase()] = meTypeFor(entry) ?? WORKSHEET;
	}
	return out;
}

function documentTypeFor(entry: ModuleEntry | undefined): EventHandlerDocumentType | undefined {
	if (!entry || entry.type !== 'document') {
		return undefined;
	}
	return eventHandlerDocumentTypeForContext({
		moduleName: entry.name,
		moduleKind: 'document',
		documentType: entry.documentType,
	});
}

class VbaMemberCompletionProvider
	implements
		vscode.CompletionItemProvider,
		vscode.HoverProvider,
		vscode.SignatureHelpProvider
{
	private readonly _cache = new Map<string, CachedModules>();
	private readonly _typeCache = new Map<string, CachedProjectTypes>();
	private readonly _procedureCache = new Map<string, CachedProjectProcedures>();
	private _applyingCanonicalCase = false;

	constructor(
		private readonly _bridge: PythonBridge,
		private readonly _docs?: DocRegistry,
	) {}

	/** Drop cached module lists for a workbook (e.g. after a project change). */
	invalidate(xlsmPath?: string): void {
		if (xlsmPath === undefined) {
			this._cache.clear();
			this._typeCache.clear();
			this._procedureCache.clear();
		} else {
			const key = this._key(xlsmPath);
			this._cache.delete(key);
			this._typeCache.delete(key);
			this._procedureCache.delete(key);
		}
	}

	async provideCompletionItems(
		document: vscode.TextDocument,
		position: vscode.Position,
	): Promise<vscode.CompletionItem[]> {
		const source = document.getText();
		const offset = document.offsetAt(position);
		const range = this._completionRange(document, position);

		const memberCtx = await this._buildContext(document);
		const members = resolveMemberCompletions(source, offset, memberCtx);
		if (members.length > 0) {
			return members.map((mem) => this._toItem(mem, range, source, offset));
		}

		const typeCtx = await this._buildTypeContext(document, source);
		const types = resolveTypeCompletions(source, offset, typeCtx);
		if (types.length > 0) {
			return types.map((t) => this._toTypeItem(t, range));
		}

		const eventCtx = await this._buildEventHandlerContext(document);
		const events = resolveEventHandlerCompletions(source, offset, eventCtx);
		if (events.length > 0) {
			return events.map((event) => this._toEventHandlerItem(event, range));
		}

		const identCtx = await this._buildIdentifierContext(document);
		const idents = resolveIdentifierCompletions(source, offset, identCtx);
		return idents.map((id) => this._toIdentItem(id, range, source, offset));
	}

	async applyCanonicalCase(
		document: vscode.TextDocument,
		candidateEnd: vscode.Position,
		editorHint?: vscode.TextEditor,
	): Promise<void> {
		if (this._applyingCanonicalCase) {
			return;
		}
		this._applyingCanonicalCase = true;
		try {
			const editor = editorHint?.document === document
				? editorHint
				: vscode.window.visibleTextEditors.find((candidate) => candidate.document === document);
			if (!editor || editor.document !== document) {
				return;
			}
			const source = document.getText();
			const offset = document.offsetAt(candidateEnd);
			const edit = resolveCanonicalCaseEdit(source, offset, {
				member: await this._buildContext(document),
				identifier: await this._buildIdentifierContext(document),
				type: await this._buildTypeContext(document, source),
			});
			if (!edit) {
				return;
			}
			const range = new vscode.Range(
				document.positionAt(edit.start),
				document.positionAt(edit.end),
			);
			if (document.getText(range) === edit.text) {
				return;
			}
			await editor.edit((builder) => builder.replace(range, edit.text), {
				undoStopBefore: false,
				undoStopAfter: false,
			});
		} finally {
			this._applyingCanonicalCase = false;
		}
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
			projectProcedures: await this._buildProjectProcedureContext(document),
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
		const source = document.getText();
		const typeContext = await this._buildTypeContext(document, source);
		if (document.uri.scheme !== XLIDE_SCHEME) {
			return { ...typeContext, docRegistry: this._docs };
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
				meProjectType: meProjectTypeFor(current),
				moduleName: decoded.moduleName,
				moduleKind: current ? this._moduleKind(current.type) : 'standard',
				projectTypes: typeContext.projectTypes,
				projectClassMembers: entries
					? await this._loadProjectClassMembers(decoded.xlsmPath, entries)
					: undefined,
				projectProcedures: entries
					? await this._loadCrossModuleProcedureSignatures(
						decoded.xlsmPath,
						entries,
						decoded.moduleName.toLowerCase(),
					)
					: undefined,
				docRegistry: this._docs,
			};
		} catch {
			return { ...typeContext, docRegistry: this._docs };
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
		const projectClassMembers = await this._loadProjectClassMembers(
			decoded.xlsmPath,
			entries,
		);
		return {
			codeNames: codeNamesFor(entries),
			meType: meTypeFor(current),
			meProjectType: meProjectTypeFor(current),
			projectClassMembers,
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

	private async _loadProjectClassMembers(
		xlsmPath: string,
		entries: ModuleEntry[],
	): Promise<VbaProjectClassMembers[]> {
		const project = new ProjectIndex();
		for (const entry of entries) {
			const kind = this._moduleKind(entry.type);
			if (kind !== 'class' && kind !== 'userform' && kind !== 'document') {
				continue;
			}
			const source = await this._moduleSource(xlsmPath, entry);
			if (source === undefined) {
				continue;
			}
			project.setModule({
				moduleName: entry.name,
				moduleKind: kind,
				source,
			});
		}
		return project.projectClassMembers();
	}

	private async _loadCrossModuleProcedureSignatures(
		xlsmPath: string,
		entries: ModuleEntry[],
		currentLower: string,
	): Promise<VbaProcedureSignature[]> {
		const allProcedures = await this._loadProjectProcedureSignatures(xlsmPath, entries);
		return allProcedures.filter(
			(procedure) => procedure.moduleName.toLowerCase() !== currentLower,
		);
	}

	private async _loadProjectProcedureSignatures(
		xlsmPath: string,
		entries: ModuleEntry[],
	): Promise<VbaProcedureSignature[]> {
		const key = this._key(xlsmPath);
		const cached = this._procedureCache.get(key);
		if (cached && Date.now() - cached.loadedAt < MODULE_CACHE_TTL_MS) {
			return cached.procedures;
		}
		const project = new ProjectIndex();
		for (const entry of entries) {
			const kind = this._moduleKind(entry.type);
			if (kind !== 'standard') {
				continue;
			}
			const source = await this._moduleSource(xlsmPath, entry);
			if (source === undefined) {
				continue;
			}
			project.setModule({
				moduleName: entry.name,
				moduleKind: kind,
				source,
			});
		}
		const procedures = project.visibleProcedureSignatures('__xlide_external_module__');
		this._procedureCache.set(key, { procedures, loadedAt: Date.now() });
		return procedures;
	}

	private async _moduleSource(
		xlsmPath: string,
		entry: ModuleEntry,
	): Promise<string | undefined> {
		const open = this._openModuleSource(xlsmPath, entry.name);
		if (open !== undefined) {
			return open;
		}
		try {
			const res = await this._bridge.call<{ source: string }>('readModule', {
				path: xlsmPath,
				module: entry.name,
			});
			return res.source;
		} catch {
			return undefined;
		}
	}

	private _openModuleSource(
		xlsmPath: string,
		moduleName: string,
	): string | undefined {
		const workbookKey = this._key(xlsmPath);
		const moduleKey = moduleName.toLowerCase();
		for (const doc of vscode.workspace.textDocuments) {
			if (doc.uri.scheme !== XLIDE_SCHEME) {
				continue;
			}
			try {
				const decoded = decodeModuleUri(doc.uri);
				if (
					this._key(decoded.xlsmPath) === workbookKey &&
					decoded.moduleName.toLowerCase() === moduleKey
				) {
					return doc.getText();
				}
			} catch {
				// Ignore virtual documents outside XLIDE's module URI shape.
			}
		}
		return undefined;
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
		const push = (
			name: string,
			kind: ProjectTypeName['kind'],
			doc?: ProjectTypeName['doc'],
		): void => {
			const key = name.toLowerCase();
			if (!name) {
				return;
			}
			if (!seen.has(key)) {
				seen.add(key);
				projectTypes.push({ name, kind, doc });
				return;
			}
			if (doc) {
				const existing = projectTypes.find((item) => item.name.toLowerCase() === key);
				if (existing && !existing.doc) {
					existing.doc = doc;
				}
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
					const kind = this._moduleKind(entry.type);
					if (kind === 'class' || kind === 'document' || kind === 'userform') {
						push(entry.name, kind);
					}
					if (entry.name.toLowerCase() === currentName.toLowerCase()) {
						currentKind = kind;
					}
				}
				// Public Type/Enum declared in OTHER modules of the project.
				const crossModule = await this._loadCrossModuleTypes(
					decoded.xlsmPath,
					entries ?? [],
					currentName.toLowerCase(),
				);
				for (const t of crossModule) {
					push(t.name, t.kind, t.doc);
				}
			} catch {
				// Fall back to a primitives + host-types only context.
			}
		}

		// User-defined Type/Enum declarations in the module currently being edited.
		// Read live from the editor so unsaved declarations are offered too.
		try {
			const mod = buildModuleSymbols(currentName, currentKind, source);
			if (currentKind === 'class' || currentKind === 'document' || currentKind === 'userform') {
				push(currentName, currentKind, mod.root.doc);
			}
			for (const child of mod.root.children ?? []) {
				if (child.kind === 'type') {
					push(child.name, 'userType', child.doc);
				} else if (child.kind === 'enum') {
					push(child.name, 'enum', child.doc);
				}
			}
		} catch {
			// Ignore parse failures; primitives + host types still apply.
		}

		return { projectTypes };
	}

	/**
	 * Collects workbook object-module docs plus Public (non-`Private`) `Type`
	 * and `Enum` names declared across the workbook's modules, excluding the
	 * current module (handled live). Reads each module's source via the live
	 * editor/bridge path and caches the result per workbook.
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
			const source = await this._moduleSource(xlsmPath, entry);
			if (source === undefined) {
				continue; // Skip modules we cannot read.
			}
			try {
				const kind = this._moduleKind(entry.type);
				const mod = buildModuleSymbols(entry.name, kind, source);
				if (kind === 'class' || kind === 'document' || kind === 'userform') {
					types.push({ name: entry.name, kind, doc: mod.root.doc });
				}
				for (const child of mod.root.children ?? []) {
					if (child.visibility === 'Private') {
						continue;
					}
					if (child.kind === 'type') {
						types.push({ name: child.name, kind: 'userType', doc: child.doc });
					} else if (child.kind === 'enum') {
						types.push({ name: child.name, kind: 'enum', doc: child.doc });
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
	 * Builds the event-handler completion context for the current module.
	 * Event procedures are scoped by host document module type and intentionally
	 * stay outside object-member completion.
	 */
	private async _buildEventHandlerContext(
		document: vscode.TextDocument,
	): Promise<EventHandlerCompletionContext> {
		if (document.uri.scheme !== XLIDE_SCHEME) {
			return {};
		}
		try {
			const decoded = decodeModuleUri(document.uri);
			const entries = await this._loadModules(decoded.xlsmPath);
			const current = entries?.find(
				(e) => e.name.toLowerCase() === decoded.moduleName.toLowerCase(),
			);
			return {
				moduleName: decoded.moduleName,
				moduleKind: current ? this._moduleKind(current.type) : 'standard',
				documentType: documentTypeFor(current),
			};
		} catch {
			return {};
		}
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
			const currentLower = decoded.moduleName.toLowerCase();
			for (const entry of entries ?? []) {
				if (entry.type === 'document') {
					codeNames.push(entry.name);
				}
				if (entry.name.toLowerCase() === currentLower) {
					moduleKind = this._moduleKind(entry.type);
				}
			}
			return {
				codeNames,
				moduleName: decoded.moduleName,
				moduleKind,
				projectProcedures: entries
					? await this._loadCrossModuleProcedureSignatures(
						decoded.xlsmPath,
						entries,
						currentLower,
					)
					: undefined,
			};
		} catch {
			return {};
		}
	}

	private async _buildProjectProcedureContext(
		document: vscode.TextDocument,
	): Promise<VbaProcedureSignature[] | undefined> {
		if (document.uri.scheme !== XLIDE_SCHEME) {
			return undefined;
		}
		try {
			const decoded = decodeModuleUri(document.uri);
			const entries = await this._loadModules(decoded.xlsmPath);
			return entries
				? await this._loadCrossModuleProcedureSignatures(
					decoded.xlsmPath,
					entries,
					decoded.moduleName.toLowerCase(),
				)
				: undefined;
		} catch {
			return undefined;
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

	private _toItem(
		mem: MemberCompletion,
		range: vscode.Range,
		source: string,
		offset: number,
	): vscode.CompletionItem {
		const item = new vscode.CompletionItem(mem.name, this._memberItemKind(mem));
		const ownerName = getHostType(mem.owner)?.displayName ?? mem.owner;
		const kindLabel = mem.kind;
		if (mem.signature) {
			item.detail = `${ownerName}.${mem.signature}`;
		} else {
			item.detail = `${ownerName} ${kindLabel}`;
		}
		if (!mem.signature && mem.returns) {
			const returnName = getHostType(mem.returns)?.displayName ?? mem.returns;
			item.detail += ` -> ${returnName}`;
		}
		if (mem.documentation) {
			item.documentation = new vscode.MarkdownString(mem.documentation);
		}
		this._applyCompletionInsert(
			item,
			mem.name,
			range,
			mem.kind === 'method',
			callableCompletionShouldInsertParens(source, offset),
		);
		return item;
	}

	private _memberItemKind(mem: MemberCompletion): vscode.CompletionItemKind {
		switch (mem.kind) {
			case 'method':
				return vscode.CompletionItemKind.Method;
			case 'event':
				return vscode.CompletionItemKind.Event;
			default:
				return vscode.CompletionItemKind.Property;
		}
	}

	private _toTypeItem(t: TypeCompletion, range: vscode.Range): vscode.CompletionItem {
		const item = new vscode.CompletionItem(t.name, this._typeItemKind(t));
		item.detail = t.detail;
		if (t.documentation) {
			item.documentation = new vscode.MarkdownString(t.documentation);
		}
		this._applyCompletionInsert(item, t.name, range, false);
		return item;
	}

	private _typeItemKind(t: TypeCompletion): vscode.CompletionItemKind {
		switch (t.kind) {
			case 'enum':
				return vscode.CompletionItemKind.Enum;
			case 'host':
			case 'class':
			case 'document':
			case 'userform':
				return vscode.CompletionItemKind.Class;
			default:
				return vscode.CompletionItemKind.Struct;
		}
	}

	private _toEventHandlerItem(
		event: EventHandlerCompletion,
		range: vscode.Range,
	): vscode.CompletionItem {
		const item = new vscode.CompletionItem(event.name, vscode.CompletionItemKind.Event);
		item.detail = event.detail;
		const documentation = new vscode.MarkdownString();
		documentation.appendCodeblock(event.signature, 'vba');
		documentation.appendMarkdown('\n\n');
		documentation.appendMarkdown(event.documentation);
		item.documentation = documentation;
		item.range = range;
		item.filterText = event.name;
		item.sortText = `0:${event.name}`;
		item.insertText = new vscode.SnippetString(event.insertText);
		return item;
	}

	private _toIdentItem(
		id: IdentifierCompletion,
		range: vscode.Range,
		source: string,
		offset: number,
	): vscode.CompletionItem {
		const item = new vscode.CompletionItem(id.name, this._identItemKind(id));
		item.detail = id.detail;
		if (id.documentation) {
			item.documentation = new vscode.MarkdownString(id.documentation);
		}
		const callable = id.kind === 'runtime' || id.kind === 'procedure';
		this._applyCompletionInsert(
			item,
			id.name,
			range,
			callable,
			callableCompletionShouldInsertParens(source, offset),
		);
		return item;
	}

	private _applyCompletionInsert(
		item: vscode.CompletionItem,
		name: string,
		range: vscode.Range,
		callable: boolean,
		insertParens: boolean = callable,
	): void {
		item.range = range;
		item.filterText = name;
		if (!callable || !insertParens) {
			item.insertText = name;
			return;
		}
		item.insertText = new vscode.SnippetString(`${name}($0)`);
		item.command = {
			command: 'editor.action.triggerParameterHints',
			title: 'Trigger Parameter Hints',
		};
	}

	private _completionRange(
		document: vscode.TextDocument,
		position: vscode.Position,
	): vscode.Range {
		const line = document.lineAt(position.line).text;
		let start = position.character;
		while (start > 0 && /[A-Za-z0-9_]/.test(line[start - 1])) {
			start -= 1;
		}
		let end = position.character;
		while (end < line.length && /[A-Za-z0-9_]/.test(line[end])) {
			end += 1;
		}
		return new vscode.Range(position.line, start, position.line, end);
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
	let lastCanonicalCandidate = canonicalCandidateFromEditor(vscode.window.activeTextEditor);
	const flushCanonicalCandidate = (): void => {
		const candidate = lastCanonicalCandidate;
		if (!candidate || !isVbaDocument(candidate.editor.document)) {
			return;
		}
		void provider.applyCanonicalCase(
			candidate.editor.document,
			candidate.position,
			candidate.editor,
		);
	};

	context.subscriptions.push(
		vscode.languages.registerCompletionItemProvider(
			selector,
			provider,
			'.',
			' ',
		),
		vscode.workspace.onDidChangeTextDocument((event) => {
			if (!isVbaDocument(event.document) || event.contentChanges.length !== 1) {
				return;
			}
			const change = event.contentChanges[0];
			if (!change.range.isEmpty || !isCanonicalCaseBoundary(change.text)) {
				return;
			}
			void provider.applyCanonicalCase(event.document, change.range.start);
		}),
		vscode.window.onDidChangeTextEditorSelection((event) => {
			const previous = lastCanonicalCandidate;
			if (previous?.editor === event.textEditor) {
				void provider.applyCanonicalCase(
					previous.editor.document,
					previous.position,
					previous.editor,
				);
			}
			lastCanonicalCandidate = canonicalCandidateFromEditor(event.textEditor);
		}),
		vscode.window.onDidChangeActiveTextEditor((editor) => {
			flushCanonicalCandidate();
			lastCanonicalCandidate = canonicalCandidateFromEditor(editor);
		}),
		vscode.window.onDidChangeWindowState((state) => {
			if (!state.focused) {
				flushCanonicalCandidate();
			}
		}),
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

function canonicalCandidateFromEditor(
	editor: vscode.TextEditor | undefined,
): { editor: vscode.TextEditor; position: vscode.Position } | undefined {
	if (!editor || !isVbaDocument(editor.document)) {
		return undefined;
	}
	return { editor, position: editor.selection.active };
}

function isVbaDocument(document: vscode.TextDocument): boolean {
	return document.languageId === 'vba' || document.uri.scheme === XLIDE_SCHEME;
}

function isCanonicalCaseBoundary(text: string): boolean {
	return (
		/^[ \t]$/.test(text) ||
		/^\r?\n[ \t]*$/.test(text) ||
		['(', ')', '.', ',', '=', ':', '+', '-', '*', '/', '\\', '&', '<', '>'].includes(text)
	);
}
