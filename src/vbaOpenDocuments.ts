import * as vscode from 'vscode';
import {
	XLIDE_SCHEME,
	decodeModuleUri,
	moduleIdentityKey,
	sameWorkbookPath,
} from './xlideFileSystem';

export interface VbaOpenDocumentLike {
	uri: vscode.Uri;
	getText(): string;
}

export interface VbaOpenModuleSource {
	xlsmPath: string;
	moduleName: string;
	source: string;
}

export interface VbaSourceModule {
	moduleName: string;
	source: string;
}

export function openXlideModuleSources(
	documents: readonly VbaOpenDocumentLike[] = vscode.workspace.textDocuments,
): VbaOpenModuleSource[] {
	const out: VbaOpenModuleSource[] = [];
	for (const document of documents) {
		if (document.uri.scheme !== XLIDE_SCHEME) {
			continue;
		}
		try {
			const decoded = decodeModuleUri(document.uri);
			out.push({
				xlsmPath: decoded.xlsmPath,
				moduleName: decoded.moduleName,
				source: document.getText(),
			});
		} catch {
			// Ignore non-standard xlide-vba URIs.
		}
	}
	return out;
}

export function openModuleSourceForWorkbook(
	xlsmPath: string,
	moduleName: string,
	documents: readonly VbaOpenDocumentLike[] = vscode.workspace.textDocuments,
): string | undefined {
	const moduleKey = moduleIdentityKey(moduleName);
	for (const open of openXlideModuleSources(documents)) {
		if (
			sameWorkbookPath(open.xlsmPath, xlsmPath) &&
			moduleIdentityKey(open.moduleName) === moduleKey
		) {
			return open.source;
		}
	}
	return undefined;
}

export function applyOpenDocumentSources<T extends VbaSourceModule>(
	modules: readonly T[],
	xlsmPath: string,
	documents: readonly VbaOpenDocumentLike[] = vscode.workspace.textDocuments,
): T[] {
	const out = modules.map((mod) => ({ ...mod }));
	const byName = new Map(out.map((mod) => [moduleIdentityKey(mod.moduleName), mod]));
	for (const open of openXlideModuleSources(documents)) {
		if (!sameWorkbookPath(open.xlsmPath, xlsmPath)) {
			continue;
		}
		const mod = byName.get(moduleIdentityKey(open.moduleName));
		if (!mod) {
			continue;
		}
		mod.source = open.source;
	}
	return out;
}
