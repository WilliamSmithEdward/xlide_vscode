import * as vscode from 'vscode';
import {
	XLIDE_SCHEME,
	decodeModuleUri,
	moduleIdentityKey,
	sameProjectPath,
} from './xlideFileSystem';

export interface VbaOpenDocumentLike {
	uri: vscode.Uri;
	getText(): string;
}

export interface VbaOpenModuleSource {
	projectPath: string;
	moduleName: string;
	source: string;
}

export interface VbaSourceModule {
	moduleName: string;
	source: string;
}

export function openXlideModuleSources(
	documents: readonly VbaOpenDocumentLike[] = vscode.workspace.textDocuments ?? [],
): VbaOpenModuleSource[] {
	const out: VbaOpenModuleSource[] = [];
	for (const document of documents) {
		if (document.uri.scheme !== XLIDE_SCHEME) {
			continue;
		}
		try {
			const decoded = decodeModuleUri(document.uri);
			out.push({
				projectPath: decoded.projectPath,
				moduleName: decoded.moduleName,
				source: document.getText(),
			});
		} catch {
			// Ignore non-standard xlide-vba URIs.
		}
	}
	return out;
}

export function openModuleSourceForProject(
	projectPath: string,
	moduleName: string,
	documents: readonly VbaOpenDocumentLike[] = vscode.workspace.textDocuments ?? [],
): string | undefined {
	const moduleKey = moduleIdentityKey(moduleName);
	return openModuleSourceMapForProject(projectPath, documents).get(moduleKey);
}

export function openModuleSourceMapForProject(
	projectPath: string,
	documents: readonly VbaOpenDocumentLike[] = vscode.workspace.textDocuments ?? [],
): Map<string, string> {
	const out = new Map<string, string>();
	for (const open of openXlideModuleSources(documents)) {
		if (sameProjectPath(open.projectPath, projectPath)) {
			out.set(moduleIdentityKey(open.moduleName), open.source);
		}
	}
	return out;
}

export function applyOpenDocumentSources<T extends VbaSourceModule>(
	modules: readonly T[],
	projectPath: string,
	documents: readonly VbaOpenDocumentLike[] = vscode.workspace.textDocuments ?? [],
): T[] {
	const out = modules.map((mod) => ({ ...mod }));
	const byName = new Map(out.map((mod) => [moduleIdentityKey(mod.moduleName), mod]));
	for (const [moduleKey, source] of openModuleSourceMapForProject(projectPath, documents)) {
		const mod = byName.get(moduleKey);
		if (!mod) {
			continue;
		}
		mod.source = source;
	}
	return out;
}
