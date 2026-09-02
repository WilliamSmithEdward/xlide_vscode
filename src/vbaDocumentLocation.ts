// Where a VBA document lives, for every surface that needs the container
// and module behind an editor: a workbook module carries both in its
// `xlide-vba:` URI, and a VB6 module is a file on disk whose project the
// locator knows. One answer for both, so no provider has to ask twice.

import * as vscode from 'vscode';
import { XLIDE_SCHEME, decodeModuleUri, encodeModuleUri } from './xlideFileSystem';
import { vb6ModuleOwnerOf } from './vb6ProjectLocator';
import { blankDesignerHeader } from './vba/moduleSource';

export interface ModuleLocation {
	/** The container: a workbook path, or a VB6 project's `.vbp`. */
	xlsmPath: string;
	moduleName: string;
	/** The module's kind, when the locator knows it (VB6 modules only). */
	moduleType?: string;
	/** True when the document IS the module's file rather than a virtual view of it. */
	native: boolean;
}

/**
 * The container and module a document belongs to, or undefined for a loose
 * file no project claims (which is analyzed as a module on its own).
 */
export function moduleLocationOfDocument(document: vscode.TextDocument): ModuleLocation | undefined {
	return moduleLocationOfUri(document.uri);
}

/** The same answer from a URI alone, for caches keyed by URI string. */
export function moduleLocationOfUri(uri: vscode.Uri): ModuleLocation | undefined {
	if (uri.scheme === XLIDE_SCHEME) {
		try {
			const decoded = decodeModuleUri(uri);
			return { xlsmPath: decoded.xlsmPath, moduleName: decoded.moduleName, native: false };
		} catch {
			return undefined;
		}
	}
	if (uri.scheme !== 'file') {
		return undefined;
	}
	const owner = vb6ModuleOwnerOf(uri.fsPath);
	if (!owner) {
		return undefined;
	}
	return { xlsmPath: owner.vbpPath, moduleName: owner.moduleName, moduleType: owner.moduleType, native: true };
}

/** The location, or an error for the surfaces that cannot work without one. */
export function moduleLocationOrThrow(document: vscode.TextDocument): ModuleLocation {
	const location = moduleLocationOfDocument(document);
	if (!location) {
		throw new Error('Not a project VBA module: no workbook or VB6 project claims this document.');
	}
	return location;
}

/**
 * The text the analyzer should see for a document. A workbook module's
 * virtual document is already the code; a file on disk may open with a
 * designer block (`VERSION 5.00` / `Begin VB.Form` ... `End`, or a class
 * preamble) that is not VBA, so it is blanked to whitespace - every offset
 * and line number stays the file's own, and nothing the block says reaches
 * the parser.
 */
export function analysisSourceForDocument(document: vscode.TextDocument): string {
	const text = document.getText();
	if (document.uri.scheme === XLIDE_SCHEME) {
		return text;
	}
	return blankDesignerHeader(text);
}

/**
 * The URI an editor opens for a module the project index knows: the file
 * itself when the module is one (a VB6 project), else the workbook module's
 * virtual document.
 */
export function moduleDocumentUri(
	xlsmPath: string,
	module: { moduleName: string; filePath?: string },
): vscode.Uri {
	return module.filePath ? vscode.Uri.file(module.filePath) : encodeModuleUri(xlsmPath, module.moduleName);
}
