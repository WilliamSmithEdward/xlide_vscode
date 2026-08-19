// UI-side container facts, decided from the file extension.
//
// The engine classifies containers from CONTENT (macroContainer.ts) and is
// the authority for every actual read and write; these helpers exist for the
// surfaces that must answer synchronously before any file is opened - tree
// context values, file-system stat permissions, discovery globs, and which
// host token an analysis request should carry.

import { hostTokenForFileName, type VbaHostToken } from './analyzer/host/hostRegistry';

/** Every macro-container extension the engine reads, lowercase, no dot. */
export const MACRO_CONTAINER_EXTENSIONS = [
	'xlsm', 'xlsb', 'xlam', 'xls',
	'docm', 'dotm', 'doc',
	'pptm', 'potm', 'ppsm', 'ppt',
	'accdb', 'accda', 'mdb',
] as const;

/** Every macro container the engine reads, for workspace discovery. */
export const MACRO_CONTAINER_GLOB = `**/*.{${MACRO_CONTAINER_EXTENSIONS.join(',')}}`;

/** Alternation of the extensions, for building recognition regexes. */
export const MACRO_CONTAINER_EXTENSION_PATTERN = MACRO_CONTAINER_EXTENSIONS.join('|');

/** The host a container path belongs to ('excel' when unrecognized). */
export function containerHostForPath(fsPath: string): VbaHostToken {
	return hostTokenForFileName(fsPath) ?? 'excel';
}

/** Access is the one read-only host: it runs VBA from compiled p-code, so
 * source writes cannot take effect. Everything else writes. */
export function isReadOnlyContainerPath(fsPath: string): boolean {
	return containerHostForPath(fsPath) === 'access';
}

/** Containers the Excel-specific surfaces (launcher, VBA tests) accept. */
export function isExcelContainerPath(fsPath: string): boolean {
	return containerHostForPath(fsPath) === 'excel';
}

/** The application display name for user-facing messages about a container. */
export function containerAppNameForPath(fsPath: string): string {
	switch (containerHostForPath(fsPath)) {
		case 'word': return 'Word';
		case 'powerpoint': return 'PowerPoint';
		case 'access': return 'Access';
		default: return 'Excel';
	}
}

/** The tree item context value that gates a workbook node's menu surface. */
export function containerContextValue(fsPath: string): 'xlsm' | 'macroDocument' | 'macroReadOnly' {
	if (isReadOnlyContainerPath(fsPath)) {
		return 'macroReadOnly';
	}
	return isExcelContainerPath(fsPath) ? 'xlsm' : 'macroDocument';
}
