// UI-side container facts, decided from the file extension.
//
// The engine classifies containers from CONTENT (macroContainer.ts) and is
// the authority for every actual read and write; these helpers exist for the
// surfaces that must answer synchronously before any file is opened - tree
// context values, file-system stat permissions, discovery globs, and which
// host token an analysis request should carry.

import { hostTokenForFileName, type VbaHostToken } from './analyzer/host/hostRegistry';

/**
 * Every macro-container extension the engine reads, lowercase, no dot. The
 * VB6 project manifest is listed with them: it is not an Office container,
 * but it is a file the engine answers module questions for, and discovery
 * is where "the workspace's projects" is decided.
 */
export const MACRO_CONTAINER_EXTENSIONS = [
	'xlsm', 'xlsb', 'xlam', 'xltm', 'xls', 'xlt', 'xla',
	'docm', 'dotm', 'doc', 'dot',
	'pptm', 'potm', 'ppsm', 'ppam', 'ppt', 'ppa',
	'accdb', 'accda', 'mdb', 'mda',
	'vbp',
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

/** A VB6 project: modules are the files on disk, not streams in a container. */
export function isVb6ProjectPath(fsPath: string): boolean {
	return containerHostForPath(fsPath) === 'vb6';
}

/** The application display name for user-facing messages about a container. */
export function containerAppNameForPath(fsPath: string): string {
	switch (containerHostForPath(fsPath)) {
		case 'word': return 'Word';
		case 'powerpoint': return 'PowerPoint';
		case 'access': return 'Access';
		case 'vb6': return 'Visual Basic 6';
		default: return 'Excel';
	}
}

/** The tree item context value that gates a workbook node's menu surface. */
export function containerContextValue(fsPath: string): 'xlsm' | 'macroDocument' | 'macroReadOnly' | 'vb6Project' {
	if (isVb6ProjectPath(fsPath)) {
		return 'vb6Project';
	}
	if (isReadOnlyContainerPath(fsPath)) {
		return 'macroReadOnly';
	}
	return isExcelContainerPath(fsPath) ? 'xlsm' : 'macroDocument';
}
