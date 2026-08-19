// UI-side container facts, decided from the file extension.
//
// The engine classifies containers from CONTENT (macroContainer.ts) and is
// the authority for every actual read and write; these helpers exist for the
// surfaces that must answer synchronously before any file is opened - tree
// context values, file-system stat permissions, discovery globs, and which
// host token an analysis request should carry.

import { hostTokenForFileName, type VbaHostToken } from './analyzer/host/hostRegistry';

/** Every macro container the engine reads, for workspace discovery. */
export const MACRO_CONTAINER_GLOB =
	'**/*.{xlsm,xlsb,xlam,xls,docm,dotm,doc,pptm,potm,ppsm,ppt,accdb,accda,mdb}';

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

/** The tree item context value that gates a workbook node's menu surface. */
export function containerContextValue(fsPath: string): 'xlsm' | 'macroDocument' | 'macroReadOnly' {
	if (isReadOnlyContainerPath(fsPath)) {
		return 'macroReadOnly';
	}
	return isExcelContainerPath(fsPath) ? 'xlsm' : 'macroDocument';
}
