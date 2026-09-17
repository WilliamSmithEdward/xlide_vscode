// The Office applications XLIDE drives, described once.
//
// Everything that talks to a host application reads its identity from this
// table: the launcher, write coordination, the COM availability probe and the
// test host. A host is added here and nowhere else.

import { containerHostForPath } from './macroContainerUi';

/** The Office applications that own a macro container. */
export type OfficeHostApp = 'excel' | 'word' | 'powerpoint' | 'access';

export interface OfficeHostAppInfo {
	/** COM ProgID of the application object. */
	progId: string;
	/** Process image name without `.exe`: the force-close last resort. */
	processName: string;
	/** Display name for user-facing text: "Word". */
	noun: string;
	/** What the application calls one of its files: "document". */
	fileNoun: string;
	/** The application bundle `open -a` takes on macOS; Access has none. */
	macAppName?: string;
	/** The LibreOffice module that opens this host's files on Linux. */
	libreOfficeFlag?: string;
}

export const OFFICE_HOST_APPS: Record<OfficeHostApp, OfficeHostAppInfo> = {
	excel: {
		progId: 'Excel.Application',
		processName: 'EXCEL',
		noun: 'Excel',
		fileNoun: 'workbook',
		macAppName: 'Microsoft Excel',
		libreOfficeFlag: '--calc',
	},
	word: {
		progId: 'Word.Application',
		processName: 'WINWORD',
		noun: 'Word',
		fileNoun: 'document',
		macAppName: 'Microsoft Word',
		libreOfficeFlag: '--writer',
	},
	powerpoint: {
		progId: 'PowerPoint.Application',
		processName: 'POWERPNT',
		noun: 'PowerPoint',
		fileNoun: 'presentation',
		macAppName: 'Microsoft PowerPoint',
		libreOfficeFlag: '--impress',
	},
	access: {
		progId: 'Access.Application',
		processName: 'MSACCESS',
		noun: 'Access',
		fileNoun: 'database',
	},
};

/**
 * The Office application that owns a container path. A VB6 project has none:
 * its modules are plain files that no application holds open.
 */
export function officeHostForPath(fsPath: string): OfficeHostApp | undefined {
	const host = containerHostForPath(fsPath);
	return host in OFFICE_HOST_APPS ? (host as OfficeHostApp) : undefined;
}
