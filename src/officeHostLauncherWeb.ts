// Launching Office from a browser: not possible, and not pretended.
//
// The only reachable caller in a web build is officeWriteCoordinator, which
// reopens a file in its host application after a write. In a browser there is
// no host application, so the coordinator's Windows-only branches never fire
// and this is never called. It refuses rather than resolving quietly, so a
// wrong assumption shows up as a message naming the desktop editor instead of
// a write that silently skipped its coordination step.
//
// Only the members a reachable module actually uses are here. The script
// builders and the macro runners are desktop-only by construction: every
// caller of those is behind platformFeatures.

import type { OfficeHostApp } from './officeHostApps';

export type HostMacroFailureCode = 'REOPEN_BLOCKED' | 'REOPEN_FAILED' | 'RUN_FAILED' | 'UNKNOWN';

export class HostMacroError extends Error {
	constructor(message: string, readonly code: HostMacroFailureCode = 'UNKNOWN') {
		super(message);
		this.name = 'HostMacroError';
	}
}

function noHostApplication(): never {
	throw new HostMacroError(
		'XLIDE in the browser cannot open a file in Excel, Word, PowerPoint or Access. This needs the desktop editor.',
		'REOPEN_BLOCKED',
	);
}

export async function openFileInHost(): Promise<never> {
	return noHostApplication();
}

export async function runHostMacro(): Promise<never> {
	return noHostApplication();
}

export async function showAccessDesign(): Promise<never> {
	return noHostApplication();
}

export function hostMacroReference(
	host: OfficeHostApp,
	moduleName: string,
	procedureName: string,
): string {
	// Pure string composition, and the only member here with a real answer:
	// procedures are named this way whether or not a host application exists.
	// Kept identical to the desktop version, Access's unqualified form
	// included.
	return host === 'access' ? procedureName : `${moduleName}.${procedureName}`;
}
