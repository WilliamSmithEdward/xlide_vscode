// Where a reference to another VBA project points (issue #1503), and whether
// XLIDE may read it.
//
// A project reference records the referenced file twice ([MS-OVBA]
// 2.3.4.2.2.2): an absolute libid, `*\CC:\Libs\ProbeLibrary.xlam`, and one
// relative to the referring file, `*\CProbeLibrary.xlam`, which is where the
// VBE looks once the absolute path is gone (measured in Excel 16.0: Excel
// wrote both). XLIDE tries them in that order.
//
// The path comes from inside the user's file, so it is not trusted. Only a
// local file is read: a UNC path, `\\host\share\x.xlam`, would have Windows
// contact that host and offer the user's credentials as it does, so a path
// that is or resolves to one is never opened, nor a device path. The file
// must be a macro-enabled Office file of a kind a reference can name, and of
// a size a VBA project file has. A path that fails any of these is left
// unread, and the reference is judged as one XLIDE could not read.

import * as path from 'path';
import { hostPlatform } from './hostPlatform';

/** The files a project reference can name: workbooks and add-ins, documents and templates, presentations and add-ins, databases. */
const REFERENCEABLE = new Set([
	'.xlsm', '.xlsb', '.xlam', '.xls', '.xla', '.xltm', '.xlt',
	'.docm', '.dotm', '.doc', '.dot',
	'.pptm', '.ppam', '.potm', '.ppt', '.ppa', '.pot',
	'.accdb', '.accda', '.accde', '.mdb', '.mda', '.mde',
]);
/** Larger than any VBA project file has a reason to be; a guard against reading whatever a path names. */
const MAX_BYTES = 256 * 1024 * 1024;

/** The path a project libid carries, `*\C<path>`, or undefined. */
function libidPath(libid: string | undefined): string | undefined {
	const match = /^\*\\[A-Za-z](.+)$/s.exec(libid ?? '');
	return match ? match[1] : undefined;
}

/** True for a path Windows would reach over the network or as a device: `\\host\share`, `\\?\`, `\\.\`, `//host`. */
function isRemoteOrDevice(candidate: string): boolean {
	return /^[\\/]{2}/.test(candidate);
}

/** A name Windows reads as a device whatever its extension: `CON.xlam` is the console. */
const DEVICE_NAME = /^(con|prn|aux|nul|com[0-9]|lpt[0-9])(\..*)?$/i;

/**
 * The file name alone of a recorded path, or undefined when it is no plain
 * name: empty, `.` or `..`, a device, or one with a colon, which Windows
 * reads as a drive (`C:Lib.xlam`) or an alternate stream (`Book.xlsm:Lib`).
 * A name recorded on a network or device path gives none, so a local file
 * never answers for `\\host\share\Lib.xlam`.
 */
function plainFileName(recorded: string): string | undefined {
	if (isRemoteOrDevice(recorded)) {
		return undefined;
	}
	const name = recorded.split(/[\\/]/).pop() ?? '';
	if (name === '' || name === '.' || name === '..' || name.includes(':') || DEVICE_NAME.test(name) || /[\u0000-\u001f]/.test(name)) {
		return undefined;
	}
	return name;
}

/** A local file of a referenceable kind and a plausible size, or undefined. */
function readableFile(candidate: string): string | undefined {
	if (isRemoteOrDevice(candidate) || !REFERENCEABLE.has(path.extname(candidate).toLowerCase())
		|| DEVICE_NAME.test(path.basename(candidate))) {
		return undefined;
	}
	// A folder that happens to end in .xlam fails the read that follows,
	// which leaves the reference unread as any other failure does.
	const stat = hostPlatform().statIfPresent(candidate);
	return stat && stat.size > 0 && stat.size <= MAX_BYTES ? candidate : undefined;
}

/**
 * The local file a project reference names, from its absolute libid and then
 * its relative one against the referring file's folder, or undefined when
 * neither names one XLIDE may read. The referring file itself is never one.
 *
 * Then the file name each records, looked for beside the referring file,
 * the absolute one's first: that is where the VBE finds a provider moved
 * together with its consumer, and only Excel records a relative path that
 * still names it after a move (Word and PowerPoint record the absolute path
 * again, Access one relative to another folder; issue #1532).
 */
export function referencedProjectFile(
	libid: string,
	relativeLibid: string | undefined,
	referringFile: string,
): string | undefined {
	const referring = path.resolve(referringFile);
	const candidates: string[] = [];
	const absolute = libidPath(libid);
	// A drive path only: C:\Libs\x.xlam. Where paths are not Windows' it
	// names nothing.
	if (absolute && /^[A-Za-z]:[\\/]/.test(absolute) && path.sep === '\\') {
		candidates.push(path.normalize(absolute));
	}
	const relative = libidPath(relativeLibid);
	if (relative && !isRemoteOrDevice(relative) && !/^[A-Za-z]:/.test(relative) && !isRemoteOrDevice(referring)) {
		candidates.push(path.resolve(path.dirname(referring), ...relative.split(/[\\/]/)));
	}
	if (!isRemoteOrDevice(referring)) {
		for (const recorded of [absolute, relative]) {
			const name = recorded === undefined ? undefined : plainFileName(recorded);
			const beside = name === undefined ? undefined : path.join(path.dirname(referring), name);
			if (beside !== undefined && !candidates.some((candidate) => candidate.toLowerCase() === beside.toLowerCase())) {
				candidates.push(beside);
			}
		}
	}
	for (const candidate of candidates) {
		if (path.resolve(candidate).toLowerCase() === referring.toLowerCase()) {
			continue;
		}
		const file = readableFile(candidate);
		if (file) {
			return file;
		}
	}
	return undefined;
}
