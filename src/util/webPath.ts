// A POSIX `path` for the web extension host, which has no node:path.
//
// The browser build aliases `path` to this module (see webBuild.js). That is
// sound because every path in a web workspace comes from a vscode.Uri and is
// POSIX by construction - `vscode-vfs://github/owner/repo/Book.xlsm` - so
// there are no drive letters or backslashes to handle.
//
// node:path's own posix implementation is the oracle: tests/webPath.test.ts
// asserts identical output across a generated corpus of path shapes. Where
// the browser has no equivalent at all the difference is documented below
// rather than guessed at.

/** Working directory for resolve(). A browser has no cwd; the root stands in. */
const CWD = '/';

export const sep = '/';
export const delimiter = ':';

/**
 * Applies '.' and '..' to a split path. `allowAboveRoot` keeps leading '..'
 * segments, which a relative path may have and an absolute one may not.
 */
function normalizeSegments(parts: readonly string[], allowAboveRoot: boolean): string[] {
	const out: string[] = [];
	for (const part of parts) {
		if (part === '' || part === '.') {
			continue;
		}
		if (part === '..') {
			if (out.length > 0 && out[out.length - 1] !== '..') {
				out.pop();
			} else if (allowAboveRoot) {
				out.push('..');
			}
			continue;
		}
		out.push(part);
	}
	return out;
}

export function isAbsolute(p: string): boolean {
	return p.charCodeAt(0) === 47;
}

export function normalize(p: string): string {
	if (p.length === 0) {
		return '.';
	}
	const absolute = isAbsolute(p);
	const trailingSlash = p.charCodeAt(p.length - 1) === 47;

	let joined = normalizeSegments(p.split('/'), !absolute).join('/');
	if (joined.length === 0) {
		// Everything cancelled out. A trailing slash still distinguishes
		// './' from '.', which join() relies on.
		return absolute ? '/' : trailingSlash ? './' : '.';
	}
	if (trailingSlash) {
		joined += '/';
	}
	return absolute ? `/${joined}` : joined;
}

export function join(...parts: string[]): string {
	const joined = parts.filter((part) => part.length > 0).join('/');
	return joined.length === 0 ? '.' : normalize(joined);
}

export function resolve(...parts: string[]): string {
	let resolved = '';
	let absolute = false;

	for (let i = parts.length - 1; i >= 0 && !absolute; i--) {
		const part = parts[i];
		if (part.length === 0) {
			continue;
		}
		resolved = resolved.length > 0 ? `${part}/${resolved}` : part;
		absolute = isAbsolute(part);
	}
	if (!absolute) {
		resolved = resolved.length > 0 ? `${CWD}${resolved}` : CWD;
	}

	const joined = normalizeSegments(resolved.split('/'), false).join('/');
	return `/${joined}`;
}

export function dirname(p: string): string {
	if (p.length === 0) {
		return '.';
	}
	const absolute = isAbsolute(p);
	let end = -1;
	let sawNonSlash = false;
	for (let i = p.length - 1; i >= 1; i--) {
		if (p.charCodeAt(i) === 47) {
			if (sawNonSlash) {
				end = i;
				break;
			}
		} else {
			sawNonSlash = true;
		}
	}
	if (end === -1) {
		return absolute ? '/' : '.';
	}
	if (absolute && end === 1) {
		return '//';
	}
	return p.slice(0, end);
}

export function basename(p: string, ext?: string): string {
	// A path that is nothing but the extension has no base name at all:
	// basename('.txt', '.txt') is '', while basename('/a/.txt', '.txt') keeps
	// '.txt', because stripping it would leave the file nameless.
	if (ext !== undefined && ext.length > 0 && ext === p) {
		return '';
	}

	let start = 0;
	let end = -1;
	let sawNonSlash = false;

	for (let i = p.length - 1; i >= 0; i--) {
		if (p.charCodeAt(i) === 47) {
			if (sawNonSlash) {
				start = i + 1;
				break;
			}
		} else if (end === -1) {
			sawNonSlash = true;
			end = i + 1;
		}
	}
	if (end === -1) {
		return '';
	}

	const base = p.slice(start, end);
	if (ext !== undefined && ext.length > 0 && ext.length < base.length && base.endsWith(ext)) {
		return base.slice(0, base.length - ext.length);
	}
	return base;
}

/**
 * The rules here are subtler than "text after the last dot": '.bashrc' and
 * '..' have no extension, but '...' has one, and '/a/.txt' differs from
 * '.txt'. This follows node:path's own scan so the two cannot drift.
 */
export function extname(p: string): string {
	let startDot = -1;
	let startPart = 0;
	let end = -1;
	let sawNonSlash = false;
	// 0 until a dot is seen, 1 once a run of dots precedes it, -1 once a
	// non-dot character does. Only the -1 and the long-dot-run cases have an
	// extension.
	let preDotState = 0;

	for (let i = p.length - 1; i >= 0; i--) {
		const code = p.charCodeAt(i);
		if (code === 47) {
			if (sawNonSlash) {
				startPart = i + 1;
				break;
			}
			continue;
		}
		if (end === -1) {
			sawNonSlash = true;
			end = i + 1;
		}
		if (code === 46) {
			if (startDot === -1) {
				startDot = i;
			} else if (preDotState !== 1) {
				preDotState = 1;
			}
		} else if (startDot !== -1) {
			preDotState = -1;
		}
	}

	if (
		startDot === -1 ||
		end === -1 ||
		preDotState === 0 ||
		(preDotState === 1 && startDot === end - 1 && startDot === startPart + 1)
	) {
		return '';
	}
	return p.slice(startDot, end);
}

export function relative(from: string, to: string): string {
	const fromResolved = resolve(from);
	const toResolved = resolve(to);
	if (fromResolved === toResolved) {
		return '';
	}

	const fromParts = fromResolved.split('/').filter((part) => part.length > 0);
	const toParts = toResolved.split('/').filter((part) => part.length > 0);

	let shared = 0;
	while (shared < fromParts.length && shared < toParts.length && fromParts[shared] === toParts[shared]) {
		shared++;
	}

	const up = new Array(fromParts.length - shared).fill('..');
	return [...up, ...toParts.slice(shared)].join('/');
}

export const posix = {
	sep,
	delimiter,
	isAbsolute,
	normalize,
	join,
	resolve,
	dirname,
	basename,
	extname,
	relative,
};

/**
 * Windows path semantics, which a web workspace never has: its paths are
 * vscode.Uri paths. Every caller reaches this through a platform test that is
 * false in a browser, so these throw rather than answering a question they
 * cannot answer correctly - being loudly absent beats being quietly wrong
 * about where a user's file is.
 */
function noWin32(): never {
	throw new Error('XLIDE in the browser has no Windows path support; workspace paths are POSIX.');
}

export const win32 = {
	sep: '\\',
	delimiter: ';',
	isAbsolute: noWin32,
	normalize: noWin32,
	join: noWin32,
	resolve: noWin32,
	dirname: noWin32,
	basename: noWin32,
	extname: noWin32,
	relative: noWin32,
};

export default {
	sep,
	delimiter,
	isAbsolute,
	normalize,
	join,
	resolve,
	dirname,
	basename,
	extname,
	relative,
	posix,
	win32,
};
