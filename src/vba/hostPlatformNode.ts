// The desktop platform: node:fs and node:crypto, which is what every
// non-browser build and the whole test suite run on.
//
// The browser build never reaches this module - webBuild.js aliases it to
// hostPlatformWeb.ts - which is what keeps node:fs out of the web bundle.

import * as fs from 'fs';
import * as path from 'path';
import { randomBytes } from 'crypto';
import type { HostFileStat, HostPlatform } from './hostPlatform';

function isMissing(err: unknown): boolean {
	return (err as NodeJS.ErrnoException)?.code === 'ENOENT';
}

export const platformHost: HostPlatform = {
	name: 'node',

	readFile(filePath: string): Buffer {
		return fs.readFileSync(filePath);
	},

	readFileIfPresent(filePath: string): Buffer | undefined {
		try {
			return fs.readFileSync(filePath);
		} catch (err) {
			if (isMissing(err)) {
				return undefined;
			}
			throw err;
		}
	},

	stat(filePath: string): HostFileStat {
		const stat = fs.statSync(filePath);
		return { mtimeMs: stat.mtimeMs, size: stat.size };
	},

	statIfPresent(filePath: string): HostFileStat | undefined {
		try {
			const stat = fs.statSync(filePath);
			return { mtimeMs: stat.mtimeMs, size: stat.size };
		} catch {
			return undefined;
		}
	},

	exists(filePath: string): boolean {
		return fs.existsSync(filePath);
	},

	/**
	 * A temp file beside the target, then a rename over it, so a crash
	 * mid-write cannot leave a half-written file. Every project save and
	 * every VB6 module write lands this way.
	 */
	writeFile(filePath: string, data: Buffer): void {
		const dir = path.dirname(path.resolve(filePath));
		const tmp = path.join(dir, `.xlide-${process.pid}-${Date.now()}.tmp`);
		try {
			fs.writeFileSync(tmp, data);
			try {
				// Preserve the original file mode; a fresh temp file would
				// otherwise narrow permissions on POSIX.
				const stat = fs.statSync(filePath);
				fs.chmodSync(tmp, stat.mode);
			} catch { /* new file: keep the default mode */ }
			fs.renameSync(tmp, filePath);
		} catch (err) {
			try { fs.unlinkSync(tmp); } catch { /* nothing to clean up */ }
			throw err;
		}
	},

	randomBytes(count: number): Buffer {
		return randomBytes(count);
	},
};
