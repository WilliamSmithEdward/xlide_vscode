// One way to land bytes on disk: a temp file beside the target, then a
// rename over it, so a crash mid-write cannot leave a half-written file.
// Every project save and every VB6 module write goes through here.

import * as fs from 'fs';
import * as path from 'path';

export function atomicWrite(filePath: string, data: Buffer): void {
	const dir = path.dirname(path.resolve(filePath));
	const tmp = path.join(dir, `.xlide-${process.pid}-${Date.now()}.tmp`);
	try {
		fs.writeFileSync(tmp, data);
		try {
			// Preserve the original file mode; a fresh temp file would otherwise
			// narrow permissions on POSIX.
			const stat = fs.statSync(filePath);
			fs.chmodSync(tmp, stat.mode);
		} catch { /* new file: keep the default mode */ }
		fs.renameSync(tmp, filePath);
	} catch (err) {
		try { fs.unlinkSync(tmp); } catch { /* nothing to clean up */ }
		throw err;
	}
}
