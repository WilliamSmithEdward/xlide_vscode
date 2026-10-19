import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';

// Folders a test file makes under the system temp folder, removed together
// once the file's tests have run: call `afterAll(removeTempDirs)` in it.
const made: string[] = [];

export function makeTempDir(prefix: string): string {
	const dir = fs.mkdtempSync(path.join(os.tmpdir(), prefix));
	made.push(dir);
	return dir;
}

export function removeTempDirs(): void {
	for (const dir of made.splice(0)) {
		// Windows can hold a file a moment after it is closed.
		fs.rmSync(dir, { recursive: true, force: true, maxRetries: 5, retryDelay: 100 });
	}
}
