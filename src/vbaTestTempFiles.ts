import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';

export const XLIDE_VBA_TEST_HOST_TEMP_PREFIX = 'xlide-vba-test-host-';
export const DEFAULT_STALE_VBA_TEST_HOST_TEMP_AGE_MS = 24 * 60 * 60 * 1000;

export interface CleanupStaleVbaTestHostTempDirsOptions {
    tmpDir?: string;
    olderThanMs?: number;
    nowMs?: number;
}

export interface CleanupStaleVbaTestHostTempDirsResult {
    scanned: number;
    deleted: number;
    failed: number;
}

export function createVbaTestHostTempDir(tmpDir = os.tmpdir()): string {
    return fs.mkdtempSync(path.join(tmpDir, XLIDE_VBA_TEST_HOST_TEMP_PREFIX));
}

export function cleanupStaleVbaTestHostTempDirs(
    options: CleanupStaleVbaTestHostTempDirsOptions = {},
): CleanupStaleVbaTestHostTempDirsResult {
    const tmpDir = options.tmpDir ?? os.tmpdir();
    const olderThanMs = options.olderThanMs ?? DEFAULT_STALE_VBA_TEST_HOST_TEMP_AGE_MS;
    const nowMs = options.nowMs ?? Date.now();
    let scanned = 0;
    let deleted = 0;
    let failed = 0;

    for (const entry of fs.readdirSync(tmpDir, { withFileTypes: true })) {
        if (!entry.isDirectory() || !entry.name.startsWith(XLIDE_VBA_TEST_HOST_TEMP_PREFIX)) {
            continue;
        }
        scanned++;
        const fullPath = path.join(tmpDir, entry.name);
        try {
            const stat = fs.statSync(fullPath);
            if (nowMs - stat.mtimeMs < olderThanMs) {
                continue;
            }
            fs.rmSync(fullPath, { recursive: true, force: true });
            deleted++;
        } catch {
            failed++;
        }
    }

    return { scanned, deleted, failed };
}
