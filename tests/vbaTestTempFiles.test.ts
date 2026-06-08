import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { afterEach, describe, expect, it } from 'vitest';
import {
    cleanupStaleVbaTestHostTempDirs,
    cleanupStaleVbaTestHostTempDirsAsync,
    createVbaTestHostTempDir,
    XLIDE_VBA_TEST_HOST_TEMP_PREFIX,
} from '../src/vbaTestTempFiles';

const tempRoots: string[] = [];

afterEach(() => {
    for (const root of tempRoots.splice(0)) {
        fs.rmSync(root, { recursive: true, force: true });
    }
});

describe('VBA test temp files', () => {
    it('creates host temp directories with the XLIDE test-host prefix', () => {
        const root = fs.mkdtempSync(path.join(os.tmpdir(), 'xlide-temp-test-'));
        tempRoots.push(root);

        const dir = createVbaTestHostTempDir(root);

        expect(path.basename(dir).startsWith(XLIDE_VBA_TEST_HOST_TEMP_PREFIX)).toBe(true);
        expect(fs.existsSync(dir)).toBe(true);
    });

    it('cleans only stale XLIDE VBA test host temp directories', () => {
        const root = fs.mkdtempSync(path.join(os.tmpdir(), 'xlide-temp-test-'));
        tempRoots.push(root);
        const stale = path.join(root, `${XLIDE_VBA_TEST_HOST_TEMP_PREFIX}stale`);
        const fresh = path.join(root, `${XLIDE_VBA_TEST_HOST_TEMP_PREFIX}fresh`);
        const unrelated = path.join(root, 'other-stale-folder');
        fs.mkdirSync(stale);
        fs.mkdirSync(fresh);
        fs.mkdirSync(unrelated);
        fs.writeFileSync(path.join(stale, 'workbook.xlsm'), 'copy');
        const now = Date.now();
        const staleDate = new Date(now - 60 * 60 * 1000);
        const freshDate = new Date(now - 5 * 60 * 1000);
        fs.utimesSync(stale, staleDate, staleDate);
        fs.utimesSync(fresh, freshDate, freshDate);
        fs.utimesSync(unrelated, staleDate, staleDate);

        const result = cleanupStaleVbaTestHostTempDirs({
            tmpDir: root,
            olderThanMs: 30 * 60 * 1000,
            nowMs: now,
        });

        expect(result).toEqual({ scanned: 2, deleted: 1, failed: 0 });
        expect(fs.existsSync(stale)).toBe(false);
        expect(fs.existsSync(fresh)).toBe(true);
        expect(fs.existsSync(unrelated)).toBe(true);
    });

    it('cleans stale XLIDE VBA test host temp directories asynchronously', async () => {
        const root = fs.mkdtempSync(path.join(os.tmpdir(), 'xlide-temp-test-'));
        tempRoots.push(root);
        const stale = path.join(root, `${XLIDE_VBA_TEST_HOST_TEMP_PREFIX}stale`);
        const fresh = path.join(root, `${XLIDE_VBA_TEST_HOST_TEMP_PREFIX}fresh`);
        fs.mkdirSync(stale);
        fs.mkdirSync(fresh);
        const now = Date.now();
        const staleDate = new Date(now - 60 * 60 * 1000);
        const freshDate = new Date(now - 5 * 60 * 1000);
        fs.utimesSync(stale, staleDate, staleDate);
        fs.utimesSync(fresh, freshDate, freshDate);

        const result = await cleanupStaleVbaTestHostTempDirsAsync({
            tmpDir: root,
            olderThanMs: 30 * 60 * 1000,
            nowMs: now,
        });

        expect(result).toEqual({ scanned: 2, deleted: 1, failed: 0 });
        expect(fs.existsSync(stale)).toBe(false);
        expect(fs.existsSync(fresh)).toBe(true);
    });
});
