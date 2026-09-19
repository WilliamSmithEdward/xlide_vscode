import * as fs from 'fs';
import * as os from 'os';
import * as nodePath from 'path';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

interface FakeWatcher {
    pattern: { baseUri: { fsPath: string }; pattern: string };
    change?: (uri: { fsPath: string }) => void;
    create?: (uri: { fsPath: string }) => void;
    disposed: boolean;
}

const watchers = vi.hoisted(() => [] as FakeWatcher[]);

vi.mock('vscode', async () => (await import('./helpers/vscodeMock')).vscodeMock({
    workspace: {
        createFileSystemWatcher: (pattern: FakeWatcher['pattern']) => {
            const watcher: FakeWatcher = { pattern, disposed: false };
            watchers.push(watcher);
            return {
                onDidChange: (listener: FakeWatcher['change']) => { watcher.change = listener; return { dispose: () => undefined }; },
                onDidCreate: (listener: FakeWatcher['create']) => { watcher.create = listener; return { dispose: () => undefined }; },
                onDidDelete: () => ({ dispose: () => undefined }),
                dispose: () => { watcher.disposed = true; },
            };
        },
    },
}));

import {
    checkProjectFile,
    onDidChangeProjectFile,
    recordProjectWrite,
    watchProjectFile,
} from '../src/projectFileChanges';

describe('project file changes made outside XLIDE', () => {
    let tempDir: string;
    let projectPath: string;
    let seen: string[];
    let subscription: { dispose(): void };

    beforeEach(() => {
        tempDir = fs.mkdtempSync(nodePath.join(os.tmpdir(), 'xlide-file-changes-'));
        projectPath = nodePath.join(tempDir, 'Book.xlsm');
        fs.writeFileSync(projectPath, 'one');
        setMtime(Date.parse('2024-01-01T00:00:00Z'));
        seen = [];
        subscription = onDidChangeProjectFile((changed) => seen.push(changed));
        watchers.length = 0;
    });

    afterEach(() => {
        subscription.dispose();
        vi.useRealTimers();
        fs.rmSync(tempDir, { recursive: true, force: true });
    });

    function setMtime(ms: number): void {
        fs.utimesSync(projectPath, new Date(ms), new Date(ms));
    }

    /** Another program saves the file: new content, new modification time. */
    function saveOutsideXlide(content: string, ms: number): void {
        fs.writeFileSync(projectPath, content);
        setMtime(ms);
    }

    it('records the first look at a file without reporting it', () => {
        expect(checkProjectFile(projectPath)).toBe(false);
        expect(seen).toEqual([]);
    });

    it('reports a change it did not write, once', () => {
        checkProjectFile(projectPath);
        saveOutsideXlide('two', Date.parse('2024-01-02T00:00:00Z'));

        expect(checkProjectFile(projectPath)).toBe(true);
        expect(checkProjectFile(projectPath)).toBe(false);
        expect(seen).toEqual([projectPath]);
    });

    it('takes the stamp its own write leaves as its own', async () => {
        checkProjectFile(projectPath);

        await recordProjectWrite(projectPath, async () => saveOutsideXlide('written by xlide', Date.parse('2024-01-02T00:00:00Z')));

        expect(checkProjectFile(projectPath)).toBe(false);
        expect(seen).toEqual([]);
    });

    it('still reports a change made just before its own write', async () => {
        // The VBE saves, and an agent writes before anything looked. Taking the
        // stamp the write leaves would hide the VBE's change from every module
        // the write did not touch.
        checkProjectFile(projectPath);
        saveOutsideXlide('saved in the VBE', Date.parse('2024-01-02T00:00:00Z'));

        await recordProjectWrite(projectPath, async () => saveOutsideXlide('then an agent wrote', Date.parse('2024-01-03T00:00:00Z')));

        expect(checkProjectFile(projectPath)).toBe(true);
        expect(seen).toEqual([projectPath]);
    });

    it('waits out a moment with no file, and reports the file that comes back', () => {
        checkProjectFile(projectPath);
        fs.rmSync(projectPath);

        expect(checkProjectFile(projectPath)).toBe(false);

        saveOutsideXlide('replaced', Date.parse('2024-01-02T00:00:00Z'));
        expect(checkProjectFile(projectPath)).toBe(true);
    });

    it('checks a watched file after its events settle, and ignores the rest of the folder', () => {
        vi.useFakeTimers();
        checkProjectFile(projectPath);
        const watch = watchProjectFile(projectPath);
        expect(watchers).toHaveLength(1);
        expect(watchers[0].pattern.baseUri.fsPath).toBe(tempDir);

        saveOutsideXlide('saved in the VBE', Date.parse('2024-01-02T00:00:00Z'));
        watchers[0].change?.({ fsPath: nodePath.join(tempDir, 'Other.xlsm') });
        vi.advanceTimersByTime(1000);
        expect(seen).toEqual([]);

        // A save through a rename arrives as a create, often more than once.
        watchers[0].create?.({ fsPath: projectPath });
        watchers[0].change?.({ fsPath: projectPath });
        vi.advanceTimersByTime(1000);
        expect(seen).toEqual([projectPath]);

        watch.dispose();
        expect(watchers[0].disposed).toBe(true);
    });

    it('shares one watcher between holders and keeps it until the last lets go', () => {
        const first = watchProjectFile(projectPath);
        const second = watchProjectFile(projectPath);
        expect(watchers).toHaveLength(1);

        first.dispose();
        first.dispose();
        expect(watchers[0].disposed).toBe(false);

        second.dispose();
        expect(watchers[0].disposed).toBe(true);
    });
});
