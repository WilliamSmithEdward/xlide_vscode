import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { afterEach, beforeEach, expect, it, vi } from 'vitest';
import { makeTempDir, removeTempDirs } from './helpers/tempDirs';

vi.mock('fs', () => ({ mkdtempSync: vi.fn(), rmSync: vi.fn() }));

beforeEach(() => {
    vi.resetAllMocks();
    let next = 0;
    vi.mocked(fs.mkdtempSync).mockImplementation(prefix => `${prefix}owned-${next++}`);
});

afterEach(() => {
    // Drain retained paths through the mock even if an assertion fails.
    vi.mocked(fs.rmSync).mockReset();
    removeTempDirs();
});

it('removes only successfully created directories and drains successful batches', () => {
    const first = makeTempDir('xlide-native-');
    const second = makeTempDir('xlide-access-');
    expect(fs.mkdtempSync).toHaveBeenNthCalledWith(1, path.join(os.tmpdir(), 'xlide-native-'));
    removeTempDirs();
    expect(vi.mocked(fs.rmSync).mock.calls.map(([dir]) => dir)).toEqual([first, second]);
    expect(fs.rmSync).toHaveBeenCalledWith(first,
        { recursive: true, force: true, maxRetries: 5, retryDelay: 100 });
    removeTempDirs();
    expect(fs.rmSync).toHaveBeenCalledTimes(2);
    const third = makeTempDir('xlide-lines-');
    removeTempDirs();
    expect(fs.rmSync).toHaveBeenLastCalledWith(third,
        { recursive: true, force: true, maxRetries: 5, retryDelay: 100 });
});

it('does not register a directory when creation fails', () => {
    vi.mocked(fs.mkdtempSync).mockImplementationOnce(() => { throw new Error('creation failed'); });
    expect(() => makeTempDir('xlide-native-')).toThrow('creation failed');
    removeTempDirs();
    expect(fs.rmSync).not.toHaveBeenCalled();
});

it('attempts the whole batch, reports every failure, and retries only failed directories', () => {
    const first = makeTempDir('xlide-native-');
    const second = makeTempDir('xlide-access-');
    const third = makeTempDir('xlide-lines-');
    const firstError = new Error('first directory locked');
    const thirdError = new Error('third directory locked');
    vi.mocked(fs.rmSync).mockImplementation(dir => {
        if (dir === first) { throw firstError; }
        if (dir === third) { throw thirdError; }
    });
    let failure: unknown;
    try { removeTempDirs(); } catch (error) { failure = error; }
    expect(failure).toBeInstanceOf(AggregateError);
    expect((failure as AggregateError).errors).toEqual([firstError, thirdError]);
    expect(vi.mocked(fs.rmSync).mock.calls.map(([dir]) => dir)).toEqual([first, second, third]);

    vi.mocked(fs.rmSync).mockReset();
    const later = makeTempDir('xlide-create-');
    removeTempDirs();
    expect(vi.mocked(fs.rmSync).mock.calls.map(([dir]) => dir)).toEqual([first, third, later]);
    removeTempDirs();
    expect(fs.rmSync).toHaveBeenCalledTimes(3);
});
