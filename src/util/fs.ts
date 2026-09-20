import * as path from 'path';
import { fileExists as platformFileExists } from './fsNode';

/**
 * Whether a path exists.
 *
 * The answer comes from a leaf the browser build swaps (see webBuild.js):
 * github.dev has no node:fs, and its files live behind a virtual filesystem
 * only the editor can see. Kept behind this module so the four callers, and
 * the tests that exercise them, stay unchanged.
 */
export async function fileExists(filePath: string): Promise<boolean> {
    return platformFileExists(filePath);
}

export function isPathInside(baseDir: string, targetPath: string): boolean {
    const base = path.resolve(baseDir);
    const target = path.resolve(targetPath);
    return target === base || target.startsWith(base + path.sep);
}
