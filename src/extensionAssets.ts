import * as fs from 'fs';
import * as path from 'path';

/**
 * Resolves bundled extension assets (assets/**) at runtime.
 *
 * The installed extension layout is `<extensionRoot>/out/extension.js` plus
 * `<extensionRoot>/assets/...`, so assets cannot be resolved relative to the
 * src/ tree. activate() registers the authoritative root from
 * `context.extensionUri.fsPath`; the `__dirname/..` fallback only serves
 * contexts that never activate the extension (vitest runs from src/, where
 * `..` is the repo root, and the bundle runs from out/, where `..` is the
 * extension root).
 */

let extensionAssetRoot: string | undefined;
const assetCache = new Map<string, string>();

/** Called once from activate() with context.extensionUri.fsPath. */
export function setExtensionAssetRoot(rootFsPath: string): void {
    if (extensionAssetRoot !== rootFsPath) {
        extensionAssetRoot = rootFsPath;
        assetCache.clear();
    }
}

/**
 * Reads a UTF-8 text asset by extension-root-relative path (e.g.
 * 'assets/testhost/XlideTestModalWatcher.cs'). Line endings are normalized to
 * LF so behavior does not depend on checkout/packaging EOL conversion. The
 * content is cached: assets are static for the lifetime of the extension host.
 */
export function readExtensionTextAsset(relativePath: string): string {
    const cached = assetCache.get(relativePath);
    if (cached !== undefined) {
        return cached;
    }
    const root = extensionAssetRoot ?? path.resolve(__dirname, '..');
    const content = fs
        .readFileSync(path.join(root, relativePath), 'utf8')
        .replace(/\r\n/g, '\n');
    assetCache.set(relativePath, content);
    return content;
}
