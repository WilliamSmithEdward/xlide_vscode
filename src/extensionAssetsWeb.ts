// The web extension host's half of extensionAssets.ts.
//
// A browser has no extension directory to read, so assets travel inside the
// bundle: webBuild.js compiles assets/webview/** into bundledAssets.ts, and
// registerBundledAssets can add more at runtime. The API matches
// extensionAssets.ts exactly, because webBuild.js aliases that module to this
// one for the web bundle (the same swap containerCodecNode/Web uses).
//
// Anything not carried names itself rather than the mechanism: a missing
// asset means one feature is desktop-only, which is what the user needs told.

import { bundledTextAssets } from './webview/bundledAssets';

// The webview templates are compiled into the bundle (see webBuild.js);
// registerBundledAssets can add more at runtime.
const textAssets = new Map<string, string>(Object.entries(bundledTextAssets));
const binaryAssets = new Map<string, Buffer>();

/**
 * Accepted for API parity with the desktop module. A browser has no asset
 * root, so this records nothing.
 */
export function setExtensionAssetRoot(_rootFsPath: string): void {
	/* no filesystem to resolve against */
}

/**
 * Supplies the assets this build carries, keyed by the same
 * extension-root-relative paths the desktop reader uses ('assets/...').
 * Text is stored LF-normalized, matching readExtensionTextAsset.
 */
export function registerBundledAssets(assets: {
	text?: Record<string, string>;
	binary?: Record<string, Uint8Array>;
}): void {
	for (const [key, value] of Object.entries(assets.text ?? {})) {
		textAssets.set(key, value.replace(/\r\n/g, '\n'));
	}
	for (const [key, value] of Object.entries(assets.binary ?? {})) {
		binaryAssets.set(key, Buffer.from(value));
	}
}

export function readExtensionTextAsset(relativePath: string): string {
	const content = textAssets.get(relativePath);
	if (content === undefined) {
		throw new Error(missing(relativePath));
	}
	return content;
}

export function readExtensionBinaryAsset(relativePath: string): Buffer {
	const content = binaryAssets.get(relativePath);
	if (content === undefined) {
		throw new Error(missing(relativePath));
	}
	return content;
}

/**
 * Names the asset rather than the mechanism: what reaches the user is that
 * one feature is desktop-only, not that a bundle map was empty.
 */
function missing(relativePath: string): string {
	return `XLIDE in the browser does not carry ${relativePath}. This feature is available in the desktop editor.`;
}
