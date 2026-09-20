// Everything the container engine asks of the platform it runs on: reading
// and replacing files, and random bytes for the GUIDs a new form or Access
// control needs. Desktop builds get node:fs and node:crypto. The web
// extension host has neither, so the browser build swaps the leaf module
// underneath this one (see webBuild.js), exactly as containerCodec does.
//
// This interface is synchronous and stays that way. A browser cannot read a
// file synchronously, but it does not have to: every engine operation reads a
// whole container, works in memory, and writes it back, so the asynchronous
// work belongs at that boundary. The web leaf serves from a cache the
// extension layer primes before it calls in and drains afterwards, rather
// than threading promises through every reader in src/vba/.

import { platformHost } from './hostPlatformNode';

/** The two fields the engine's caches key on. */
export interface HostFileStat {
	/** Last-write time in milliseconds. */
	mtimeMs: number;
	size: number;
}

export interface HostPlatform {
	/** A name for diagnostics and the support bundle. */
	readonly name: string;

	/** Throws when there is no such file. */
	readFile(filePath: string): Buffer;

	/** Undefined when there is no such file; every other failure throws. */
	readFileIfPresent(filePath: string): Buffer | undefined;

	/** Throws when there is no such file. */
	stat(filePath: string): HostFileStat;

	/** Undefined when there is no such file. */
	statIfPresent(filePath: string): HostFileStat | undefined;

	exists(filePath: string): boolean;

	/**
	 * Replaces a file's contents so no reader ever sees it half written. On a
	 * real filesystem that means writing a sibling and renaming over the
	 * target, carrying the original's mode across.
	 */
	writeFile(filePath: string, data: Buffer): void;

	/** Cryptographically strong bytes. Used only to mint GUIDs. */
	randomBytes(count: number): Buffer;
}

let current: HostPlatform = platformHost;

/** The platform in force. */
export function hostPlatform(): HostPlatform {
	return current;
}

/**
 * Replaces the platform. The browser build does not need this - it swaps the
 * leaf at bundle time - but tests use it to run the engine against an
 * in-memory filesystem, which is how the web leaf is held to the same
 * behavior as the real one.
 */
export function setHostPlatform(platform: HostPlatform): void {
	current = platform;
}
