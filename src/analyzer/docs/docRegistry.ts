// External documentation registry.
//
// Holds the doc entries parsed from every external metadata file in the project
// and resolves a symbol name (with an optional qualifier - module name for user
// symbols, receiver type for host members) to its documentation.
//
// Lookup precedence inside the registry:
//   1. Qualified match     - `Qualifier.Name` (most specific)
//   2. Bare match          - an unqualified `<member name="Name">` entry
//   3. Member-of-qualified - the trailing `Name` of any `X.Name` entry
//
// The overall tooltip precedence (developer overrides library) is enforced at
// the call sites: an inline `'''` comment on a symbol wins over this registry,
// and this registry wins over the curated host/runtime library. See
// docs/vba-doc-comments.md.

import { VbaDoc } from './docModel';
import { ExternalDocEntry } from './externalDoc';

export class DocRegistry {
	private readonly _qualified = new Map<string, VbaDoc>();
	private readonly _bare = new Map<string, VbaDoc>();
	private readonly _byMember = new Map<string, VbaDoc>();

	/** Removes every entry. */
	clear(): void {
		this._qualified.clear();
		this._bare.clear();
		this._byMember.clear();
	}

	/** True when no entries are loaded. */
	get isEmpty(): boolean {
		return (
			this._qualified.size === 0 &&
			this._bare.size === 0 &&
			this._byMember.size === 0
		);
	}

	/** Adds entries from one parsed metadata file. Later adds override earlier. */
	add(entries: ExternalDocEntry[]): void {
		for (const entry of entries) {
			const lower = entry.name.toLowerCase();
			const dot = lower.lastIndexOf('.');
			if (dot >= 0) {
				this._qualified.set(lower, entry.doc);
				const member = lower.slice(dot + 1);
				if (member) {
					this._byMember.set(member, entry.doc);
				}
			} else {
				this._bare.set(lower, entry.doc);
			}
		}
	}

	/**
	 * Resolves documentation for `name`, optionally within `qualifier` (the owning
	 * module for a user symbol, or the receiver type for a host member). Returns
	 * undefined when nothing matches.
	 */
	lookup(name: string, qualifier?: string): VbaDoc | undefined {
		const lowerName = name.toLowerCase();
		if (qualifier) {
			const q = this._qualified.get(`${qualifier.toLowerCase()}.${lowerName}`);
			if (q) {
				return q;
			}
		}
		return this._bare.get(lowerName) ?? this._byMember.get(lowerName);
	}
}
