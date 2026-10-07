import type { VbaToken } from '../lexer/tokenKinds';
import type { ReachingAssignments } from './straightLineValues';

// Small states keep native Map behavior. Large states share their unchanged
// base and copy at most 32 overrides, with one bounded level of lookup.
const SHARED_STATE_MIN_SIZE = 64;
const MAX_OVERRIDES = 32;
const immutableStarts = new WeakSet<ReachingAssignments>();

/** Only analyzer-owned, immutable maps may opt into sharing their base. */
export function shareImmutableReachingStart(start: ReachingAssignments): ReachingAssignments {
	immutableStarts.add(start);
	return start;
}

class ReachingSnapshot implements ReachingAssignments {
	readonly #base: ReachingAssignments;
	readonly #changes: ReadonlyMap<string, readonly VbaToken[]>;
	readonly #size: number;
	get size(): number { return this.#size; }
	constructor(base: ReachingAssignments, changes: ReadonlyMap<string, readonly VbaToken[]>, size: number) {
		this.#base = base; this.#changes = changes; this.#size = size;
	}
	get(key: string): readonly VbaToken[] | undefined { return this.#changes.get(key) ?? this.#base.get(key); }
	has(key: string): boolean { return this.#changes.has(key) || this.#base.has(key); }
	withValue(key: string, value: readonly VbaToken[]): ReachingAssignments {
		if (!this.#changes.has(key) && this.#changes.size >= MAX_OVERRIDES) {
			const next = new Map(this); next.set(key, value); return next;
		}
		const changes = new Map(this.#changes); changes.set(key, value);
		return new ReachingSnapshot(this.#base, changes, this.size + (this.has(key) ? 0 : 1));
	}
	*entries(): MapIterator<[string, readonly VbaToken[]]> {
		for (const [key, value] of this.#base) yield [key, this.#changes.get(key) ?? value];
		for (const [key, value] of this.#changes) if (!this.#base.has(key)) yield [key, value];
	}
	*keys(): MapIterator<string> { for (const [key] of this.entries()) yield key; }
	*values(): MapIterator<readonly VbaToken[]> { for (const [, value] of this.entries()) yield value; }
	[Symbol.iterator](): MapIterator<[string, readonly VbaToken[]]> { return this.entries(); }
	forEach(callback: (value: readonly VbaToken[], key: string, map: ReachingAssignments) => void, thisArg?: unknown): void {
		for (const [key, value] of this.entries()) callback.call(thisArg, value, key, this);
	}
}

/** Change one value without copying every unaffected fact into the snapshot. */
export function withReachingValue(before: ReachingAssignments, key: string, value: readonly VbaToken[]): ReachingAssignments {
	if (before instanceof ReachingSnapshot) return before.withValue(key, value);
	if (before.size < SHARED_STATE_MIN_SIZE) { const next = new Map(before); next.set(key, value); return next; }
	// Detach the shared base once: callers may own a mutable initial Map.
	return new ReachingSnapshot(immutableStarts.has(before) ? before : new Map(before), new Map([[key, value]]), before.size + (before.has(key) ? 0 : 1));
}
