/**
 * Evicts the oldest entries of an insertion-ordered map until it holds at
 * most `cap`. A cache that re-inserts on hit gets LRU behaviour from it.
 */
export function evictOldest<K, V>(map: Map<K, V>, cap: number): void {
    while (map.size > cap) {
        const oldest = map.keys().next().value;
        if (oldest === undefined) {
            return;
        }
        map.delete(oldest);
    }
}
