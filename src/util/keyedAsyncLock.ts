/**
 * Serializes async actions per key: actions sharing a key run one at a time in
 * call order, while actions with different keys run concurrently. Use this to
 * make a shared resource (a file, a folder) safe against interleaved read /
 * write / delete operations from independent callers.
 *
 * Mirrors the inline queue in workbookSettings.withWorkbookSettingsWriteLock.
 */
export function createKeyedAsyncLock(): <T>(key: string, action: () => Promise<T>) => Promise<T> {
    const queues = new Map<string, Promise<unknown>>();
    return async function withKeyedLock<T>(key: string, action: () => Promise<T>): Promise<T> {
        const previous = queues.get(key) ?? Promise.resolve();
        let release: () => void = () => undefined;
        const current = new Promise<void>((resolve) => {
            release = resolve;
        });
        const queued = previous.catch(() => undefined).then(() => current);
        queues.set(key, queued);
        await previous.catch(() => undefined);
        try {
            return await action();
        } finally {
            release();
            if (queues.get(key) === queued) {
                queues.delete(key);
            }
        }
    };
}
