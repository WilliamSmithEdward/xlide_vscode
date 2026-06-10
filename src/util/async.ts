export async function yieldToExtensionHost(): Promise<void> {
    await new Promise<void>((resolve) => setTimeout(resolve, 0));
}

export async function mapWithConcurrency<T, R>(
    items: readonly T[],
    concurrency: number,
    worker: (item: T, index: number) => Promise<R | undefined>,
): Promise<R[]> {
    const results: Array<R | undefined> = new Array(items.length);
    let nextIndex = 0;
    const workerCount = Math.min(Math.max(1, concurrency), items.length);
    await Promise.all(Array.from({ length: workerCount }, async () => {
        while (nextIndex < items.length) {
            const index = nextIndex++;
            results[index] = await worker(items[index], index);
        }
    }));
    return results.filter((value): value is R => value !== undefined);
}
