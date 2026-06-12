export interface DisposableDebounce {
    (): void;
    /** Drops any pending invocation without disabling the debounce. */
    cancel(): void;
    /** Drops any pending invocation; alias of cancel so the result is a Disposable. */
    dispose(): void;
}

/** Trailing-edge debounce: invokes `fn` once `delayMs` after the latest call. */
export function debounce(fn: () => void, delayMs: number): DisposableDebounce {
    let timer: ReturnType<typeof setTimeout> | undefined;
    const debounced = () => {
        if (timer !== undefined) {
            clearTimeout(timer);
        }
        timer = setTimeout(() => {
            timer = undefined;
            fn();
        }, delayMs);
    };
    debounced.cancel = () => {
        if (timer !== undefined) {
            clearTimeout(timer);
            timer = undefined;
        }
    };
    debounced.dispose = debounced.cancel;
    return debounced;
}
