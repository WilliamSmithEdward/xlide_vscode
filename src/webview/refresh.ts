/**
 * Serializes panel operations. Refreshers sharing one gate never refresh while
 * any gated operation is in flight; refreshes that fire meanwhile rerun once idle.
 */
export class RefreshGate {
    private depth = 0;
    private readonly idleCallbacks = new Set<() => void>();

    get inFlight(): boolean {
        return this.depth > 0;
    }

    async runExclusive<T>(operation: () => Promise<T>): Promise<T> {
        this.depth += 1;
        try {
            return await operation();
        } finally {
            this.depth -= 1;
            if (this.depth === 0) {
                const callbacks = [...this.idleCallbacks];
                this.idleCallbacks.clear();
                for (const callback of callbacks) {
                    callback();
                }
            }
        }
    }

    onceIdle(callback: () => void): void {
        this.idleCallbacks.add(callback);
    }
}

export interface DebouncedRefresherOptions {
    /** Runs the actual refresh; failures are routed to onError. */
    refresh: () => Promise<void>;
    onError: (error: unknown) => void;
    defaultDelayMs: number;
    gate?: RefreshGate;
}

/**
 * Debounced webview refresh pump: trailing-edge debounce, stale-request
 * version guard, and queue-behind-exclusive-operation semantics.
 */
export class DebouncedRefresher {
    private timer: ReturnType<typeof setTimeout> | undefined;
    private version = 0;
    private disposed = false;
    private queuedBehindGate = false;
    private readonly gate: RefreshGate;

    constructor(private readonly options: DebouncedRefresherOptions) {
        this.gate = options.gate ?? new RefreshGate();
    }

    /** Trailing-edge debounce; later calls supersede pending ones. */
    schedule(delayMs = this.options.defaultDelayMs): void {
        if (this.disposed) {
            return;
        }
        const requestVersion = ++this.version;
        if (this.timer) {
            clearTimeout(this.timer);
        }
        this.timer = setTimeout(() => {
            this.timer = undefined;
            void this.run(requestVersion);
        }, delayMs);
    }

    /** Cancels any pending schedule and refreshes immediately. */
    async refreshNow(): Promise<void> {
        const requestVersion = ++this.version;
        if (this.timer) {
            clearTimeout(this.timer);
            this.timer = undefined;
        }
        await this.run(requestVersion);
    }

    /**
     * Runs a panel operation that refreshes must not interleave with;
     * a refresh that fires meanwhile reruns once the operation completes.
     */
    runExclusive<T>(operation: () => Promise<T>): Promise<T> {
        return this.gate.runExclusive(operation);
    }

    dispose(): void {
        this.disposed = true;
        if (this.timer) {
            clearTimeout(this.timer);
            this.timer = undefined;
        }
    }

    private async run(requestVersion: number): Promise<void> {
        if (this.disposed || requestVersion !== this.version) {
            return;
        }
        if (this.gate.inFlight) {
            if (!this.queuedBehindGate) {
                this.queuedBehindGate = true;
                this.gate.onceIdle(() => {
                    this.queuedBehindGate = false;
                    this.schedule();
                });
            }
            return;
        }
        try {
            await this.gate.runExclusive(this.options.refresh);
        } catch (err) {
            this.options.onError(err);
        }
    }
}
