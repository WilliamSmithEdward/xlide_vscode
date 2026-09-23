// Just enough of `vscode` for the coordinator, the launcher and the engine to
// run in node, for the live Office suite. `settings` is the configuration the
// coordinator reads; a check sets what it needs and clears it after.
export const settings = new Map<string, unknown>();

export class Disposable {
    static from(...items: Array<{ dispose?: () => void }>): Disposable {
        return new Disposable(() => items.forEach((item) => item.dispose?.()));
    }
    constructor(private readonly onDispose: () => void = () => undefined) {}
    dispose(): void { this.onDispose(); }
}

export class EventEmitter<T> {
    private readonly listeners = new Set<(value: T) => void>();
    readonly event = (listener: (value: T) => void): Disposable => {
        this.listeners.add(listener);
        return new Disposable(() => this.listeners.delete(listener));
    };
    fire(value: T): void { for (const listener of [...this.listeners]) { listener(value); } }
    dispose(): void { this.listeners.clear(); }
}

export class CancellationError extends Error {}

export class RelativePattern {
    constructor(readonly baseUri: unknown, readonly pattern: string) {}
}

export const Uri = {
    file: (fsPath: string) => ({ fsPath, scheme: 'file', path: fsPath }),
    joinPath: (base: { fsPath: string }, ...parts: string[]) => ({ fsPath: [base.fsPath, ...parts].join('\\') }),
};

export const workspace = {
    getConfiguration: () => ({
        get: <T>(key: string, fallback?: T): T => (settings.has(key) ? settings.get(key) as T : fallback as T),
        inspect: (key: string) => (settings.has(key) ? { key, globalValue: settings.get(key) } : { key }),
        has: (key: string) => settings.has(key),
    }),
    createFileSystemWatcher: () => ({
        onDidChange: () => new Disposable(),
        onDidCreate: () => new Disposable(),
        onDidDelete: () => new Disposable(),
        dispose: () => undefined,
    }),
    workspaceFolders: [],
};

export const window = {};
export const env = {};
