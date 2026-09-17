import type * as vscode from 'vscode';

/**
 * A WorkspaceConfiguration over a plain object: `get` reads it, `update`
 * writes it and records the call, and `inspect` reports a machine-scoped
 * value for the keys named in `machineKeys`.
 */
export function fakeConfig(
    values: Record<string, unknown>,
    machineKeys = new Set<string>(),
    updates: Array<{ key: string; value: unknown; target: unknown }> = [],
): vscode.WorkspaceConfiguration {
    return {
        get: (key: string, fallback?: unknown) => key in values ? values[key] : fallback,
        inspect: (key: string) => machineKeys.has(key) ? { globalValue: values[key] } : {},
        update: (key: string, value: unknown, target?: unknown) => {
            if (value === undefined) {
                delete values[key];
                machineKeys.delete(key);
            } else {
                values[key] = value;
                machineKeys.add(key);
            }
            updates.push({ key, value, target });
            return Promise.resolve();
        },
    } as unknown as vscode.WorkspaceConfiguration;
}
