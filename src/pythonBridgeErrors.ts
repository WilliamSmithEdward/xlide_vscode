/** JSON-RPC 2.0 error code python/server.py returns for unknown methods. */
export const JSONRPC_METHOD_NOT_FOUND = -32601;

/**
 * Rejection raised by PythonBridge when the backend returns a JSON-RPC error,
 * preserving the error code alongside the message so callers can branch on
 * the code instead of regex-matching the message text.
 */
export class BridgeError extends Error {
    constructor(message: string, readonly code: number) {
        super(message);
        this.name = 'BridgeError';
    }
}

/**
 * Recognizes the bridge rejection for backends (and test fakes) that do not
 * implement the batch readModules RPC, so callers can fall back to
 * listModules plus per-module readModule calls. Callers only use this inside
 * catch blocks that wrap a single readModules call, so the method-not-found
 * code alone identifies the missing capability.
 */
export function isReadModulesUnavailable(err: unknown): boolean {
    return err instanceof BridgeError && err.code === JSONRPC_METHOD_NOT_FOUND;
}
