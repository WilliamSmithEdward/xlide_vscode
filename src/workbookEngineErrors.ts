/** JSON-RPC 2.0 error code the engine returns for unknown methods. */
export const JSONRPC_METHOD_NOT_FOUND = -32601;

/**
 * Rejection raised by WorkbookEngine when a call fails, preserving a
 * JSON-RPC-style error code alongside the message so callers can branch on
 * the code instead of regex-matching the message text.
 */
export class WorkbookEngineError extends Error {
    constructor(message: string, readonly code: number) {
        super(message);
        this.name = 'WorkbookEngineError';
    }
}

/**
 * Recognizes the rejection from engines (and test fakes) that do not
 * implement the batch readModules RPC, so callers can fall back to
 * listModules plus per-module readModule calls. Callers only use this inside
 * catch blocks that wrap a single readModules call, so the method-not-found
 * code alone identifies the missing capability.
 */
export function isReadModulesUnavailable(err: unknown): boolean {
    return err instanceof WorkbookEngineError && err.code === JSONRPC_METHOD_NOT_FOUND;
}
