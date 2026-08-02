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
