/** A timed-out analysis must never be retried on the editor host. */
export class AnalysisWorkerTimeoutError extends Error {
    constructor(timeoutMs: number) {
        super(`Analysis worker request timed out after ${timeoutMs} ms.`);
        this.name = 'AnalysisWorkerTimeoutError';
    }
}

export function isAnalysisWorkerTimeoutError(error: unknown): boolean {
    return error instanceof Error && error.name === 'AnalysisWorkerTimeoutError';
}
