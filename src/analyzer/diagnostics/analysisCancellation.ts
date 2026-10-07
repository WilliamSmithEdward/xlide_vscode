/** Cooperative cancellation is control flow, never an analyzer failure. */
export class AnalysisCancelled extends Error {
	constructor() { super('Analysis snapshot superseded.'); this.name = 'AnalysisSnapshotSuperseded'; }
}

export function checkAnalysisCancellation(options: { isCancelled?: () => boolean }): void {
	if (options.isCancelled?.()) { throw new AnalysisCancelled(); }
}
