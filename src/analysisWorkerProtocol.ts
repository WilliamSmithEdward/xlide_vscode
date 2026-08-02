// Message protocol between the extension host and the analysis worker thread.
// Everything crossing the boundary must be structured-cloneable plain data.

import type { VbaModuleAnalysisDiagnostic } from './vbaModuleAnalysis';

export interface WorkerSeedModule {
	moduleName: string;
	source: string;
	type?: string;
	documentType?: string;
}

export type AnalysisWorkerRequest =
	| {
		kind: 'seed';
		workbookKey: string;
		generation: number;
		modules: WorkerSeedModule[];
	}
	| {
		kind: 'analyze';
		requestId: number;
		docKey: string;
		/** Present for workbook-backed modules; undefined analyzes standalone. */
		workbookKey?: string;
		/** The cross-module generation the seed must match for incremental reuse. */
		generation?: number;
		source: string;
		moduleName: string;
		moduleType?: string;
		moduleKind?: string;
		documentType?: string;
		severityOverrides?: Record<string, string>;
		activeIncompleteExpressionOffset?: number;
	}
	| { kind: 'forget'; docKey: string };

export type AnalysisWorkerResponse =
	| {
		kind: 'result';
		requestId: number;
		docKey: string;
		diagnostics: VbaModuleAnalysisDiagnostic[];
		/** Findings silenced by suppression directives; workbook analysis reports them separately. */
		suppressedDiagnostics: VbaModuleAnalysisDiagnostic[];
		incrementalMode?: 'full' | 'incremental';
	}
	| {
		kind: 'error';
		requestId: number;
		docKey: string;
		message: string;
	}
	| {
		/** The worker has no (or a stale) seed for this workbook; reseed and retry. */
		kind: 'needSeed';
		requestId: number;
		docKey: string;
		workbookKey: string;
	};
