// Message protocol between the extension host and the analysis worker thread.
// Everything crossing the boundary must be structured-cloneable plain data.

import type { VbaModuleAnalysisDiagnostic } from './vbaModuleAnalysis';

/**
 * A member a module has that its own text never declares - a UserForm control,
 * declared by the designer. A host that reads the designer itself (rather than
 * a `.frm` file, which is the only place the worker can find them on its own)
 * supplies them here.
 */
export interface WorkerImplicitMember {
	name: string;
	type: string;
}

export interface WorkerSeedModule {
	moduleName: string;
	source: string;
	type?: string;
	documentType?: string;
	/**
	 * Designer-declared members of THIS module. An empty array asserts the
	 * module has none; omit the field to leave the worker's own `.frm` header
	 * parse in charge.
	 */
	implicitMembers?: WorkerImplicitMember[];
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
		/**
		 * Designer-declared members of the analyzed module, overriding whatever
		 * the seed carried. A host whose designer state changes between seeds
		 * (and a host that analyzes without seeding a project at all) sends
		 * them here.
		 */
		implicitMembers?: WorkerImplicitMember[];
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
