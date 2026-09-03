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
	/**
	 * True when the module carries `Attribute VB_PredeclaredId = True`, giving
	 * it a default instance so its own name is usable as a value. Absent means
	 * the attribute header was not read, never "no".
	 */
	predeclaredId?: boolean;
}

export type AnalysisWorkerRequest =
	| {
		kind: 'seed';
		projectKey: string;
		generation: number;
		modules: WorkerSeedModule[];
		/**
		 * The project's own `#If` constants, raw as the VBE property holds them
		 * (`Name = Value : Name = Value`). Raw rather than parsed so
		 * `parseProjectConditionalConstants` stays the one implementation of the
		 * format and a consumer cannot disagree with it about what `-1` means.
		 *
		 * Without them every `#If MY_FLAG` is undecidable and both arms are
		 * analyzed, so a consumer that seeds them and one that does not get
		 * different answers for the same file
		 * (github.com/WilliamSmithEdward/xlide_vscode/issues/63).
		 */
		conditionalConstants?: string;
	}
	| {
		kind: 'analyze';
		requestId: number;
		docKey: string;
		/** Present for project-backed modules; undefined analyzes standalone. */
		projectKey?: string;
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
		 * Which Office host the module belongs to (`excel`, `word`,
		 * `powerpoint`, `access`, ...). Absent means Excel; a named host with
		 * no model yet asserts no host knowledge at all (issue #24).
		 */
		host?: string;
		/**
		 * Designer-declared members of the analyzed module, overriding whatever
		 * the seed carried. A host whose designer state changes between seeds
		 * (and a host that analyzes without seeding a project at all) sends
		 * them here.
		 */
		implicitMembers?: WorkerImplicitMember[];
		/**
		 * The host type the module's designer makes it (`VB.MDIForm`,
		 * `VB.UserControl`). Its members are the module's own, so `Me` reaches
		 * them and a bare call binds to them.
		 */
		designerClass?: string;
	}
	| { kind: 'forget'; docKey: string };

export type AnalysisWorkerResponse =
	| {
		kind: 'result';
		requestId: number;
		docKey: string;
		diagnostics: VbaModuleAnalysisDiagnostic[];
		/** Findings silenced by suppression directives; project analysis reports them separately. */
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
		/** The worker has no (or a stale) seed for this project; reseed and retry. */
		kind: 'needSeed';
		requestId: number;
		docKey: string;
		projectKey: string;
	};
