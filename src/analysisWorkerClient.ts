// Extension-host client for the analysis worker thread. Spawns lazily, tracks
// per-workbook seed generations, and degrades permanently to the caller's
// in-host fallback path when the worker cannot start or dies: analysis
// correctness never depends on the worker being alive.

import { Worker } from 'worker_threads';
import * as fs from 'fs';
import type {
	AnalysisWorkerRequest,
	AnalysisWorkerResponse,
	WorkerSeedModule,
} from './analysisWorkerProtocol';
import type { VbaModuleAnalysisDiagnostic } from './vbaModuleAnalysis';

export interface WorkerAnalyzeRequest {
	docKey: string;
	workbookKey?: string;
	generation?: number;
	source: string;
	moduleName: string;
	moduleType?: string;
	moduleKind?: string;
	documentType?: string;
	severityOverrides?: Record<string, string>;
	activeIncompleteExpressionOffset?: number;
}

export interface WorkerAnalyzeResult {
	diagnostics: VbaModuleAnalysisDiagnostic[];
	incrementalMode?: 'full' | 'incremental';
}

interface PendingRequest {
	resolve: (result: WorkerAnalyzeResult) => void;
	reject: (err: Error) => void;
	request: WorkerAnalyzeRequest;
	retried: boolean;
}

export class AnalysisWorkerClient {
	private _worker: Worker | undefined;
	private _failed = false;
	private _nextRequestId = 1;
	private readonly _pending = new Map<number, PendingRequest>();
	private readonly _seededGenerations = new Map<string, number>();

	private readonly _seedProviders = new Map<string, () => WorkerSeedModule[]>();

	constructor(
		private readonly _workerPath: string,
		private readonly _log?: (line: string) => void,
	) {}

	/**
	 * Ensures the worker holds this workbook's module sources at `generation`.
	 * The provider is retained so a needSeed round-trip can reseed on its own.
	 */
	ensureSeeded(workbookKey: string, generation: number, modules: () => WorkerSeedModule[]): void {
		this._seedProviders.set(workbookKey, modules);
		const worker = this._ensureWorker();
		if (!worker) {
			return;
		}
		this._postSeed(worker, workbookKey, generation);
	}

	/** False once the worker failed to start or died; callers use the sync path. */
	get available(): boolean {
		return !this._failed;
	}

	dispose(): void {
		this._failed = true;
		this._rejectAll(new Error('Analysis worker disposed.'));
		void this._worker?.terminate();
		this._worker = undefined;
	}

	analyze(request: WorkerAnalyzeRequest): Promise<WorkerAnalyzeResult> {
		const worker = this._ensureWorker();
		if (!worker) {
			return Promise.reject(new Error('Analysis worker unavailable.'));
		}
		if (request.workbookKey !== undefined && request.generation !== undefined) {
			this._postSeed(worker, request.workbookKey, request.generation);
		}
		return new Promise<WorkerAnalyzeResult>((resolve, reject) => {
			const requestId = this._nextRequestId++;
			this._pending.set(requestId, { resolve, reject, request, retried: false });
			worker.postMessage({ kind: 'analyze', requestId, ...request } satisfies AnalysisWorkerRequest);
		});
	}

	forget(docKey: string): void {
		if (this._worker && !this._failed) {
			this._worker.postMessage({ kind: 'forget', docKey } satisfies AnalysisWorkerRequest);
		}
	}

	private _postSeed(worker: Worker, workbookKey: string, generation: number): void {
		if (this._seededGenerations.get(workbookKey) === generation) {
			return;
		}
		const modules = this._seedProviders.get(workbookKey)?.();
		if (!modules) {
			return;
		}
		worker.postMessage({ kind: 'seed', workbookKey, generation, modules } satisfies AnalysisWorkerRequest);
		this._seededGenerations.set(workbookKey, generation);
	}

	private _ensureWorker(): Worker | undefined {
		if (this._failed) {
			return undefined;
		}
		if (this._worker) {
			return this._worker;
		}
		try {
			if (!fs.existsSync(this._workerPath)) {
				throw new Error(`worker bundle not found at ${this._workerPath}`);
			}
			const worker = new Worker(this._workerPath);
			worker.unref();
			worker.on('message', (response: AnalysisWorkerResponse) => this._onResponse(worker, response));
			worker.on('error', (err) => this._fail(`worker error: ${err.message}`));
			worker.on('exit', (code) => {
				if (!this._failed) {
					this._fail(`worker exited with code ${code}`);
				}
			});
			this._worker = worker;
			this._log?.('Analysis worker started.');
			return worker;
		} catch (err) {
			this._fail(`worker start failed: ${err instanceof Error ? err.message : String(err)}`);
			return undefined;
		}
	}

	private _onResponse(worker: Worker, response: AnalysisWorkerResponse): void {
		const pending = this._pending.get(response.requestId);
		if (!pending) {
			return;
		}
		if (response.kind === 'needSeed') {
			// Reseed once and retry the same request; a second miss is an error.
			this._pending.delete(response.requestId);
			if (pending.retried) {
				pending.reject(new Error('Analysis worker seed mismatch.'));
				return;
			}
			this._seededGenerations.delete(response.workbookKey);
			if (pending.request.workbookKey !== undefined && pending.request.generation !== undefined) {
				this._postSeed(worker, pending.request.workbookKey, pending.request.generation);
			}
			const requestId = this._nextRequestId++;
			this._pending.set(requestId, { ...pending, retried: true });
			worker.postMessage({ kind: 'analyze', requestId, ...pending.request } satisfies AnalysisWorkerRequest);
			return;
		}
		this._pending.delete(response.requestId);
		if (response.kind === 'error') {
			pending.reject(new Error(response.message));
			return;
		}
		pending.resolve({ diagnostics: response.diagnostics, incrementalMode: response.incrementalMode });
	}

	private _fail(reason: string): void {
		if (this._failed) {
			return;
		}
		this._failed = true;
		this._log?.(`Analysis worker disabled (${reason}); falling back to in-host analysis.`);
		this._rejectAll(new Error(`Analysis worker unavailable: ${reason}`));
		void this._worker?.terminate();
		this._worker = undefined;
	}

	private _rejectAll(err: Error): void {
		for (const pending of this._pending.values()) {
			pending.reject(err);
		}
		this._pending.clear();
	}
}
