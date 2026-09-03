// Extension-host client for the analysis worker thread. Spawns lazily, tracks
// per-project seed generations, and degrades permanently to the caller's
// in-host fallback path when the worker cannot start or dies: analysis
// correctness never depends on the worker being alive.

import { Worker } from 'worker_threads';
import * as fs from 'fs';
import type {
	AnalysisWorkerRequest,
	AnalysisWorkerResponse,
	WorkerImplicitMember,
	WorkerSeedModule,
} from './analysisWorkerProtocol';
import type { VbaModuleAnalysisDiagnostic } from './vbaModuleAnalysis';

export interface WorkerAnalyzeRequest {
	docKey: string;
	projectKey?: string;
	generation?: number;
	source: string;
	moduleName: string;
	moduleType?: string;
	moduleKind?: string;
	documentType?: string;
	severityOverrides?: Record<string, string>;
	activeIncompleteExpressionOffset?: number;
	/** Office host token for the module's container. Absent means Excel. */
	host?: string;
	/** Designer-declared members of this module, when the caller knows them. */
	implicitMembers?: WorkerImplicitMember[];
	/** The host type the module's designer makes it, when the caller knows it. */
	designerClass?: string;
}

export interface WorkerAnalyzeResult {
	diagnostics: VbaModuleAnalysisDiagnostic[];
	suppressedDiagnostics: VbaModuleAnalysisDiagnostic[];
	incrementalMode?: 'full' | 'incremental';
}

interface PendingRequest {
	resolve: (result: WorkerAnalyzeResult) => void;
	reject: (err: Error) => void;
	request: WorkerAnalyzeRequest;
	retried: boolean;
	watchdog: ReturnType<typeof setTimeout>;
}

// A worker stuck in a pathological loop emits no error or exit event, so a
// request that never answers would otherwise hang its promise forever with
// `available` still true - live diagnostics stall for the session and the
// in-host fallback never engages. Far beyond any legitimate analysis (the
// giant-module corpus completes in single-digit seconds), so firing means
// the worker is gone: fail it and let callers take the in-host path.
const WORKER_REQUEST_TIMEOUT_MS = 30_000;

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
		private readonly _requestTimeoutMs: number = WORKER_REQUEST_TIMEOUT_MS,
	) {}

	/**
	 * Ensures the worker holds this project's module sources at `generation`.
	 * The provider is retained so a needSeed round-trip can reseed on its own.
	 */
	ensureSeeded(projectKey: string, generation: number, modules: () => WorkerSeedModule[]): void {
		this._seedProviders.set(projectKey, modules);
		const worker = this._ensureWorker();
		if (!worker) {
			return;
		}
		this._postSeed(worker, projectKey, generation);
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
		if (request.projectKey !== undefined && request.generation !== undefined) {
			this._postSeed(worker, request.projectKey, request.generation);
		}
		return new Promise<WorkerAnalyzeResult>((resolve, reject) => {
			const requestId = this._nextRequestId++;
			this._track(requestId, { resolve, reject, request, retried: false });
			worker.postMessage({ kind: 'analyze', requestId, ...request } satisfies AnalysisWorkerRequest);
		});
	}

	private _track(requestId: number, base: Omit<PendingRequest, 'watchdog'>): void {
		const watchdog = setTimeout(() => {
			if (this._pending.has(requestId)) {
				this._fail(`request timed out after ${this._requestTimeoutMs} ms`);
			}
		}, this._requestTimeoutMs);
		watchdog.unref?.();
		this._pending.set(requestId, { ...base, watchdog });
	}

	forget(docKey: string): void {
		if (this._worker && !this._failed) {
			this._worker.postMessage({ kind: 'forget', docKey } satisfies AnalysisWorkerRequest);
		}
	}

	private _postSeed(worker: Worker, projectKey: string, generation: number): void {
		if (this._seededGenerations.get(projectKey) === generation) {
			return;
		}
		const modules = this._seedProviders.get(projectKey)?.();
		if (!modules) {
			return;
		}
		worker.postMessage({ kind: 'seed', projectKey, generation, modules } satisfies AnalysisWorkerRequest);
		this._seededGenerations.set(projectKey, generation);
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
			clearTimeout(pending.watchdog);
			if (pending.retried) {
				pending.reject(new Error('Analysis worker seed mismatch.'));
				return;
			}
			this._seededGenerations.delete(response.projectKey);
			if (pending.request.projectKey !== undefined && pending.request.generation !== undefined) {
				this._postSeed(worker, pending.request.projectKey, pending.request.generation);
			}
			const requestId = this._nextRequestId++;
			this._track(requestId, { ...pending, retried: true });
			worker.postMessage({ kind: 'analyze', requestId, ...pending.request } satisfies AnalysisWorkerRequest);
			return;
		}
		this._pending.delete(response.requestId);
		clearTimeout(pending.watchdog);
		if (response.kind === 'error') {
			pending.reject(new Error(response.message));
			return;
		}
		pending.resolve({
			diagnostics: response.diagnostics,
			suppressedDiagnostics: response.suppressedDiagnostics,
			incrementalMode: response.incrementalMode,
		});
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
			clearTimeout(pending.watchdog);
			pending.reject(err);
		}
		this._pending.clear();
	}
}
