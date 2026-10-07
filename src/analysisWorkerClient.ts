// Extension-host client for the analysis worker thread. Spawns lazily, tracks
// per-project seed generations, and serializes requests. Timeouts restart the
// worker without retrying stuck work on the host; startup/crash failures retain
// the caller's ordinary fallback path.

import { Worker } from 'worker_threads';
import { AnalysisWorkerTimeoutError } from './analysisWorkerErrors';
import * as fs from 'fs';
import type {
	AnalysisWorkerRequest,
	AnalysisWorkerResponse,
	WorkerImplicitMember,
	WorkerSeedModule,
} from './analysisWorkerProtocol';
import type { VbaModuleAnalysisDiagnostic, VbaModuleAnalysisFailure } from './vbaModuleAnalysis';
import type { WorkbookSheetInfo } from './analyzer/symbols/sheetChanges';

export interface WorkerAnalyzeRequest {
	errorsOnly?: boolean;
	/** Live diagnostics replace queued snapshots and cancel superseded work for this document. */
	latestOnly?: boolean;
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
	/** Host tokens for the libraries the project references, if any. */
	referencedHosts?: readonly string[];
	/** The names of the libraries the project references, when known. */
	referencedLibraries?: readonly string[];
	/** Designer-declared members of this module, when the caller knows them. */
	implicitMembers?: WorkerImplicitMember[];
	/** The host type the module's designer makes it, when the caller knows it. */
	designerClass?: string;
	/** The workbook's sheets as saved, when the container is a workbook. */
	workbookSheets?: readonly WorkbookSheetInfo[];
}

export interface WorkerAnalyzeResult {
	diagnostics: VbaModuleAnalysisDiagnostic[];
	suppressedDiagnostics: VbaModuleAnalysisDiagnostic[];
	incrementalMode?: 'full' | 'incremental';
	analysisFailures?: VbaModuleAnalysisFailure[];
}

interface PendingRequest {
	cancellation?: Int32Array;
	resolve: (result: WorkerAnalyzeResult) => void;
	reject: (err: Error) => void;
	request: WorkerAnalyzeRequest;
	retried: boolean;
	seedProvider?: () => WorkerSeedModule[];
	watchdog: ReturnType<typeof setTimeout>;
}

// A worker stuck in a pathological loop emits no error or exit event, so a
// request that never answers would otherwise hang its promise forever with
// `available` still true - live diagnostics stall for the session and the
// request never settles. A timeout fails that run and restarts the worker;
// callers must not repeat the same potentially stuck analysis in-host.
// Only the dispatched request is timed; queue waiting is not analysis time.
const WORKER_REQUEST_TIMEOUT_MS = 30_000;

export class AnalysisWorkerClient {
	private _worker: Worker | undefined;
	private _failed = false;
	private _nextRequestId = 1;
	private readonly _queue: Omit<PendingRequest, 'watchdog'>[] = [];
	private readonly _pending = new Map<number, PendingRequest>();
	private readonly _seededGenerations = new Map<string, number>();

	private readonly _seedProviders = new Map<string, () => WorkerSeedModule[]>();

	constructor(
		private readonly _workerPath: string,
		private readonly _log?: (line: string) => void,
		private readonly _requestTimeoutMs: number = WORKER_REQUEST_TIMEOUT_MS,
	) {}

	/**
	 * Retains the seed provider for dispatch and a needSeed retry. Each queued
	 * request captures its provider, so a newer project cannot change its seed.
	 */
	ensureSeeded(projectKey: string, _generation: number, modules: () => WorkerSeedModule[]): void {
		this._seedProviders.set(projectKey, modules);
		// Seed immediately before dispatch so queued generations cannot interfere.
		this._ensureWorker();
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
		return new Promise<WorkerAnalyzeResult>((resolve, reject) => {
			if (request.latestOnly) {
				this._cancelLiveDocument(request.docKey);
				for (let i = this._queue.length - 1; i >= 0; i--) {
					const queued = this._queue[i];
					if (queued.request.latestOnly && queued.request.docKey === request.docKey) {
						this._queue.splice(i, 1);
						const error = new Error('Analysis snapshot superseded.');
						error.name = 'AnalysisSnapshotSuperseded';
						queued.reject(error);
					}
				}
			}
			this._queue.push({ resolve, reject, request, retried: false,
				cancellation: request.latestOnly && typeof SharedArrayBuffer !== 'undefined' ? new Int32Array(new SharedArrayBuffer(4)) : undefined,
				seedProvider: request.projectKey ? this._seedProviders.get(request.projectKey) : undefined });
			this._dispatchNext();
		});
	}

	private _dispatchNext(): void {
		const worker = this._worker;
		if (this._failed || this._pending.size > 0 || !worker) { return; }
		// A bad seed or non-cloneable request rejects only that request. In
		// particular, dispatch from a response event must never throw into the host
		// or strand a shifted request without a watchdog.
		while (this._queue.length > 0) {
			const next = this._queue.shift()!;
			if (this._submit(worker, next)) { return; }
		}
	}

	private _submit(worker: Worker, next: Omit<PendingRequest, 'watchdog'>): boolean {
		const requestId = this._nextRequestId++;
		try {
			if (next.request.projectKey !== undefined && next.request.generation !== undefined) {
				this._postSeed(worker, next.request.projectKey, next.request.generation, next.seedProvider);
			}
			this._track(requestId, next);
			worker.postMessage({ kind: 'analyze', requestId, ...next.request,
				...(next.cancellation ? { cancellationSignal: next.cancellation } : {}) } satisfies AnalysisWorkerRequest);
			return true;
		} catch (err) {
			const pending = this._pending.get(requestId);
			if (pending) { clearTimeout(pending.watchdog); }
			this._pending.delete(requestId);
			next.reject(err instanceof Error ? err : new Error(String(err)));
			return false;
		}
	}

	private _track(requestId: number, base: Omit<PendingRequest, 'watchdog'>): void {
		const watchdog = setTimeout(() => {
			if (this._pending.has(requestId)) {
				this._restartAfterTimeout(requestId);
			}
		}, this._requestTimeoutMs);
		watchdog.unref?.();
		this._pending.set(requestId, { ...base, watchdog });
	}

	private _restartAfterTimeout(requestId: number): void {
		const pending = this._pending.get(requestId);
		if (!pending || this._failed) { return; }
		clearTimeout(pending.watchdog);
		this._pending.delete(requestId);
		pending.reject(new AnalysisWorkerTimeoutError(this._requestTimeoutMs));
		// Retrying a stuck analysis on the extension host can freeze Backspace
		// and hovers. Fail this run; give queued jobs a fresh off-thread worker.
		const previous = this._worker;
		this._worker = undefined;
		this._seededGenerations.clear();
		void previous?.terminate();
		this._log?.(`Analysis request timed out after ${this._requestTimeoutMs} ms; restarting worker.`);
		if (this._ensureWorker()) { this._dispatchNext(); }
	}

	forget(docKey: string): void {
		this._cancelLiveDocument(docKey);
		for (let i = this._queue.length - 1; i >= 0; i--) {
			if (this._queue[i].request.docKey !== docKey) { continue; }
			const queued = this._queue.splice(i, 1)[0];
			const error = new Error('Analysis document forgotten.');
			error.name = 'AnalysisSnapshotSuperseded';
			queued.reject(error);
		}
		if (this._worker && !this._failed) {
			this._worker.postMessage({ kind: 'forget', docKey } satisfies AnalysisWorkerRequest);
		}
	}

	private _cancelLiveDocument(docKey: string): void {
		for (const pending of this._pending.values()) {
			if (pending.request.docKey === docKey && pending.cancellation) { Atomics.store(pending.cancellation, 0, 1); }
		}
	}

	private _postSeed(worker: Worker, projectKey: string, generation: number, provider = this._seedProviders.get(projectKey)): void {
		if (this._seededGenerations.get(projectKey) === generation) {
			return;
		}
		const modules = provider?.();
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
			worker.on('error', (err) => {
				if (worker === this._worker) { this._fail(`worker error: ${err.message}`); }
			});
			worker.on('exit', (code) => {
				if (!this._failed && worker === this._worker) {
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
		if (worker !== this._worker) { return; }
		const pending = this._pending.get(response.requestId);
		if (!pending) {
			return;
		}
		if (response.kind === 'cancelled' || pending.cancellation && Atomics.load(pending.cancellation, 0) !== 0) {
			this._pending.delete(response.requestId);
			clearTimeout(pending.watchdog);
			const error = new Error('Analysis snapshot superseded.');
			error.name = 'AnalysisSnapshotSuperseded';
			pending.reject(error);
			this._dispatchNext();
			return;
		}
		if (response.kind === 'needSeed') {
			// Reseed once and retry the same request; a second miss is an error.
			this._pending.delete(response.requestId);
			clearTimeout(pending.watchdog);
			if (pending.retried) {
				pending.reject(new Error('Analysis worker seed mismatch.'));
				this._dispatchNext();
				return;
			}
			this._seededGenerations.delete(response.projectKey);
			if (!this._submit(worker, { ...pending, retried: true })) {
				this._dispatchNext();
			}
			return;
		}
		this._pending.delete(response.requestId);
		clearTimeout(pending.watchdog);
		if (response.kind === 'error') {
			pending.reject(new Error(response.message));
			this._dispatchNext();
			return;
		}
		pending.resolve({
			diagnostics: response.diagnostics,
			suppressedDiagnostics: response.suppressedDiagnostics,
			incrementalMode: response.incrementalMode,
			...(response.analysisFailures ? { analysisFailures: response.analysisFailures } : {}),
		});
		this._dispatchNext();
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
		for (const queued of this._queue.splice(0)) { queued.reject(err); }
	}
}
