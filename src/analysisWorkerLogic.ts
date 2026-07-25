// Pure request handler for the analysis worker: all state and behavior live
// here (unit-testable in-process); the worker entry is a thin parentPort shim.
// Must stay free of any `vscode` import - it runs on a worker thread.

import { analyzeVbaModuleSource } from './vbaModuleAnalysis';
import type { ModuleRulesIncrementalState } from './analyzer';
import {
	buildVbaProjectIndex,
	projectAnalysisOptionsForModule,
	projectProcedureSignatures,
	type VbaProjectAnalysisOptions,
} from './vbaProjectAnalysis';
import type {
	AnalysisWorkerRequest,
	AnalysisWorkerResponse,
	WorkerSeedModule,
} from './analysisWorkerProtocol';
import type { ModuleSymbolKind } from './analyzer/symbols/symbolModel';
import type { EventHandlerDocumentType } from './analyzer';
import type { DiagnosticSeverityOverrides } from './analyzer/diagnostics/analysisContext';

interface WorkbookState {
	generation: number;
	modules: WorkerSeedModule[];
	project: ReturnType<typeof buildVbaProjectIndex>;
	procedures: ReturnType<typeof projectProcedureSignatures>;
	/** Memoized per analyzed module name (lowercased). */
	optionsByModule: Map<string, VbaProjectAnalysisOptions>;
}

export class AnalysisWorkerState {
	private readonly _workbooks = new Map<string, WorkbookState>();
	private readonly _incrementalByDoc = new Map<string, ModuleRulesIncrementalState>();

	handle(request: AnalysisWorkerRequest): AnalysisWorkerResponse | undefined {
		switch (request.kind) {
			case 'seed': {
				const project = buildVbaProjectIndex(request.modules.map((m) => ({
					moduleName: m.moduleName,
					source: m.source,
					type: m.type,
					documentType: m.documentType as EventHandlerDocumentType | undefined,
				})));
				this._workbooks.set(request.workbookKey, {
					generation: request.generation,
					modules: request.modules,
					project,
					procedures: projectProcedureSignatures(project),
					optionsByModule: new Map(),
				});
				return undefined;
			}
			case 'forget': {
				this._incrementalByDoc.delete(request.docKey);
				return undefined;
			}
			case 'analyze': {
				try {
					return this._analyze(request);
				} catch (err) {
					return {
						kind: 'error',
						requestId: request.requestId,
						docKey: request.docKey,
						message: err instanceof Error ? err.message : String(err),
					};
				}
			}
		}
	}

	private _analyze(
		request: Extract<AnalysisWorkerRequest, { kind: 'analyze' }>,
	): AnalysisWorkerResponse {
		let projectOptions: VbaProjectAnalysisOptions = {};
		if (request.workbookKey !== undefined) {
			const workbook = this._workbooks.get(request.workbookKey);
			if (!workbook || (request.generation !== undefined && workbook.generation !== request.generation)) {
				return {
					kind: 'needSeed',
					requestId: request.requestId,
					docKey: request.docKey,
					workbookKey: request.workbookKey,
				};
			}
			const moduleKey = request.moduleName.toLowerCase();
			let options = workbook.optionsByModule.get(moduleKey);
			if (!options) {
				options = projectAnalysisOptionsForModule(
					workbook.project,
					request.moduleName,
					workbook.procedures,
				);
				workbook.optionsByModule.set(moduleKey, options);
			}
			projectOptions = options;
		}

		const fingerprint = [
			request.workbookKey ?? '',
			request.generation ?? -1,
			JSON.stringify(request.severityOverrides ?? null),
			request.moduleType ?? '',
			request.moduleKind ?? '',
			request.documentType ?? '',
		] as const;

		const result = analyzeVbaModuleSource({
			source: request.source,
			moduleName: request.moduleName,
			moduleType: request.moduleType,
			moduleKind: request.moduleKind as ModuleSymbolKind | undefined,
			documentType: request.documentType as EventHandlerDocumentType | undefined,
			severityOverrides: request.severityOverrides as DiagnosticSeverityOverrides | undefined,
			...projectOptions,
			activeIncompleteExpressionOffset: request.activeIncompleteExpressionOffset,
			rulesIncremental: {
				state: this._incrementalByDoc.get(request.docKey),
				fingerprint,
			},
		});
		if (result.rulesIncrementalState) {
			this._incrementalByDoc.set(request.docKey, result.rulesIncrementalState);
		}
		return {
			kind: 'result',
			requestId: request.requestId,
			docKey: request.docKey,
			diagnostics: result.diagnostics,
			incrementalMode: result.rulesIncrementalMode,
		};
	}
}
