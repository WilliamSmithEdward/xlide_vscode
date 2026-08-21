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
	WorkerImplicitMember,
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
	/**
	 * Digest of the cross-module surface each module actually consumes,
	 * memoized beside its options (issue #42). Incremental reuse keys on this
	 * rather than on the seed's generation counter: a generation changes on
	 * every re-seed, including re-seeds caused by an edit in another module
	 * that cannot affect this one, and discarding the state costs a full
	 * re-analysis of every module in the project.
	 */
	surfaceDigestByModule: Map<string, string>;
	/** Host-supplied designer members, per seeded module name (lowercased). */
	implicitMembersByModule: Map<string, WorkerImplicitMember[]>;
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
					// Folded into the form's member surface, so a CALLER module's
					// qualified `EntryForm.NameBox` resolves - not only the form's
					// own code-behind (#22).
					implicitMembers: m.implicitMembers,
				})));
				this._workbooks.set(request.workbookKey, {
					generation: request.generation,
					modules: request.modules,
					project,
					procedures: projectProcedureSignatures(project),
					optionsByModule: new Map(),
					surfaceDigestByModule: new Map(),
					implicitMembersByModule: new Map(
						request.modules
							.filter((m) => m.implicitMembers !== undefined)
							.map((m) => [m.moduleName.toLowerCase(), m.implicitMembers as WorkerImplicitMember[]]),
					),
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
		let seededImplicitMembers: WorkerImplicitMember[] | undefined;
		let surfaceDigest = '';
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
				// Computed once per module per seed, next to the options it
				// summarises, so the extra pass costs the same order as
				// building them (issue #42).
				workbook.surfaceDigestByModule.set(moduleKey, projectSurfaceDigest(options));
			}
			projectOptions = options;
			surfaceDigest = workbook.surfaceDigestByModule.get(moduleKey) ?? '';
			seededImplicitMembers = workbook.implicitMembersByModule.get(moduleKey);
		}

		// A form's controls reach the analysis from whoever actually knows them:
		// the request (a host reading the live designer), else the seed, else
		// the worker's own parse of a `.frm` header. The first two are the only
		// route for a host that seeds CodeModule text, which carries no header.
		const implicitMembers = request.implicitMembers
			?? seededImplicitMembers
			?? projectOptions.implicitMembers;
		if (implicitMembers !== projectOptions.implicitMembers) {
			projectOptions = { ...projectOptions, implicitMembers };
		}

		const fingerprint = [
			request.workbookKey ?? '',
			// The project surface this module consumes, NOT the seed's
			// generation: a re-seed with unchanged cross-module content keeps
			// every module's incremental state, while a changed signature,
			// type or member surface still invalidates it (issue #42).
			surfaceDigest,
			JSON.stringify(request.severityOverrides ?? null),
			request.moduleType ?? '',
			request.moduleKind ?? '',
			request.documentType ?? '',
			// A host change re-types every host lookup in the module.
			request.host ?? '',
			// Editing the designer changes diagnostics without changing a line
			// of code, so incremental reuse has to see the control list.
			JSON.stringify(implicitMembers ?? null),
		] as const;

		const result = analyzeVbaModuleSource({
			source: request.source,
			moduleName: request.moduleName,
			moduleType: request.moduleType,
			host: request.host,
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
			suppressedDiagnostics: result.suppressedDiagnostics,
			incrementalMode: result.rulesIncrementalMode,
		};
	}
}

/**
 * A digest of the cross-module surface a module's analysis consumes: the
 * project-visible procedure, identifier and non-type names, the visible project
 * types, and the source-backed member surfaces.
 *
 * Order-independent by construction, because the project index builds these
 * collections by walking modules and a re-seed must not invalidate on iteration
 * order alone. Content-addressed, so two seeds carrying the same project
 * surface produce the same digest and incremental state survives (issue #42).
 */
function projectSurfaceDigest(options: VbaProjectAnalysisOptions): string {
    let names = 0;
    let accumulator = 0;
    const fold = (text: string): void => {
        // FNV-1a over the string, then summed into the accumulator so the
        // result does not depend on the order items are visited in.
        let hash = 0x811c9dc5;
        for (let i = 0; i < text.length; i += 1) {
            hash ^= text.charCodeAt(i);
            hash = Math.imul(hash, 0x01000193);
        }
        accumulator = (accumulator + (hash >>> 0)) >>> 0;
        names += 1;
    };
    for (const set of [options.knownProcedures, options.knownIdentifiers, options.knownNonTypeNames]) {
        for (const name of set ?? []) {
            fold(name);
        }
    }
    for (const type of options.projectTypes ?? []) {
        fold(`${type.name}:${type.kind}:${type.moduleName ?? ''}`);
    }
    for (const surface of options.projectClassMembers ?? []) {
        fold(`${surface.name}:${surface.kind}:${surface.exhaustive === true ? '1' : '0'}`);
        for (const member of surface.members) {
            fold(`${surface.name}.${member.name}:${member.kind}:${member.returns ?? ''}`);
        }
    }
    // projectProcedures is keyed by lowercased name; the signatures are what
    // a call site in another module is checked against.
    for (const [key, signatures] of options.projectProcedures ?? []) {
        for (const signature of signatures) {
            fold(`${key}:${signature.moduleName ?? ''}.${signature.name}:${signature.kind}:${signature.returnType ?? ''}:${
                signature.params
                    .map((param) => `${param.name}|${param.type ?? ''}|${param.optional ? '1' : '0'}|${param.paramArray ? '1' : '0'}`)
                    .join(',')
            }`);
        }
    }
    return `${names}:${accumulator.toString(36)}`;
}
