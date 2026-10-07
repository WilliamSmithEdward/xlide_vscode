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
import { parseProjectConditionalConstants, type EventHandlerDocumentType } from './analyzer';
import type { DiagnosticSeverityOverrides } from './analyzer/diagnostics/analysisContext';
import { AnalysisCancelled, checkAnalysisCancellation } from './analyzer/diagnostics/analysisCancellation';

interface ProjectState {
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
	/**
	 * The exported-signature half of the surface digest. Identical for every
	 * module in the project, so it is folded once per seed rather than once
	 * per module - folding it per module cost more than building the options
	 * it summarises (measured 8.6 ms/module on a 40-module project).
	 */
	proceduresDigest: string;
}

export class AnalysisWorkerState {
	private readonly _workbooks = new Map<string, ProjectState>();
	private readonly _incrementalByDoc = new Map<string, ModuleRulesIncrementalState>();
	// The local and full live passes often request the identical snapshot.
	// Incremental rules still repeat all eager module checks, so keep one exact
	// completed result per document, separately from procedure-level reuse.
	private readonly _completedByDoc = new Map<string, {
		projectKey: string | undefined;
		source: string;
		fingerprint: readonly unknown[];
		activeIncompleteExpressionOffset: number | undefined;
		result: Extract<AnalysisWorkerResponse, { kind: 'result' }>;
	}>();

	handle(request: AnalysisWorkerRequest): AnalysisWorkerResponse | undefined {
		switch (request.kind) {
			case 'seed': {
				for (const [key, completed] of this._completedByDoc) {
					if (completed.projectKey === request.projectKey) { this._completedByDoc.delete(key); }
				}
				const project = buildVbaProjectIndex(request.modules.map((m) => ({
					moduleName: m.moduleName,
					source: m.source,
					type: m.type,
					documentType: m.documentType as EventHandlerDocumentType | undefined,
					// Folded into the form's member surface, so a CALLER module's
					// qualified `EntryForm.NameBox` resolves - not only the form's
					// own code-behind (#22).
					implicitMembers: m.implicitMembers,
					predeclaredId: m.predeclaredId,
					designerClass: m.designerClass,
				})), undefined, {
					// Both halves must be built under the same constants: giving
					// them to the rules alone would leave a branch dropped from
					// the symbols still being analyzed (issues/63).
					conditionalCompilation: {
						projectConstants: parseProjectConditionalConstants(request.conditionalConstants),
					},
				});
				this._workbooks.set(request.projectKey, {
					generation: request.generation,
					modules: request.modules,
					project,
					procedures: projectProcedureSignatures(project),
					optionsByModule: new Map(),
					surfaceDigestByModule: new Map(),
					proceduresDigest: '',
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
				this._completedByDoc.delete(request.docKey);
				return undefined;
			}
			case 'analyze': {
				try {
					return this._analyze(request);
				} catch (err) {
					if (err instanceof AnalysisCancelled) { return { kind: 'cancelled', requestId: request.requestId, docKey: request.docKey }; }
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
		const cancellation = { isCancelled: request.cancellationSignal ? () => Atomics.load(request.cancellationSignal!, 0) !== 0 : undefined };
		checkAnalysisCancellation(cancellation);
		let projectOptions: VbaProjectAnalysisOptions = {};
		let seededImplicitMembers: WorkerImplicitMember[] | undefined;
		let surfaceDigest = '';
		if (request.projectKey !== undefined) {
			const state = this._workbooks.get(request.projectKey);
			if (!state || (request.generation !== undefined && state.generation !== request.generation)) {
				return {
					kind: 'needSeed',
					requestId: request.requestId,
					docKey: request.docKey,
					projectKey: request.projectKey,
				};
			}
			const moduleKey = request.moduleName.toLowerCase();
			let options = state.optionsByModule.get(moduleKey);
			if (!options) {
				options = projectAnalysisOptionsForModule(
					state.project,
					request.moduleName,
					state.procedures,
				);
				state.optionsByModule.set(moduleKey, options);
				if (state.proceduresDigest === '') {
					state.proceduresDigest = exportedProcedureDigest(options.projectProcedures);
				}
				// The per-module half covers only what visibility filtering makes
				// specific to this module; the shared half rides along (issue #42).
				state.surfaceDigestByModule.set(
					moduleKey,
					`${state.proceduresDigest}/${moduleSurfaceDigest(options)}`,
				);
			}
			projectOptions = options;
			surfaceDigest = state.surfaceDigestByModule.get(moduleKey) ?? '';
			seededImplicitMembers = state.implicitMembersByModule.get(moduleKey);
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
			request.errorsOnly === true,
			request.moduleName,
			request.projectKey ?? '',
			// The project surface this module consumes, NOT the seed's
			// generation: a re-seed with unchanged cross-module content keeps
			// every module's incremental state, while a changed signature,
			// type or member surface still invalidates it (issue #42).
			surfaceDigest,
			JSON.stringify(request.severityOverrides ?? null),
			request.moduleType ?? '',
			request.moduleKind ?? '',
			request.documentType ?? '',
			// A host change re-types every host lookup in the module, and so
			// does a reference, which brings another application's types in.
			request.host ?? '',
			(request.referencedHosts ?? []).join('+'),
			request.referencedLibraries === undefined ? '' : `refs:${request.referencedLibraries.join('+')}`,
			// Editing the designer changes diagnostics without changing a line
			// of code, so incremental reuse has to see the control list.
			JSON.stringify(implicitMembers ?? null),
			// The class the designer makes the module is part of its scope.
			request.designerClass ?? '',
			// Saving the workbook with a sheet added or renamed changes what
			// `ThisWorkbook.Sheets("x")` reaches without changing a line of code.
			JSON.stringify(request.workbookSheets ?? null),
		] as const;
		const completed = this._completedByDoc.get(request.docKey);
		if (completed && completed.source === request.source
			&& completed.activeIncompleteExpressionOffset === request.activeIncompleteExpressionOffset
			&& completed.fingerprint.length === fingerprint.length
			&& completed.fingerprint.every((value, i) => Object.is(value, fingerprint[i]))) {
			return { ...completed.result, requestId: request.requestId };
		}

		const result = analyzeVbaModuleSource({
			...cancellation,
			errorsOnly: request.errorsOnly,
			source: request.source,
			moduleName: request.moduleName,
			moduleType: request.moduleType,
			host: request.host,
			referencedHosts: request.referencedHosts,
			referencedLibraries: request.referencedLibraries,
			designerClass: request.designerClass,
			workbookSheets: request.workbookSheets,
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
		checkAnalysisCancellation(cancellation);
		if (result.rulesIncrementalState) {
			this._incrementalByDoc.set(request.docKey, result.rulesIncrementalState);
		}
		const response: Extract<AnalysisWorkerResponse, { kind: 'result' }> = {
			kind: 'result',
			requestId: request.requestId,
			docKey: request.docKey,
			diagnostics: result.diagnostics,
			suppressedDiagnostics: result.suppressedDiagnostics,
			incrementalMode: result.rulesIncrementalMode,
			...(result.analysisFailures ? { analysisFailures: result.analysisFailures } : {}),
		};
		if (!result.analysisFailures?.length) {
			this._completedByDoc.set(request.docKey, {
				projectKey: request.projectKey,
				source: request.source, fingerprint,
				activeIncompleteExpressionOffset: request.activeIncompleteExpressionOffset,
				result: response,
			});
		} else {
			this._completedByDoc.delete(request.docKey);
		}
		return response;
	}
}

/**
 * Order-independent content fold, used by both halves of the surface digest.
 *
 * Two accumulators rather than one: a 32-bit sum collides far too readily for
 * a key that decides whether a stale diagnostic is allowed to survive, so a
 * second, differently-weighted accumulator widens the state to 64 bits at the
 * cost of one multiply per item.
 */
class SurfaceFold {
    private _items = 0;
    private _sum = 0;
    private _mixed = 0;

    add(text: string): void {
        let hash = 0x811c9dc5;
        for (let i = 0; i < text.length; i += 1) {
            hash ^= text.charCodeAt(i);
            hash = Math.imul(hash, 0x01000193);
        }
        const value = hash >>> 0;
        this._items += 1;
        this._sum = (this._sum + value) >>> 0;
        this._mixed = (this._mixed ^ Math.imul(value ^ text.length, 0x9e3779b1)) >>> 0;
    }

    toString(): string {
        return `${this._items}:${this._sum.toString(36)}:${this._mixed.toString(36)}`;
    }
}

/**
 * The project's exported procedure signatures - what a call site in ANY module
 * is checked against. Identical for every module, so the worker folds it once
 * per seed (issue #42).
 */
function exportedProcedureDigest(
    projectProcedures: VbaProjectAnalysisOptions['projectProcedures'],
): string {
    const fold = new SurfaceFold();
    for (const [key, signatures] of projectProcedures ?? []) {
        for (const signature of signatures) {
            fold.add(`${key}:${signature.moduleName ?? ''}.${signature.name}:${signature.kind}:${signature.returnType ?? ''}:${
                signature.params
                    .map((param) => `${param.name}|${param.type ?? ''}|${param.optional ? '1' : '0'}|${param.paramArray ? '1' : '0'}`)
                    .join(',')
            }`);
        }
    }
    return fold.toString();
}

/**
 * The half of the cross-module surface that visibility filtering makes specific
 * to one module: the names it can see, the types visible to it, and the
 * source-backed member surfaces it resolves against.
 *
 * Order-independent by construction, because the project index builds these by
 * walking modules and a re-seed must not invalidate on iteration order alone.
 */
function moduleSurfaceDigest(options: VbaProjectAnalysisOptions): string {
    const fold = new SurfaceFold();
    for (const set of [
        options.knownProcedures,
        options.knownIdentifiers,
        options.knownNonTypeNames,
        options.hiddenTypeNames,
        // A string in another module can name one of this module's private
        // procedures, so a new mention there has to clear the finding here.
        options.projectStringLiteralWords,
    ]) {
        for (const name of set ?? []) {
            fold.add(name);
        }
    }
    for (const type of options.projectTypes ?? []) {
        fold.add(`${type.name}:${type.kind}:${type.moduleName ?? ''}`);
    }
    for (const surface of options.projectClassMembers ?? []) {
        // The default-instance bit changes which bare qualifiers are legal, so a
        // re-seed that answers it differently must not reuse the old analysis. All
        // THREE states are distinct: unknown is not the same claim as false (#47).
        fold.add(`${surface.name}:${surface.kind}:${surface.exhaustive === true ? '1' : '0'}`
            + `:${surface.predeclaredId === undefined ? '?' : surface.predeclaredId ? '1' : '0'}`);
        for (const member of surface.members) {
            fold.add(`${surface.name}.${member.name}:${member.kind}:${member.returns ?? ''}`);
        }
    }
    // Another module adding or naming a sheet decides whether this one's
    // `ThisWorkbook.Sheets("x")` is reported (issue #229).
    const sheets = options.projectSheetChanges;
    if (sheets) {
        fold.add(`sheets:${sheets.addsSheets ? '1' : '0'}${sheets.assignsComputedName ? '1' : '0'}`);
        for (const name of sheets.namesAssigned) {
            fold.add(`sheet:${name}`);
        }
    }
    return fold.toString();
}
