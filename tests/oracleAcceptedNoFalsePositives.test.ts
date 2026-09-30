import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import { diagnosticMetadataForCode } from '../src/analyzer';
import { analyzeVbaModuleSource } from '../src/vbaModuleAnalysis';
import {
	buildVbaProjectIndex,
	effectiveModuleKind,
	projectAnalysisOptionsForModule,
	projectProcedureSignatures,
	type VbaProjectModuleInput,
} from '../src/vbaProjectAnalysis';

// Corpus-wide no-false-positive sweep: every ACCEPTED oracle case is code the
// Excel/VBE oracle verified as valid, so no constrained diagnostic may fire on
// it. Compile-error diagnostics are constrained by every accepted case; a
// deterministic-runtime-error / runtime-risk diagnostic is only constrained by
// a case verified at RUNTIME (a runtime diagnostic on compile-only-verified
// code flags a real runtime fault, not a false positive); style-policy
// diagnostics (Option Explicit, mismatched End keyword, ...) are advisory and
// never constrained. Each module goes through analyzeVbaModuleSource with the
// full cross-module project context its case provides, as in the extension,
// so a runtime error that an active On Error Resume Next handles is judged
// the way the user sees it (issue #199).
//
// This is the gate the juxtaposed-values regression slipped past: the rule
// flagged the oracle-accepted control suffix_long_amp_glued_concat_accepted,
// and only a targeted per-rule test existed to notice.

interface OracleModule {
	name: string;
	type?: string;
	source: string;
}

interface OracleCase {
	id: string;
	expected: 'accepted' | 'rejected' | 'observe';
	evidencePhase?: 'compile' | 'runtime';
	moduleName?: string;
	source?: string;
	modules?: OracleModule[];
}

const corpus: { cases: OracleCase[] } = JSON.parse(
	readFileSync(join(process.cwd(), 'syntax_corpus', 'oracle', 'vbe_oracle_cases.json'), 'utf8'),
);

function caseModules(oracleCase: OracleCase): OracleModule[] {
	if (oracleCase.modules) {
		return oracleCase.modules;
	}
	return [{
		name: oracleCase.moduleName ?? 'Module1',
		type: 'standard',
		source: oracleCase.source ?? '',
	}];
}

const RUNTIME_KINDS = new Set(['deterministic-runtime-error', 'runtime-risk']);

/** True when the accepted case asserts the given code must NOT fire on it. */
function constrainedByAcceptedCase(code: string, evidencePhase: string | undefined): boolean {
	const kind = diagnosticMetadataForCode(code)?.diagnosticKind;
	if (kind === 'compile-error') {
		return true;
	}
	if (kind && RUNTIME_KINDS.has(kind)) {
		return evidencePhase === 'runtime';
	}
	return false;
}

const accepted = corpus.cases.filter((oracleCase) => oracleCase.expected === 'accepted');

describe('oracle accepted cases - corpus-wide no-false-positive sweep', () => {
	it('has accepted cases to sweep', () => {
		expect(accepted.length).toBeGreaterThan(100);
	});

	it.each(accepted.map((oracleCase) => [oracleCase.id, oracleCase] as const))(
		'no constrained diagnostic fires on %s',
		(_id, oracleCase) => {
			const modules: VbaProjectModuleInput[] = caseModules(oracleCase).map((mod) => ({
				moduleName: mod.name,
				type: mod.type,
				source: mod.source,
			}));
			// The extension always supplies the project's conditional
			// constants, and an Excel case has none: a name no #Const defines
			// is then provably undefined, as the VBE treats it (issue #208).
			const project = buildVbaProjectIndex(modules, undefined, { conditionalCompilation: { projectConstants: {} } });
			const procedures = projectProcedureSignatures(project);
			const falsePositives: string[] = [];
			for (const mod of modules) {
				const { diagnostics } = analyzeVbaModuleSource({
					source: mod.source,
					moduleName: mod.moduleName,
					moduleKind: effectiveModuleKind(mod),
					...projectAnalysisOptionsForModule(project, mod.moduleName, procedures),
				});
				for (const diagnostic of diagnostics) {
					if (constrainedByAcceptedCase(diagnostic.code ?? '', oracleCase.evidencePhase)) {
						falsePositives.push(`${mod.moduleName}: ${diagnostic.code}: ${diagnostic.message}`);
					}
				}
			}
			expect(falsePositives).toEqual([]);
		},
	);
});
