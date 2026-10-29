import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import { oracleConditionalCompilation, oracleModuleInput } from '../helpers/oracleModuleInput';
import { allStructuralComparisonSamples } from '../helpers/structuralCorpus';
import { analyzeVbaModuleSource } from '../../src/vbaModuleAnalysis';
import { buildVbaProjectIndex, projectAnalysisOptionsForModule, projectProcedureSignatures } from '../../src/vbaProjectAnalysis';

// Fixtures measured in 32-bit Office (`bitness: 32`) are analyzed with Win64
// off, as their Office compiled them; analyzed as 64-bit, the same code is
// rightly refused by the 64-bit rules.

interface OracleCase {
	id: string;
	host?: string;
	bitness?: number;
	expected: string;
	source?: string;
}

const corpus: { cases: OracleCase[] } = JSON.parse(
	readFileSync(join(process.cwd(), 'syntax_corpus', 'oracle', 'vbe_oracle_cases.json'), 'utf8'),
);
const in32Bit = corpus.cases.filter((kase) => kase.bitness === 32);

function errorCodes(kase: OracleCase, asIf64Bit = false): string[] {
	const mod = oracleModuleInput({ name: 'Module1', type: 'standard', source: kase.source ?? '' });
	const environment = oracleConditionalCompilation(asIf64Bit ? {} : kase);
	const project = buildVbaProjectIndex([mod], undefined, { conditionalCompilation: environment });
	const { diagnostics } = analyzeVbaModuleSource({
		source: mod.source,
		moduleName: mod.moduleName,
		moduleKind: 'standard',
		...projectAnalysisOptionsForModule(project, mod.moduleName, projectProcedureSignatures(project)),
		...(kase.host && kase.host !== 'excel' ? { host: kase.host } : {}),
	});
	return diagnostics.filter((diagnostic) => diagnostic.severity === 'error').map((diagnostic) => diagnostic.code ?? '').sort();
}

describe('oracle fixtures measured in 32-bit Office', () => {
	it('gives the environment Win64 off for 32-bit and nothing for 64-bit', () => {
		expect(oracleConditionalCompilation({ bitness: 32 })).toEqual({ projectConstants: {}, compilerConstants: { Win64: 0 } });
		expect(oracleConditionalCompilation({ bitness: 64 })).toEqual({ projectConstants: {} });
		expect(oracleConditionalCompilation({})).toEqual({ projectConstants: {} });
	});

	it('has a first set in each of the four hosts', () => {
		const hosts = new Set(in32Bit.map((kase) => kase.host ?? 'excel'));
		expect([...hosts].sort()).toEqual(['access', 'excel', 'powerpoint', 'word']);
		expect(in32Bit.every((kase) => typeof kase.source === 'string')).toBe(true);
	});

	it.each(in32Bit.filter((kase) => kase.expected === 'accepted').map((kase) => [kase.id, kase] as const))(
		'%s has no error as 32-bit', (_id, kase) => {
			expect(errorCodes(kase)).toEqual([]);
		},
	);

	it('refuses the 32-bit Declare as 64-bit, so the environment is what keeps it clean', () => {
		const declare = in32Bit.filter((kase) => kase.id.endsWith('_declare_no_ptrsafe'));
		expect(declare).toHaveLength(4);
		for (const kase of declare) {
			expect(errorCodes(kase, true), kase.id).toEqual(['declare-missing-ptrsafe']);
		}
	});

	it('marks 32-bit samples in the structural corpus', () => {
		const samples = allStructuralComparisonSamples().filter((sample) => sample.bitness === 32);
		expect(samples.map((sample) => sample.id).sort()).toEqual(in32Bit.map((kase) => `oracle:${kase.id}`).sort());
	});
});
