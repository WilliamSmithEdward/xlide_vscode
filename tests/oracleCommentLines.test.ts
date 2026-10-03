import { readFileSync } from 'fs';
import { join } from 'path';
import { describe, expect, it } from 'vitest';
import { analyzeVbaModuleSource } from '../src/vbaModuleAnalysis';
import {
	buildVbaProjectIndex,
	effectiveModuleKind,
	projectAnalysisOptionsForModule,
	projectProcedureSignatures,
} from '../src/vbaProjectAnalysis';

// Issue #249: a comment line is no statement. A `' note` line after every
// complete line of an oracle case changes nothing the VBE does - the cases
// the flow rules used to misread were each measured in Excel 16.0 with and
// without the comments - so it must change nothing the analyzer reports.
// The comment line used to be unreachable code after Exit, the line a GoSub
// target was fallen into from, and the statement that hid a handler's next
// line from the rules. Every finding of every severity is compared, at the
// line it was on before the comments went in. No comment goes in front of
// an Attribute line, whose place is fixed.

interface OracleModule {
	name: string;
	type?: string;
	source: string;
}

interface OracleCase {
	id: string;
	moduleName?: string;
	source?: string;
	modules?: OracleModule[];
}

const corpus: { cases: OracleCase[] } = JSON.parse(
	readFileSync(join(process.cwd(), 'syntax_corpus', 'oracle', 'vbe_oracle_cases.json'), 'utf8'),
);

function modulesOf(oracleCase: OracleCase): OracleModule[] {
	return oracleCase.modules ?? [{ name: oracleCase.moduleName ?? 'Module1', type: 'standard', source: oracleCase.source ?? '' }];
}

/** The source with a comment after each complete line, and the original line of every line of it (0 for a comment). */
function commented(source: string): { source: string; origin: number[] } {
	const lines = source.split(/\r\n|\n/);
	const out: string[] = [];
	const origin: number[] = [];
	lines.forEach((line, i) => {
		out.push(line);
		origin.push(i + 1);
		const next = lines[i + 1];
		if (next !== undefined && !/ _\s*$/.test(line) && !/^\s*Attribute\s/i.test(next)) {
			out.push("' note");
			origin.push(0);
		}
	});
	return { source: out.join('\r\n'), origin };
}

/** Every finding as `module: code @ line`, sorted, with lines mapped through `origins`. */
function findings(modules: readonly OracleModule[], origins?: ReadonlyMap<string, number[]>): string[] {
	const inputs = modules.map((mod) => ({ moduleName: mod.name, type: mod.type, source: mod.source }));
	const project = buildVbaProjectIndex(inputs, undefined, { conditionalCompilation: { projectConstants: {} } });
	const procedures = projectProcedureSignatures(project);
	const out: string[] = [];
	for (const mod of inputs) {
		const { diagnostics } = analyzeVbaModuleSource({
			source: mod.source,
			moduleName: mod.moduleName,
			moduleKind: effectiveModuleKind(mod),
			...projectAnalysisOptionsForModule(project, mod.moduleName, procedures),
		});
		for (const diagnostic of diagnostics) {
			const line = mod.source.slice(0, diagnostic.span.start).split('\n').length;
			out.push(`${mod.moduleName}: ${diagnostic.code} @ ${origins ? origins.get(mod.moduleName)![line - 1] : line}`);
		}
	}
	return out.sort();
}

describe('oracle cases with a comment after every line (issue #249)', () => {
	it('report the same findings at the same lines', () => {
		const differences: string[] = [];
		for (const oracleCase of corpus.cases) {
			const modules = modulesOf(oracleCase);
			const origins = new Map<string, number[]>();
			const withComments = modules.map((mod) => {
				const { source, origin } = commented(mod.source);
				origins.set(mod.name, origin);
				return { ...mod, source };
			});
			const before = findings(modules);
			const after = findings(withComments, origins);
			if (before.join('\n') !== after.join('\n')) {
				differences.push(`${oracleCase.id}: [${before.join(', ')}] -> [${after.join(', ')}]`);
			}
		}
		expect(corpus.cases.length).toBeGreaterThan(2000);
		expect(differences).toEqual([]);
		// Two analyses of every oracle case: about 30 s alone here with 8,600
		// cases, and 115 s on a CI runner beside the suite on 2026-10-03, which
		// the old 120 s limit then failed. The oracle grows with every measured
		// issue, so the limit leaves room, as oracleWrappedInIf's does.
	}, 240_000);
});
