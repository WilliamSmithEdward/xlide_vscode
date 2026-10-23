import { readFileSync } from 'fs';
import { join } from 'path';
import { describe, expect, it } from 'vitest';
import { oracleModuleInput, type OracleFixtureModule } from './oracleModuleInput';
import { oracleShardCases, oracleShardCount } from './oracleShards';
import { analyzeVbaModuleSource } from '../../src/vbaModuleAnalysis';
import {
	buildVbaProjectIndex,
	effectiveModuleKind,
	projectAnalysisOptionsForModule,
	projectProcedureSignatures,
} from '../../src/vbaProjectAnalysis';

// Issue #237: wrapping a procedure's body in `If True Then` / `End If`
// changes nothing the VBE does - each wrapped oracle case was measured in
// Excel 16.0 to raise the same error - so it must change nothing the
// analyzer reports. Rules that track state (objects, collections, arrays,
// open files, error handlers) used to stop at the first block, and a finding
// inside one was lost. Every oracle case with an entry procedure is analyzed
// both ways, and its error findings compared. A line pushed past 1023
// characters by the extra indent would be a new line-too-long, rightly, so
// that code is left out.

interface OracleModule extends OracleFixtureModule {
	name: string;
	type?: string;
	source: string;
	entry?: boolean;
}

interface OracleCase {
	id: string;
	entryPoint?: string;
	moduleName?: string;
	source?: string;
	modules?: OracleModule[];
}

const corpus: { cases: OracleCase[] } = JSON.parse(
	readFileSync(join(process.cwd(), 'syntax_corpus', 'oracle', 'vbe_oracle_cases.json'), 'utf8'),
);

function modulesOf(oracleCase: OracleCase): OracleModule[] {
	return oracleCase.modules ?? [{ name: oracleCase.moduleName ?? 'Module1', type: 'standard', source: oracleCase.source ?? '', entry: true }];
}

/** Each module's error codes, sorted, as the extension would show them. */
function errorCodes(modules: readonly OracleModule[]): string[] {
	const inputs = modules.map(oracleModuleInput);
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
			if (diagnostic.severity === 'error' && diagnostic.code !== 'line-too-long') {
				out.push(`${mod.moduleName}: ${diagnostic.code}`);
			}
		}
	}
	return out.sort();
}

/** The source with the entry procedure's body inside `If True Then` / `End If`. */
function wrapped(source: string, entry: string): string | undefined {
	const lines = source.split(/\r\n|\n/);
	const header = lines.findIndex((line) => new RegExp(`^\\s*(Public\\s+|Private\\s+)?(Function|Sub)\\s+${entry}\\b`, 'i').test(line));
	const end = lines.findIndex((line, i) => i > header && /^\s*End\s+(Function|Sub)\b/i.test(line));
	if (header < 0 || end < 0) {
		return undefined;
	}
	// A procedure ended inside an `#If`, `End Function` in each arm, has no
	// one place for the `End If` that every build compiles (issue #1315).
	let depth = 0;
	for (const line of lines.slice(header + 1, end)) {
		depth += /^\s*#\s*If\b/i.test(line) ? 1 : /^\s*#\s*End\s*If\b/i.test(line) ? -1 : 0;
	}
	if (depth > 0) {
		return undefined;
	}
	return [...lines.slice(0, header + 1), 'If True Then', ...lines.slice(header + 1, end), 'End If', ...lines.slice(end)].join('\r\n');
}

/** The case's modules, its entry module and that module's wrapped source, or undefined when the case is not checked. */
function wrappedCase(oracleCase: OracleCase): { modules: OracleModule[]; entry: OracleModule; source: string } | undefined {
	if (!oracleCase.entryPoint) {
		return undefined;
	}
	const modules = modulesOf(oracleCase);
	const entry = modules.find((mod) => mod.entry) ?? modules[0];
	const source = wrapped(entry.source, oracleCase.entryPoint);
	return source === undefined ? undefined : { modules, entry, source };
}

export function wrappedInIfShard(shard: number): void {
	describe(`oracle cases wrapped in If True Then (issue #237) - shard ${shard + 1} of ${oracleShardCount}`, () => {
		it('report the same errors wrapped as unwrapped', () => {
			const differences: string[] = [];
			for (const oracleCase of oracleShardCases(corpus.cases, shard)) {
				const checked = wrappedCase(oracleCase);
				if (checked === undefined) {
					continue;
				}
				const { modules, entry, source } = checked;
				const before = errorCodes(modules);
				const after = errorCodes(modules.map((mod) => (mod === entry ? { ...mod, source } : mod)));
				if (before.join('\n') !== after.join('\n')) {
					differences.push(`${oracleCase.id}: [${before.join(', ')}] -> [${after.join(', ')}]`);
				}
			}
			// Checked cases over the whole corpus, not this shard, so the
			// threshold is the one the sweep had before it was split.
			expect(corpus.cases.filter((oracleCase) => wrappedCase(oracleCase) !== undefined).length).toBeGreaterThan(1000);
			expect(differences).toEqual([]);
			// Two analyses of each case in a quarter of the oracle: the eight
			// shards of both sweeps took about 90 s side by side here with 9,740
			// cases on 2026-10-16. The oracle grows with every measured issue, so
			// the limit leaves room.
		}, 240_000);
	});
}
