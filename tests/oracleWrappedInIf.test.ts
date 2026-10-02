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

// Issue #237: wrapping a procedure's body in `If True Then` / `End If`
// changes nothing the VBE does - each wrapped oracle case was measured in
// Excel 16.0 to raise the same error - so it must change nothing the
// analyzer reports. Rules that track state (objects, collections, arrays,
// open files, error handlers) used to stop at the first block, and a finding
// inside one was lost. Every oracle case with an entry procedure is analyzed
// both ways, and its error findings compared. A line pushed past 1023
// characters by the extra indent would be a new line-too-long, rightly, so
// that code is left out.

interface OracleModule {
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
	return [...lines.slice(0, header + 1), 'If True Then', ...lines.slice(header + 1, end), 'End If', ...lines.slice(end)].join('\r\n');
}

describe('oracle cases wrapped in If True Then (issue #237)', () => {
	it('report the same errors wrapped as unwrapped', () => {
		const differences: string[] = [];
		let checked = 0;
		for (const oracleCase of corpus.cases) {
			if (!oracleCase.entryPoint) {
				continue;
			}
			const modules = modulesOf(oracleCase);
			const entry = modules.find((mod) => mod.entry) ?? modules[0];
			const source = wrapped(entry.source, oracleCase.entryPoint);
			if (source === undefined) {
				continue;
			}
			checked++;
			const before = errorCodes(modules);
			const after = errorCodes(modules.map((mod) => (mod === entry ? { ...mod, source } : mod)));
			if (before.join('\n') !== after.join('\n')) {
				differences.push(`${oracleCase.id}: [${before.join(', ')}] -> [${after.join(', ')}]`);
			}
		}
		expect(checked).toBeGreaterThan(1000);
		expect(differences).toEqual([]);
		// Two analyses of every oracle case: about 17 s alone here with 6,400
		// cases, and 33 to 63 s on a CI runner beside the suite. The oracle
		// grows with every measured issue, so the limit leaves room.
	}, 240_000);
});
