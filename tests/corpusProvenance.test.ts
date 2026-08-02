import { existsSync, readdirSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import { DIAGNOSTIC_RULES, STRUCTURAL_DIAGNOSTIC_RULES } from '../src/analyzer';

type Provenance =
	| 'spec-derived'
	| 'vbe-oracle-verified'
	| 'observed-not-asserted'
	| 'pending-verification';

interface CorpusProvenanceFile {
	path: string;
	provenance: Provenance;
	notes: string;
}

interface CorpusProvenanceSidecar {
	version: number;
	allowedProvenance: Provenance[];
	files: CorpusProvenanceFile[];
}

interface OracleCase {
	id: string;
	expected: 'accepted' | 'rejected' | 'observe';
	provenance: Provenance;
	mode?: 'compile' | 'run' | 'compile_then_run';
	evidencePhase: 'compile' | 'runtime';
	diagnosticMeaning:
		| 'compile-error'
		| 'compile-valid'
		| 'runtime-error'
		| 'runtime-valid'
		| 'observation';
}

interface OracleCasesFile {
	version: number;
	cases: OracleCase[];
}

interface DiagnosticInfluenceEntry {
	code: string;
	status: Provenance;
	diagnosticKind:
		| 'compile-error'
		| 'deterministic-runtime-error'
		| 'runtime-risk'
		| 'style-policy';
	sourceOfTruth: string;
	corpusInfluence: 'none' | 'oracle-backed';
	assertedOracleCases: string[];
	observeOnlyOracleCases: string[];
	pendingLegacyCorpusFiles: string[];
	notes?: string;
}

interface DiagnosticInfluenceAudit {
	version: number;
	diagnostics: DiagnosticInfluenceEntry[];
}

const corpusRoot = join(process.cwd(), 'syntax_corpus');
const allowedProvenance: Provenance[] = [
	'spec-derived',
	'vbe-oracle-verified',
	'observed-not-asserted',
	'pending-verification',
];

function readJson<T>(path: string): T {
	return JSON.parse(readFileSync(path, 'utf8')) as T;
}

describe('syntax corpus provenance', () => {
	it('lists every corpus source file with a provenance state', () => {
		const sidecar = readJson<CorpusProvenanceSidecar>(
			join(corpusRoot, 'corpus_provenance.json'),
		);
		const allowed = new Set(allowedProvenance);

		expect(sidecar.version).toBe(1);
		expect([...sidecar.allowedProvenance].sort()).toEqual(
			[...allowedProvenance].sort(),
		);

		const expectedFiles = readdirSync(corpusRoot, { withFileTypes: true })
			.filter(
				(entry) =>
					entry.isFile() &&
					((entry.name.endsWith('.md') && entry.name !== 'README.md') ||
						(entry.name.endsWith('.json') &&
							entry.name !== 'corpus_provenance.json')),
			)
			.map((entry) => entry.name);
		expectedFiles.push('oracle/vbe_oracle_cases.json');

		expect(sidecar.files.map((entry) => entry.path).sort()).toEqual(
			expectedFiles.sort(),
		);

		for (const entry of sidecar.files) {
			expect(allowed.has(entry.provenance)).toBe(true);
			expect(entry.notes.trim().length).toBeGreaterThan(0);
			expect(existsSync(join(corpusRoot, entry.path))).toBe(true);
		}
	});

	it('keeps the audit and oracle evidence plain ASCII', () => {
		// The evidence files are the export contract downstream ports vendor
		// verbatim and checksum, and the house style is plain ASCII prose (no
		// em dashes, no section signs). Guard it so smart punctuation cannot
		// creep back in through an edited notes field.
		for (const relative of ['diagnostic_influence_audit.json', 'oracle/vbe_oracle_cases.json']) {
			const raw = readFileSync(join(corpusRoot, relative), 'utf8');
			// \t\r\n tolerated: \r appears on autocrlf checkouts, never in the blob.
			const offenders = [...new Set(raw.match(/[^\t\r\n -~]/g) ?? [])];
			expect(offenders, relative).toEqual([]);
		}
	});

	it('requires every oracle case to declare case-level provenance', () => {
		const oracle = readJson<OracleCasesFile>(
			join(corpusRoot, 'oracle', 'vbe_oracle_cases.json'),
		);
		const allowed = new Set(allowedProvenance);

		expect(oracle.version).toBe(1);

		for (const fixture of oracle.cases) {
			expect(fixture.id.trim().length).toBeGreaterThan(0);
			expect(allowed.has(fixture.provenance)).toBe(true);
			if (fixture.mode === 'compile_then_run') {
				expect(['compile', 'runtime']).toContain(fixture.evidencePhase);
			} else {
				expect(fixture.evidencePhase).toBe(
					fixture.mode === 'run' ? 'runtime' : 'compile',
				);
			}

			if (fixture.expected === 'observe') {
				expect(fixture.provenance).toBe('observed-not-asserted');
				expect(fixture.diagnosticMeaning).toBe('observation');
			} else {
				expect(fixture.provenance).toBe('vbe-oracle-verified');
				if (fixture.evidencePhase === 'runtime') {
					expect(fixture.diagnosticMeaning).toBe(
						fixture.expected === 'rejected' ? 'runtime-error' : 'runtime-valid',
					);
				} else {
					expect(fixture.diagnosticMeaning).toBe(
						fixture.expected === 'rejected' ? 'compile-error' : 'compile-valid',
					);
				}
			}
		}
	});

	it('keeps the managed backlog connected to every Markdown corpus source', () => {
		const sidecar = readJson<CorpusProvenanceSidecar>(
			join(corpusRoot, 'corpus_provenance.json'),
		);
		const backlog = readFileSync(join(corpusRoot, 'managed_backlog.md'), 'utf8');
		const markdownSources = sidecar.files
			.map((entry) => entry.path)
			.filter((path) => path.endsWith('.md') && path !== 'managed_backlog.md');

		for (const path of markdownSources) {
			expect(backlog).toContain(`\`${path}\``);
		}

		for (const category of [
			'syntax-core',
			'realtime-recovery',
			'diagnostic-ranges',
			'type-analysis',
			'runtime-resolution',
			'project-binder',
			'module-context',
			'object-member',
			'host-behavior',
			'completion-context',
			'userform-designer',
			'limits-boundaries',
			'legacy-edges',
			'tokenizer',
			'roundtrip-io',
			'canary-verdicts',
			'casing',
		]) {
			expect(backlog).toContain(`\`${category}\``);
		}
	});

	it('maps every diagnostic rule to corpus and oracle evidence', () => {
		const audit = readJson<DiagnosticInfluenceAudit>(
			join(corpusRoot, 'diagnostic_influence_audit.json'),
		);
		const oracle = readJson<OracleCasesFile>(
			join(corpusRoot, 'oracle', 'vbe_oracle_cases.json'),
		);
		const allowed = new Set(allowedProvenance);
		const oracleById = new Map(oracle.cases.map((fixture) => [fixture.id, fixture]));
		const auditedCodes = audit.diagnostics.map((entry) => entry.code);
		const ruleByCode = new Map(
			[
				...Object.values(DIAGNOSTIC_RULES),
				// Structural block-balance codes are emitted by analyzeVbaStructure
				// rather than the registry, but are held to the same evidence bar.
				...Object.values(STRUCTURAL_DIAGNOSTIC_RULES),
			].map((rule) => [rule.code, rule]),
		);
		const ruleCodes = [...ruleByCode.keys()];

		expect(audit.version).toBe(1);
		expect(new Set(auditedCodes).size).toBe(auditedCodes.length);
		expect([...auditedCodes].sort()).toEqual([...ruleCodes].sort());

		for (const entry of audit.diagnostics) {
			const rule = ruleByCode.get(entry.code);
			expect(allowed.has(entry.status)).toBe(true);
			expect(entry.diagnosticKind).toBe(rule?.diagnosticKind);
			expect(entry.sourceOfTruth.trim().length).toBeGreaterThan(0);
			expect(Array.isArray(entry.assertedOracleCases)).toBe(true);
			expect(Array.isArray(entry.observeOnlyOracleCases)).toBe(true);
			expect(Array.isArray(entry.pendingLegacyCorpusFiles)).toBe(true);

			if (entry.status === 'vbe-oracle-verified') {
				expect(entry.assertedOracleCases.length).toBeGreaterThan(0);
			}

			for (const caseId of entry.assertedOracleCases) {
				const fixture = oracleById.get(caseId);
				expect(fixture).toBeDefined();
				expect(fixture?.expected).not.toBe('observe');
				expect(fixture?.provenance).toBe('vbe-oracle-verified');
			}

			if (entry.diagnosticKind === 'deterministic-runtime-error') {
				const assertedFixtures = entry.assertedOracleCases
					.map((caseId) => oracleById.get(caseId))
					.filter((fixture): fixture is OracleCase => fixture !== undefined);
				expect(
					assertedFixtures.some(
						(fixture) => fixture.diagnosticMeaning === 'runtime-error',
					),
				).toBe(true);
			}

			for (const caseId of entry.observeOnlyOracleCases) {
				const fixture = oracleById.get(caseId);
				expect(fixture).toBeDefined();
				expect(fixture?.expected).toBe('observe');
				expect(fixture?.provenance).toBe('observed-not-asserted');
			}
		}
	});
});
