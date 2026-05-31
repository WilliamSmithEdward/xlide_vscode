import { existsSync, readdirSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import { DIAGNOSTIC_RULES } from '../src/analyzer';

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
}

interface OracleCasesFile {
	version: number;
	cases: OracleCase[];
}

interface DiagnosticInfluenceEntry {
	code: string;
	status: Provenance;
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

	it('requires every oracle case to declare case-level provenance', () => {
		const oracle = readJson<OracleCasesFile>(
			join(corpusRoot, 'oracle', 'vbe_oracle_cases.json'),
		);
		const allowed = new Set(allowedProvenance);

		expect(oracle.version).toBe(1);

		for (const fixture of oracle.cases) {
			expect(fixture.id.trim().length).toBeGreaterThan(0);
			expect(allowed.has(fixture.provenance)).toBe(true);

			if (fixture.expected === 'observe') {
				expect(fixture.provenance).toBe('observed-not-asserted');
			} else {
				expect(fixture.provenance).toBe('vbe-oracle-verified');
			}
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
		const ruleCodes = Object.values(DIAGNOSTIC_RULES).map((rule) => rule.code);

		expect(audit.version).toBe(1);
		expect(new Set(auditedCodes).size).toBe(auditedCodes.length);
		expect([...auditedCodes].sort()).toEqual([...ruleCodes].sort());

		for (const entry of audit.diagnostics) {
			expect(allowed.has(entry.status)).toBe(true);
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

			for (const caseId of entry.observeOnlyOracleCases) {
				const fixture = oracleById.get(caseId);
				expect(fixture).toBeDefined();
				expect(fixture?.expected).toBe('observe');
				expect(fixture?.provenance).toBe('observed-not-asserted');
			}
		}
	});
});
