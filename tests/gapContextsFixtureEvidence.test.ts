import { createHash } from 'node:crypto';
import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';

const root = 'syntax_corpus/oracle/';
const read = (name: string) => JSON.parse(readFileSync(root + name, 'utf8'));
const historical = read('evidence/gap_contexts_historical_20261016.json');
const corpus = read('vbe_oracle_cases.json');

describe('gap-contexts fresh measurement provenance', () => {
    it.each(['review', 'validation'])('binds every %s result to exact harness, build, input and canonical outcome', (name) => {
        const document = read(`evidence/gap_contexts_${name}_20261019.json`);
        const snapshots = read('evidence/gap_contexts_harness_20261019.json');
        const snapshot = snapshots.versions[document.harnessSnapshot.version];
        const hash = createHash('sha256').update(snapshots.coordinatorText).update(snapshot.workerText).digest('hex');
        expect(hash).toBe(document.harnessSnapshot.hash);
        expect(document.oracleFailureCount).toBe(0);
        expect(document.skippedCount).toBe(0);
        expect(document.checkpointRecords).toHaveLength(document.cases.length);
        const canonical = new Map(corpus.cases.map((c: any) => [c.id, c]));
        for (const c of document.cases) {
            const record = document.checkpointRecords.find((r: any) => r.result.caseId === c.id);
            expect(record, c.id).toBeDefined();
            expect(record.result.hostBuild).toBe(document.hostBuild);
            expect(['accepted', 'rejected']).toContain(record.result.outcome);
            expect(createHash('sha256').update(hash).update(JSON.stringify([document.hostBuild, c])).digest('hex'), c.id).toBe(record.caseHash);
            const promoted: any = canonical.get(c.id);
            expect(promoted.expected, c.id).toBe(record.result.outcome);
            expect(promoted.modules, c.id).toEqual(c.modules);
            expect(promoted.provenance).toBe('vbe-oracle-verified');
            expect(promoted.evidencePhase).toBe('compile');
        }
    });
});

describe('gap-contexts historical evidence', () => {
    it('retains the original input and result bytes with matching hashes', () => {
        for (const artifact of Object.values(historical.artifacts) as Record<string, string>[]) {
            for (const key of ['input', 'rawLog']) {
                expect(createHash('sha256').update(artifact[key + 'Text']).digest('hex')).toBe(artifact[key + 'Sha256']);
            }
        }
    });
    it('asserts only the measured compile phase and preserves every module source', () => {
        expect(historical.measurements).toHaveLength(44);
        expect(new Set(historical.measurements.map((m: any) => m.fixtureId)).size).toBe(44);
        for (const m of historical.measurements) {
            const c = corpus.cases.find((c: any) => c.id === m.fixtureId);
            expect(c, m.fixtureId).toBeDefined();
            expect(c.expected, m.fixtureId).toBe(m.result.compile);
            expect(c.provenance).toBe('vbe-oracle-verified');
            expect(c.mode).toBe('compile');
            expect(c.evidencePhase).toBe('compile');
            expect(c.diagnosticMeaning).toBe(m.result.compile === 'accepted' ? 'compile-valid' : 'compile-error');
            expect(m.selection).toBeNull();
            const modules = m.input.modules ?? [{ name: 'Module1', kind: 'standard', source: m.input.source }];
            expect(c.modules.map((module: any) => ({ name: module.name, type: module.type, source: module.source }))).toEqual(
                modules.map((module: any) => ({ name: module.name, type: module.kind, source: module.source })),
            );
        }
    });
    it('keeps unmeasured review and designer proposals as compile observations', () => {
        const proposals = [
            ...read('probes/gap_contexts_review_20261019.json').cases,
            ...read('probes/gap_contexts_designer_20261019.json').cases,
        ];
        expect(proposals).toHaveLength(415);
        expect(new Set(proposals.map((c: any) => c.id)).size).toBe(415);
        for (const c of proposals) {
            expect(c.expected, c.id).toBe('observe');
            expect(c.provenance).toBe('observed-not-asserted');
            expect(c.mode).toBe('compile');
            expect(c.evidencePhase).toBe('compile');
            expect(c.diagnosticMeaning).toBe('observation');
            expect(c.modules.every((m: any) => m.name && m.type && typeof m.source === 'string'), c.id).toBe(true);
        }
    });
});
