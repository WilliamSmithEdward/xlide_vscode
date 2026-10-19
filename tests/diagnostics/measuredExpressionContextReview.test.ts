import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { diagnosticMetadataForCode } from '../../src/analyzer';
import { analyzeVbaModuleSource } from '../../src/vbaModuleAnalysis';
import { buildVbaProjectIndex, effectiveModuleKind, projectAnalysisOptionsForModule, projectProcedureSignatures } from '../../src/vbaProjectAnalysis';
import { oracleModuleInput } from '../helpers/oracleModuleInput';

const documents = ['review', 'validation', 'completion', 'qualified', 'comparison_neighbors', 'receiver_shadows', 'ci_neighbors'].map(name => JSON.parse(readFileSync(`syntax_corpus/oracle/evidence/gap_contexts_${name}_20261019.json`, 'utf8')));
const evidence = { cases: documents.flatMap(d => d.cases), checkpointRecords: documents.flatMap(d => d.checkpointRecords) };
const results = new Map<string, any>(evidence.checkpointRecords.map((record: any) => [record.result.caseId, record.result]));
const changedRangeCodes = new Set(['is-operator-non-object', 'typeof-invalid-expression', 'non-scalar-binary-operand', 'argument-count', 'member-not-found']);

describe('measured Excel expression context parity', () => {
    it.each(evidence.cases.map((c: any) => [c.id, c]))('%s', (_id: string, c: any) => {
        const result = results.get(c.id);
        expect(['accepted', 'rejected']).toContain(result.outcome);
        const modules = c.modules.map(oracleModuleInput);
        const project = buildVbaProjectIndex(modules, undefined, { conditionalCompilation: { projectConstants: {} } });
        const signatures = projectProcedureSignatures(project);
        const diagnostics = modules.flatMap((mod: any) => analyzeVbaModuleSource({
            source: mod.source, moduleName: mod.moduleName, moduleKind: effectiveModuleKind(mod),
            ...projectAnalysisOptionsForModule(project, mod.moduleName, signatures),
        }).diagnostics.filter(d => diagnosticMetadataForCode(d.code ?? '')?.diagnosticKind === 'compile-error')
            .map(d => ({ ...d, moduleName: mod.moduleName, source: mod.source })));
        if (result.outcome === 'accepted') { expect(diagnostics).toEqual([]); return; }
        const message = result.message.replace(/^Compile error:\s*/, '').trim();
        expect(diagnostics.some((d: any) => d.message.toLowerCase().includes(message.toLowerCase())), JSON.stringify(diagnostics)).toBe(true);
        expect(diagnostics.filter((d: any) => changedRangeCodes.has(d.code)).every((d: any) => d.message.toLowerCase().includes(message.toLowerCase())), JSON.stringify(diagnostics)).toBe(true);
        // Compare retained post-dismissal selections without claiming live-dialog capture.
        const selection = result.selection;
        if (selection?.startLine === selection?.endLine && diagnostics.some((d: any) => changedRangeCodes.has(d.code))) {
            const selected = selection.source.split(/\r?\n/)[selection.startLine - 1].slice(selection.startColumn - 1, selection.endColumn - 1);
            expect(diagnostics.some((d: any) => d.moduleName === selection.moduleName && d.source.slice(d.span.start, d.span.end) === selected), selected + JSON.stringify(diagnostics)).toBe(true);
        }
    });
});
