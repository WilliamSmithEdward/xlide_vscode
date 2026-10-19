import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { retainedExpressionText } from '../../src/analyzer/diagnostics/rules/binaryOperandScalar';
import type { BodyNode } from '../../src/analyzer/parser/nodes';
import { analyzeModule } from '../../src/analyzer';
import { analyzeVbaModuleSource } from '../../src/vbaModuleAnalysis';
import { buildVbaProjectIndex, effectiveModuleKind, projectAnalysisOptionsForModule, projectProcedureSignatures } from '../../src/vbaProjectAnalysis';
import { oracleModuleInput } from '../helpers/oracleModuleInput';

const codes = new Set(['is-operator-non-object', 'typeof-invalid-expression', 'non-scalar-binary-operand']);
const batches = ['review', 'validation', 'completion'].map(name => JSON.parse(readFileSync(`syntax_corpus/oracle/evidence/gap_contexts_${name}_20261019.json`, 'utf8')));
const observations = batches.flatMap(batch => {
    const results = new Map<string, any>(batch.checkpointRecords.map((record: any) => [record.result.caseId, record.result]));
    return batch.cases.filter((c: any) => /^(contexts_review_builtin_.*_(is|typeof)|r3_builtin_.*_(is|typeof)|imported_enum_|typeof_literal_|file_slot_|intrinsic_argument_|r2_intrinsic_|r2_(lock|unlock)_file_)/.test(c.id))
        .map((c: any) => ({ c, result: results.get(c.id) }));
});

describe('measured TypeOf grammar and scalar operand selections', () => {
    it.each(observations)('$c.id', ({ c, result }) => {
        const modules = c.modules.map(oracleModuleInput);
        const project = buildVbaProjectIndex(modules, undefined, { conditionalCompilation: { projectConstants: {} } });
        const signatures = projectProcedureSignatures(project);
        const diagnostics = modules.flatMap((mod: any) => analyzeVbaModuleSource({
            source: mod.source, moduleName: mod.moduleName, moduleKind: effectiveModuleKind(mod),
            ...projectAnalysisOptionsForModule(project, mod.moduleName, signatures),
        }).diagnostics.filter(d => codes.has(d.code ?? '')).map(d => ({ ...d, source: mod.source })));
        if (result.outcome === 'accepted') { expect(diagnostics).toEqual([]); return; }
        expect(result.outcome).toBe('rejected');
        const message = result.message.replace(/^Compile error:\s*/, '').trim();
        if (message === 'User-defined type not defined') { expect(diagnostics).toEqual([]); return; }
        expect(diagnostics.some((d: any) => d.message.includes(message)), JSON.stringify(diagnostics)).toBe(true);
        // These are supplementary post-dismissal selections, not live highlight proof.
        const selection = result.selection;
        if (selection?.startLine === selection?.endLine && selection.startColumn > 1) {
            const selected = selection.source.split(/\r?\n/)[selection.startLine - 1].slice(selection.startColumn - 1, selection.endColumn - 1);
            expect(diagnostics.map((d: any) => d.source.slice(d.span.start, d.span.end))).toContain(selected);
        }
    });

    const ownDiagnostics = (source: string) => analyzeModule(source).filter(d => codes.has(d.code ?? ''));
    it.each(['CVar', 'Array', 'CVErr', 'Abs', 'Int', 'Fix', 'InStrB'])('preserves source callable shadow %s', name => {
        const source = `Function ${name}(x As Variant) As Object\nEnd Function\nSub Main()\nDebug.Print TypeOf ${name}(1) Is Collection\nDebug.Print ${name}(1) Is Nothing\nEnd Sub`;
        expect(ownDiagnostics(source)).toEqual([]);
    });
    it.each(['CVar', 'Int', 'InStrB'])('retains scalar source-function diagnostics for %s', name => {
        const source = `Function ${name}(x As Variant) As Long\nEnd Function\nSub Main()\nDebug.Print TypeOf ${name}(1) Is Collection\nDebug.Print ${name}(1) Is Nothing\nEnd Sub`;
        const found = ownDiagnostics(source);
        expect(found).toHaveLength(2);
        expect(found.every(d => d.message.includes('Type mismatch'))).toBe(true);
        expect(found.map(d => source.slice(d.span.start, d.span.end))).toEqual([name, name]);
    });
    it.each(['Object', 'Variant'])('preserves local %s and late-bound members', type => {
        const source = `Sub Main()\nDim CVar As ${type}\nDim obj As Object\nDebug.Print TypeOf CVar Is Collection\nDebug.Print TypeOf obj.CVar(1) Is Collection\nDebug.Print obj.Int(1) Is Nothing\nEnd Sub`;
        expect(ownDiagnostics(source)).toEqual([]);
    });
    it('keeps qualification out of the unqualified intrinsic syntax table', () => {
        const source = 'Sub Main()\nDebug.Print TypeOf VBA.CVar(1) Is Collection\nEnd Sub';
        expect(ownDiagnostics(source)).toEqual([]);
    });
    it('locates file-slot addition across whitespace and continuations', () => {
        const source = 'Sub Main()\nDim a(1) As Long\nPrint #(a _\n  +   1), 1\nEnd Sub';
        const found = ownDiagnostics(source);
        expect(found).toHaveLength(1);
        expect(source.slice(found[0].span.start, found[0].span.end)).toBe('+');
    });
    it('uses logarithmic containing-statement lookup without reading unrelated text', () => {
        let reads = 0;
        const body = Array.from({ length: 1024 }, (_, i) => ({
            kind: 'Statement',
            get span() { reads++; return { start: i * 10, end: i * 10 + 9 }; },
            get raw() { expect(i).toBe(900); return 'a + 1    '; },
        })) as BodyNode[];
        for (let pass = 0; pass < 2; pass++) {
            reads = 0;
            expect(retainedExpressionText(body, { start: 9002, end: 9003 })).toBe('+');
            expect(reads).toBeLessThan(20);
        }
    });
    it('preserves concatenation and Nothing operand selections', () => {
        for (const [expression, selected] of [['"x" & a', 'a'], ['Nothing + 1', 'Nothing']]) {
            const source = `Sub Main()\nDim a(1) As Long\nDebug.Print ${expression}\nEnd Sub`;
            const found = ownDiagnostics(source);
            expect(found).toHaveLength(1);
            expect(source.slice(found[0].span.start, found[0].span.end)).toBe(selected);
        }
    });
});
