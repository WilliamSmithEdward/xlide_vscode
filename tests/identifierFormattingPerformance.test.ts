import { afterEach, describe, expect, it, vi } from 'vitest';
import { writeFileSync } from 'node:fs';
import { resolveIdentifierCompletions } from '../src/analyzer/completion/identifierCompletion';
import { resolveCanonicalCaseEdits } from '../src/analyzer/completion/canonicalCasing';
import * as docs from '../src/analyzer/docs/docModel';
import * as symbols from '../src/analyzer/symbols/symbolModel';
import * as runtimeDocs from '../src/analyzer/runtime/vbaRuntimeDocs';
import type { VbaProcedureSignature } from '../src/analyzer/symbols/symbolModel';

afterEach(() => vi.restoreAllMocks());
function procedure(name: string): VbaProcedureSignature {
    return { name, kind: 'sub', visibility: 'public', moduleName: 'Api', params: [
        { name: 'value', type: 'Long', byVal: true },
    ], doc: { source: 'inline', params: [], summary: name + ' documentation' } };
}
function sourceFor(prefix: string) {
    const source = 'Sub Main()\n    ' + prefix + '\nEnd Sub\n';
    return { source, offset: source.indexOf(prefix) + prefix.length };
}

describe('identifier candidate formatting', () => {
    it('formats only a matching project procedure in a large candidate set', () => {
        const { source, offset } = sourceFor('TargetW');
        const procedures = [...Array.from({ length: 1200 }, (_, index) => procedure('Noise' + index)), procedure('TargetWork')];
        const signature = vi.spyOn(symbols, 'procedureDeclarationSignature');
        const markdown = vi.spyOn(docs, 'renderDocMarkdown');
        const results = resolveIdentifierCompletions(source, offset, {
            includeGlobals: false, includeRuntime: false, projectProcedures: procedures,
        });
        expect(results.map(item => item.name)).toEqual(['TargetWork']);
        expect(results[0].detail).toContain('value As Long');
        expect(results[0].documentation).toContain('TargetWork documentation');
        expect(markdown).toHaveBeenCalledTimes(1);
        expect(signature.mock.calls.filter(([item]) => 'moduleName' in item && item.moduleName === 'Api').length).toBeLessThanOrEqual(2);
    });

    it('does not format a case-insensitive project duplicate shadowed by a local procedure', () => {
        const source = "''' <summary>Local documentation</summary>\nSub TargetWork()\nEnd Sub\nSub Main()\n    targetw\nEnd Sub\n";
        const remote = procedure('targetwork');
        const remoteDoc = vi.fn(() => ({ source: 'inline' as const, params: [], summary: 'Remote documentation' }));
        const remoteParams = vi.fn(() => [{ name: 'other', type: 'String' }]);
        Object.defineProperties(remote, { doc: { get: remoteDoc }, params: { get: remoteParams } });
        const result = resolveIdentifierCompletions(source, source.indexOf('targetw') + 'targetw'.length,
            { includeGlobals: false, includeRuntime: false, projectProcedures: [remote] });
        expect(result.map(item => item.name)).toEqual(['TargetWork']);
        expect(result[0].documentation).toContain('Local documentation');
        expect(remoteDoc).not.toHaveBeenCalled();
        expect(remoteParams).not.toHaveBeenCalled();
    });

    it('does not render runtime or project docs for a prefix with no matches', () => {
        const { source, offset } = sourceFor('ZzNoMatchAnywhere');
        const markdown = vi.spyOn(docs, 'renderDocMarkdown');
        const descriptions = vi.spyOn(runtimeDocs, 'vbaRuntimeDescription');
        expect(resolveIdentifierCompletions(source, offset, { projectProcedures: [procedure('Noise')] })).toEqual([]);
        expect(markdown).not.toHaveBeenCalled();
        expect(descriptions).not.toHaveBeenCalled();
    });

    it('retains docs for matching runtime functions', () => {
        const { source, offset } = sourceFor('Le');
        const descriptions = vi.spyOn(runtimeDocs, 'vbaRuntimeDescription');
        const result = resolveIdentifierCompletions(source, offset);
        expect(result.find(item => item.name === 'Left')?.documentation).toContain('Left');
        expect(result.find(item => item.name === 'Len')?.documentation).toContain('Len');
        expect(descriptions.mock.calls.length).toBeGreaterThan(0);
        expect(descriptions.mock.calls.every(([name]) => name.toLowerCase().startsWith('le'))).toBe(true);
    });

    it('formats the current metadata on each request and returns strings', () => {
        const { source, offset } = sourceFor('Target');
        const candidate = procedure('TargetWork');
        const context = { includeGlobals: false, includeRuntime: false, projectProcedures: [candidate] };
        const first = resolveIdentifierCompletions(source, offset, context)[0];
        candidate.params[0].type = 'String';
        candidate.doc!.summary = 'Updated documentation';
        const second = resolveIdentifierCompletions(source, offset, context)[0];
        expect(first.detail).toContain('Long');
        expect(second.detail).toContain('String');
        expect(second.documentation).toContain('Updated documentation');
        expect(typeof second.detail).toBe('string');
        expect(typeof second.documentation).toBe('string');
    });

    it('does not format irrelevant project docs during automatic casing', () => {
        const source = 'Sub Main()\nDim TargetValue As Long\ntargetvalue = 1\nEnd Sub\n';
        const start = source.indexOf('targetvalue');
        const markdown = vi.spyOn(docs, 'renderDocMarkdown');
        const edits = resolveCanonicalCaseEdits(source, { start, end: start + 'targetvalue = 1'.length },
            { identifier: { includeGlobals: false, includeRuntime: false, projectProcedures: [procedure('Noise')] } });
        expect(edits).toEqual([{ start, end: start + 'targetvalue'.length, text: 'TargetValue' }]);
        expect(markdown).not.toHaveBeenCalled();
    });
});

it.skipIf(!process.env.XLIDE_IDENTIFIER_BENCHMARK_OUTPUT)('measures prefix-filtered identifier and casing work', () => {
    const source = Array.from({ length: 1200 }, (_, index) =>
        "''' <summary>Probe documentation</summary>\nSub Probe" + index + '(ByVal value As Long)\nEnd Sub\n').join('\n') +
        '\nSub Active()\nDim ZzUnique As Long\nzzunique = 1\nEnd Sub\n';
    const offset = source.lastIndexOf('zzunique') + 'zzunique'.length;
    const projectProcedures = Array.from({ length: 1200 }, (_, index) => procedure('External' + index));
    const context = { projectProcedures };
    const medians: Record<string, number> = {};
    const measure = (name: string, invoke: () => unknown) => {
        invoke();
        const times: number[] = [];
        for (let sample = 0; sample < 21; sample++) {
            const start = performance.now(); invoke(); times.push(performance.now() - start);
        }
        medians[name] = times.sort((a, b) => a - b)[10];
    };
    measure('identifierMs', () => resolveIdentifierCompletions(source, offset, context));
    measure('casingLineMs', () => resolveCanonicalCaseEdits(source,
        { start: offset - 'zzunique'.length, end: offset + ' = 1'.length }, { identifier: context }));
    const wideSource = source.replace('zzunique = 1\n', 'zzunique = 1\n    \n');
    const wideOffset = wideSource.indexOf('\n    \n') + 5;
    measure('unfilteredIdentifierMs', () => resolveIdentifierCompletions(wideSource, wideOffset, context));
    writeFileSync(process.env.XLIDE_IDENTIFIER_BENCHMARK_OUTPUT!, JSON.stringify({
        procedures: 1201, projectProcedures: 1200, bytes: source.length, samples: 21, medians,
    }, null, 2));
});
