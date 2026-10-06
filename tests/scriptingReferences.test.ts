import { describe, expect, it } from 'vitest';
import { analyzeModule } from '../src/analyzer/diagnostics/analyzeModule';
import { hostTokenForLibid, hostTokensForProject, referencedHostTokens } from '../src/analyzer/host/hostLibraries';
import { hostObjectModelForTokens } from '../src/analyzer/host/hostRegistry';
import { resolveMemberCompletions } from '../src/analyzer/completion/memberAccess';
import { AnalysisWorkerState } from '../src/analysisWorkerLogic';
import { resolveHover } from '../src/analyzer/hover/resolveHover';

const refs = [
    { name: 'Scripting', libid: '*\\G{420B2830-E718-11CF-893D-00A0C9054228}#1.0#0#scrrun.dll#Microsoft Scripting Runtime' },
    { name: 'VBScript_RegExp_55', libid: '*\\G{3F4DACA7-160D-11D2-A8E9-00104B365C9F}#5.5#0#vbscript.dll\\3#Microsoft VBScript Regular Expressions 5.5' },
];
const wrap = (body: string) => `Option Explicit\nSub Main()\n${body}\nEnd Sub`;
const model = () => hostObjectModelForTokens(hostTokensForProject('excel', refs));
const findings = (body: string, references = refs) => analyzeModule(wrap(body), {
    moduleName: 'Module1', moduleKind: 'standard', host: 'excel', knownIdentifiers: new Set<string>(),
    referencedHosts: referencedHostTokens('excel', references),
    referencedLibraries: ['VBA', 'Excel', 'Office', 'stdole', ...references.map(r => r.name)],
});
const members = (body: string, marker: string, references = refs) => {
    const source = wrap(body);
    return resolveMemberCompletions(source, source.indexOf(marker) + marker.length, {
        model: hostObjectModelForTokens(hostTokensForProject('excel', references)),
    }).map(m => m.name);
};

describe('referenced Scripting and RegExp libraries (VBIDE issue #56)', () => {
    it('recognizes GUID identity independently of reference names or DLL paths', () => {
        expect(refs.map(r => hostTokenForLibid(r.libid.toLowerCase()))).toEqual(['scripting', 'regexp']);
        expect(hostTokensForProject('excel', [refs[1], refs[0], refs[1]])).toEqual(['excel', 'regexp', 'scripting']);
        expect(model()?.hostName).toBe('Excel');
    });
    for (const constant of ['ForAppending', 'ForReading', 'ForWriting', 'TristateTrue', 'TristateFalse', 'TristateUseDefault']) {
        it(`resolves ${constant} only with the reference`, () => {
            expect(findings(`Debug.Print ${constant}`).map(d => d.code)).toEqual([]);
            expect(findings(`Debug.Print ${constant}`, []).some(d => d.code === 'undeclared-variable')).toBe(true);
        });
    }
    for (const expression of ['Scripting.ForAppending', 'IOMode.ForAppending', 'Scripting.IOMode.ForAppending']) {
        it(`resolves ${expression}`, () => {
            expect(findings(`Debug.Print ${expression}`).map(d => `${d.code}: ${d.message}`)).toEqual([]);
        });
    }
    it('resolves the reported GetFile/OpenAsTextStream call and its receiver chain', () => {
        const declaration = 'Dim fso As Scripting.FileSystemObject\nSet fso = New Scripting.FileSystemObject';
        expect(findings(`${declaration}\nDim stream As Scripting.TextStream\nIf fso.FileExists("x") Then\nSet stream = fso.GetFile("x").OpenAsTextStream(ForAppending)\nstream.WriteLine "text"\nstream.Close\nEnd If`).map(d => `${d.code}: ${d.message}`)).toEqual([]);
        expect(members(`${declaration}\nfso.GetFile("x").OpenAsTextStream(ForAppending).`, 'OpenAsTextStream(ForAppending).')).toContain('WriteLine');
    });
    it('limits FileSystemObject completion to its own members', () => {
        const names = members('Dim y As Scripting.FileSystemObject\nSet y = New Scripting.FileSystemObject\ny.', 'y.');
        expect(names).toEqual(expect.arrayContaining(['FileExists', 'GetFile', 'OpenTextFile']));
        expect(names).not.toContain('Pattern');
        expect(names).not.toContain('Workbook');
        expect(members('Dim y As Scripting.FileSystemObject\nSet y = New Scripting.FileSystemObject\ny.', 'y.', [])).not.toContain('FileExists');
    });
    it('limits RegExp completion to its own members', () => {
        const names = members('Dim x As RegExp\nSet x = New RegExp\nx.', 'x.');
        expect(names).toEqual(expect.arrayContaining(['Pattern', 'Test', 'Execute', 'Multiline']));
        expect(names).not.toContain('FileExists');
        expect(names).not.toContain('Workbook');
        expect(members('Dim x As RegExp\nSet x = New RegExp\nx.', 'x.', [])).not.toContain('Pattern');
        expect(findings('Dim x As New RegExp\nx.Pattern = "A"').map(d => d.code)).toEqual([]);
    });
    it('preserves genuine undeclared names and local shadowing', () => {
        expect(findings('Debug.Print ForAppeding').some(d => d.code === 'undeclared-variable')).toBe(true);
        expect(findings('Dim ForAppending As String\nForAppending = "local"\nDebug.Print ForAppending').map(d => d.code)).toEqual([]);
    });
    for (const host of ['word', 'powerpoint', 'access', 'vb6']) {
        it(`adds reference constants without changing the ${host} host`, () => {
            const source = wrap('Debug.Print ForAppending');
            const result = analyzeModule(source, {
                host, referencedHosts: ['scripting'], knownIdentifiers: new Set<string>(),
            });
            expect(result.map(d => d.code)).toEqual([]);
            expect(hostObjectModelForTokens([host, 'scripting'])?.hostName?.toLowerCase())
                .toBe(host);
        });
    }
    it('refreshes worker diagnostics when the reference is removed', () => {
        const state = new AnalysisWorkerState();
        const source = wrap('Debug.Print ForAppending');
        state.handle({ kind: 'seed', projectKey: 'refs', generation: 1,
            modules: [{ moduleName: 'Module1', source, type: 'standard' }] });
        const request = { kind: 'analyze' as const, requestId: 1, docKey: 'doc', projectKey: 'refs',
            generation: 1, source, moduleName: 'Module1', moduleType: 'standard', host: 'excel' };
        const present = state.handle({ ...request, referencedHosts: ['scripting'] });
        expect(present?.kind).toBe('result');
        if (present?.kind !== 'result') { throw new Error('Expected worker result'); }
        expect(present.diagnostics.some(d => /ForAppending/.test(d.message))).toBe(false);
        const absent = state.handle({ ...request, requestId: 2, referencedHosts: [] });
        expect(absent?.kind).toBe('result');
        if (absent?.kind !== 'result') { throw new Error('Expected worker result'); }
        expect(absent.diagnostics.some(d => d.code === 'undeclared-variable' && /ForAppending/.test(d.message))).toBe(true);
    });
    it('offers verified constant values to hover', () => {
        const source = wrap('Debug.Print ForAppending');
        const hover = resolveHover(source, source.indexOf('ForAppending') + 3, { model: model() });
        expect(hover?.signature).toMatch(/ForAppending/);
        expect(JSON.stringify(hover)).toMatch(/8/);
    });
});