import { afterEach, describe, expect, it, vi } from 'vitest';
import { ProjectIndex } from '../src/analyzer/symbols/projectIndex';
import * as valueFacts from '../src/analyzer/symbols/classMemberFacts';
import { projectEditorSymbolContextForModule, projectAnalysisOptionsForModule } from '../src/vbaProjectAnalysis';

afterEach(() => { vi.restoreAllMocks(); });

function project() {
    const index = new ProjectIndex();
    index.setModule({ moduleName: 'Probe', moduleKind: 'class', source: [
        'Public Untouched As Object',
        'Public EmptyValue As Variant',
        'Public Function NothingValue() As Object',
        ' Set NothingValue = Nothing',
        'End Function',
        'Public Property Get ScalarValue() As Variant',
        ' ScalarValue = 42',
        'End Property',
    ].join('\r\n') });
    index.setModule({ moduleName: 'Caller', moduleKind: 'standard', source: 'Sub Main()\r\nEnd Sub' });
    return index;
}

describe('editor class surfaces skip diagnostic value facts', () => {
    it.each([false, true])('keeps diagnostics complete and caches separate, diagnostics first: %s', diagnosticsFirst => {
        const index = project();
        const scan = vi.spyOn(valueFacts, 'classMemberValues');
        if (diagnosticsFirst) { projectAnalysisOptionsForModule(index, 'Caller'); }
        const before = scan.mock.calls.length;
        const editor = projectEditorSymbolContextForModule(index, 'Caller');
        expect(scan.mock.calls.length).toBe(before);
        const editorMembers = editor.analysisOptions.projectClassMembers!.find(type => type.name === 'Probe')!.members;
        expect(editorMembers.every(member => !('knownValue' in member))).toBe(true);
        const snapshot = structuredClone(editorMembers);
        const analysis = projectAnalysisOptionsForModule(index, 'Caller');
        expect(scan).toHaveBeenCalledTimes(1);
        const diagnosticMembers = analysis.projectClassMembers!.find(type => type.name === 'Probe')!.members;
        expect(Object.fromEntries(diagnosticMembers.map(member => [member.name, member.knownValue]))).toEqual({
            Untouched: 'nothing', EmptyValue: 'empty', NothingValue: 'nothing', ScalarValue: 'scalar',
        });
        expect(editorMembers).toEqual(snapshot);
        expect(diagnosticMembers.map(({ knownValue, ...member }) => member)).toEqual(editorMembers);
        expect(projectEditorSymbolContextForModule(index, 'Caller').analysisOptions.projectClassMembers).toEqual(editor.analysisOptions.projectClassMembers);
        expect(index.projectClassMembers()[0].members).toEqual(diagnosticMembers);
    });

    it('refreshes both surfaces after source edits', () => {
        const index = project();
        projectEditorSymbolContextForModule(index, 'Caller');
        projectAnalysisOptionsForModule(index, 'Caller');
        index.setModule({ moduleName: 'Probe', moduleKind: 'class', source: 'Public Changed As Object\r\nPrivate Sub AssignIt()\r\n Set Changed = New Collection\r\nEnd Sub' });
        const editor = projectEditorSymbolContextForModule(index, 'Caller').analysisOptions.projectClassMembers!.find(type => type.name === 'Probe')!;
        const diagnostic = projectAnalysisOptionsForModule(index, 'Caller').projectClassMembers!.find(type => type.name === 'Probe')!;
        expect(editor.members.map(member => member.name)).toEqual(['Changed']);
        expect(diagnostic.members[0].knownValue).toBeUndefined();
        expect(diagnostic.members).toEqual(editor.members);
    });
});

// Standard-module getters share the lazy diagnostic projection, without making
// completion scan procedure bodies or invalidating unchanged library facts.
it.each([false, true])('keeps standard getter scans lazy, diagnostics first: %s', diagnosticsFirst => {
    const index = new ProjectIndex();
    index.setModule({moduleName: 'Library', moduleKind: 'standard', source: 'Public Mutable As Variant\nPublic Property Get Value(ByVal index As Long) As Variant\nValue = index\nEnd Property'});
    index.setModule({moduleName: 'Caller', moduleKind: 'standard', source: 'Sub T()\nEnd Sub'});
    const scan = vi.spyOn(valueFacts, 'classMemberValues');
    if (diagnosticsFirst) { projectAnalysisOptionsForModule(index, 'Caller'); }
    const before = scan.mock.calls.length;
    const editor = projectEditorSymbolContextForModule(index, 'Caller');
    expect(scan.mock.calls.length).toBe(before);
    const light = editor.analysisOptions.projectClassMembers!.find(type => type.name === 'Library')!.members;
    expect(light.every(member => !('knownValue' in member))).toBe(true);
    const snapshot = structuredClone(light);
    const analysis = projectAnalysisOptionsForModule(index, 'Caller');
    const members = analysis.projectClassMembers!.find(type => type.name === 'Library')!.members;
    expect(members.find(member => member.name === 'Value')!.knownValue).toBe('scalar');
    expect(members.find(member => member.name === 'Mutable')!.knownValue).toBeUndefined();
    expect(scan).toHaveBeenCalledTimes(1);
    expect(light).toEqual(snapshot);
    index.setModule({moduleName: 'Caller', moduleKind: 'standard', source: 'Sub T()\nDim changed As Long\nEnd Sub'});
    projectAnalysisOptionsForModule(index, 'Caller');
    expect(scan).toHaveBeenCalledTimes(1);
    index.setModule({moduleName: 'Library', moduleKind: 'standard', source: 'Public Property Get Value() As Variant\nSet Value = New Collection\nEnd Property'});
    expect(projectAnalysisOptionsForModule(index, 'Caller').projectClassMembers!.find(type => type.name === 'Library')!.members[0].knownValue).toBeUndefined();
    expect(scan).toHaveBeenCalledTimes(2);
});
