import { describe, expect, it, vi } from 'vitest';
import * as path from 'path';

vi.mock('vscode', async () => (await import('./helpers/vscodeMock')).vscodeMock());

import { buildVbaProjectIndex, type VbaProjectModuleInput } from '../src/vbaProjectAnalysis';
import type { VbaNavigationModule } from '../src/vbaNavigation';
import { projectStandardModuleReferenceLocations } from '../src/vbaStandardModuleRename';

// Renaming a standard module rewrites `Module.Member` qualifiers across the
// project. The qualifier scan must only match the MODULE's name in qualifier
// position: the same identifier as a member of another receiver (`rs.Fields`),
// a With-block member (`.Fields`), a bang access (`rs!Fields`), or a name
// shadowed by a local or module-level variable refers to something else, and
// rewriting it would corrupt code the rename never owned.

function locations(modules: VbaProjectModuleInput[], oldName: string): string[] {
    const projectPath = path.join(path.sep, 'work', 'book.xlsm');
    const project = buildVbaProjectIndex(modules);
    const byModule = new Map<string, VbaNavigationModule>(
        modules.map((mod) => [mod.moduleName.toLowerCase(), mod]),
    );
    return projectStandardModuleReferenceLocations(projectPath, byModule, project, oldName)
        .map((ref) => `${ref.uri.path.split('/').pop()}:${ref.range.start.line}`);
}

const FIELDS_MODULE: VbaProjectModuleInput = {
    moduleName: 'Fields',
    type: 'standard',
    source: [
        'Public Count As Long',
        'Public Sub Item(ByVal index As Long)',
        'End Sub',
    ].join('\n'),
};

describe('standard module rename qualifier scan', () => {
    it('rewrites a genuine module qualifier and nothing in strings or comments', () => {
        const caller = [
            'Public Sub UseModule()',
            '    Fields.Item 1',
            '    Debug.Print "Fields.Item"',
            "    ' Fields.Item",
            'End Sub',
        ].join('\n');

        expect(locations([
            FIELDS_MODULE,
            { moduleName: 'Caller', type: 'standard', source: caller },
        ], 'Fields')).toEqual(['Caller.bas:1']);
    });

    it('leaves the same name alone as a member of another receiver', () => {
        const caller = [
            'Public Sub MemberOfOther(ByVal rs As Object)',
            '    rs.Fields.Item 1',
            '    rs!Fields.Item 2',
            '    With rs',
            '        .Fields.Item 3',
            '    End With',
            '    Fields.Item 4',
            'End Sub',
        ].join('\n');

        expect(locations([
            FIELDS_MODULE,
            { moduleName: 'Caller', type: 'standard', source: caller },
        ], 'Fields')).toEqual(['Caller.bas:6']);
    });

    it('leaves occurrences shadowed by a local variable alone', () => {
        const caller = [
            'Public Sub Shadowed()',
            '    Dim Fields As Object',
            '    Fields.Item 1',
            'End Sub',
            'Public Sub NotShadowed()',
            '    Fields.Item 2',
            'End Sub',
        ].join('\n');

        expect(locations([
            FIELDS_MODULE,
            { moduleName: 'Caller', type: 'standard', source: caller },
        ], 'Fields')).toEqual(['Caller.bas:5']);
    });

    it('leaves occurrences shadowed by a parameter alone', () => {
        const caller = [
            'Public Sub ParamShadow(ByVal Fields As Object)',
            '    Fields.Item 1',
            'End Sub',
        ].join('\n');

        expect(locations([
            FIELDS_MODULE,
            { moduleName: 'Caller', type: 'standard', source: caller },
        ], 'Fields')).toEqual([]);
    });

    it('leaves a whole module alone when a module-level variable shadows the name', () => {
        const shadowed = [
            'Private Fields As Object',
            'Public Sub UseShadow()',
            '    Fields.Item 1',
            'End Sub',
        ].join('\n');
        const clean = [
            'Public Sub UseModule()',
            '    Fields.Item 1',
            'End Sub',
        ].join('\n');

        expect(locations([
            FIELDS_MODULE,
            { moduleName: 'ShadowMod', type: 'standard', source: shadowed },
            { moduleName: 'CleanMod', type: 'standard', source: clean },
        ], 'Fields')).toEqual(['CleanMod.bas:1']);
    });
});
