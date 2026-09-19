import { describe, expect, it, vi } from 'vitest';
import * as path from 'path';

vi.mock('vscode', async () => (await import('./helpers/vscodeMock')).vscodeMock());

import { buildVbaProjectIndex, type VbaProjectModuleInput } from '../src/vbaProjectAnalysis';
import type { VbaNavigationModule } from '../src/vbaNavigation';
import { projectUserFormReferenceLocations } from '../src/vbaUserFormRename';

// Renaming a UserForm from the tree rewrites every use of its name: as a type
// (`As`, `New`, `TypeOf ... Is`) and as its default instance, which is a bare
// name anywhere a value stands (`PickForm.Show`, `Unload PickForm`,
// `With PickForm`, `Set PickForm = Nothing`). A member of another receiver, a
// named argument, a declaration's own name, a label and a shadowed name are
// something else, and stay as they are.

const projectPath = path.join(path.sep, 'work', 'book.xlsm');

const FORM: VbaProjectModuleInput = {
    moduleName: 'PickForm',
    type: 'userform',
    source: [
        'Private Sub OkButton_Click()',
        '    PickForm.Tag = "done"',
        '    Me.Hide',
        'End Sub',
    ].join('\n'),
};

function renamed(modules: VbaProjectModuleInput[], oldName: string, newName: string): Record<string, string> {
    const project = buildVbaProjectIndex(modules);
    const byModule = new Map<string, VbaNavigationModule>(
        modules.map((mod) => [mod.moduleName.toLowerCase(), mod]),
    );
    const out: Record<string, string> = {};
    for (const mod of modules) {
        out[mod.moduleName] = mod.source;
    }
    const locations = projectUserFormReferenceLocations(projectPath, byModule, project, oldName, newName);
    // Apply from the end of each module so earlier ranges stay valid.
    for (const location of [...locations].reverse()) {
        const file = location.uri.path.split('/').pop()!.replace(/\.bas$/, '');
        const owner = file === newName ? oldName : file;
        const lines = out[owner].split('\n');
        const { start, end } = location.range;
        const line = lines[start.line];
        lines[start.line] = line.slice(0, start.character) + newName + line.slice(end.character);
        out[owner] = lines.join('\n');
    }
    return out;
}

describe('renaming a UserForm rewrites its uses', () => {
    it('rewrites the form as a type and as its default instance, in any module and in its own code', () => {
        const caller = [
            'Public Sub ShowIt()',
            '    Dim f As PickForm',
            '    Set f = New PickForm',
            '    If TypeOf f Is PickForm Then Debug.Print 1',
            '    Load PickForm',
            '    PickForm.Show',
            '    With PickForm',
            '        .Caption = "PickForm"',
            '    End With',
            '    Call Report(PickForm)',
            '    Unload PickForm',
            '    Set PickForm = Nothing',
            "    ' PickForm stays in a comment",
            'End Sub',
        ].join('\n');

        const out = renamed([FORM, { moduleName: 'Caller', type: 'standard', source: caller }], 'PickForm', 'ChooserForm');

        expect(out.Caller.split('\n')).toEqual([
            'Public Sub ShowIt()',
            '    Dim f As ChooserForm',
            '    Set f = New ChooserForm',
            '    If TypeOf f Is ChooserForm Then Debug.Print 1',
            '    Load ChooserForm',
            '    ChooserForm.Show',
            '    With ChooserForm',
            '        .Caption = "PickForm"',
            '    End With',
            '    Call Report(ChooserForm)',
            '    Unload ChooserForm',
            '    Set ChooserForm = Nothing',
            "    ' PickForm stays in a comment",
            'End Sub',
        ]);
        expect(out.PickForm.split('\n')[1]).toBe('    ChooserForm.Tag = "done"');
    });

    it('addresses the edits to the form\'s own code at its new name', () => {
        const project = buildVbaProjectIndex([FORM]);
        const byModule = new Map<string, VbaNavigationModule>([['pickform', FORM]]);
        const files = projectUserFormReferenceLocations(projectPath, byModule, project, 'PickForm', 'ChooserForm')
            .map((location) => location.uri.path.split('/').pop());
        expect(files).toEqual(['ChooserForm.bas']);
    });

    it('leaves members of other receivers, named arguments, declarations and labels alone', () => {
        const caller = [
            'Public Sub Others(ByVal wb As Object)',
            '    wb.PickForm.Show',
            '    wb!PickForm = 1',
            '    Report PickForm:=1',
            'PickForm:',
            '    PickForm.Show',
            'End Sub',
            'Public Property Get PickFormCount() As Long',
            'End Property',
        ].join('\n');

        const out = renamed([FORM, { moduleName: 'Caller', type: 'standard', source: caller }], 'PickForm', 'ChooserForm');

        expect(out.Caller.split('\n')).toEqual([
            'Public Sub Others(ByVal wb As Object)',
            '    wb.PickForm.Show',
            '    wb!PickForm = 1',
            '    Report PickForm:=1',
            'PickForm:',
            '    ChooserForm.Show',
            'End Sub',
            'Public Property Get PickFormCount() As Long',
            'End Property',
        ]);
    });

    it('leaves a use a local or module-level declaration shadows alone', () => {
        const local = [
            'Public Sub Shadowed()',
            '    Dim PickForm As Object',
            '    PickForm.Show',
            'End Sub',
            'Public Sub NotShadowed()',
            '    PickForm.Show',
            'End Sub',
        ].join('\n');
        const moduleLevel = [
            'Private PickForm As Object',
            'Public Sub UsesTheVariable()',
            '    PickForm.Show',
            'End Sub',
        ].join('\n');

        const out = renamed([
            FORM,
            { moduleName: 'Local', type: 'standard', source: local },
            { moduleName: 'ModuleLevel', type: 'standard', source: moduleLevel },
        ], 'PickForm', 'ChooserForm');

        expect(out.Local.split('\n')[2]).toBe('    PickForm.Show');
        expect(out.Local.split('\n')[5]).toBe('    ChooserForm.Show');
        expect(out.ModuleLevel).toBe(moduleLevel);
    });

    it('keeps the brackets on a bracketed use', () => {
        const caller = 'Public Sub Go()\n    [PickForm].Show\nEnd Sub';
        const out = renamed([FORM, { moduleName: 'Caller', type: 'standard', source: caller }], 'PickForm', 'ChooserForm');
        expect(out.Caller.split('\n')[1]).toBe('    [ChooserForm].Show');
    });

    it('finds nothing for a module that is not a form', () => {
        const project = buildVbaProjectIndex([{ moduleName: 'Helpers', type: 'standard', source: 'Public Sub A()\nEnd Sub' }]);
        expect(projectUserFormReferenceLocations(projectPath, new Map(), project, 'Helpers')).toEqual([]);
    });
});
