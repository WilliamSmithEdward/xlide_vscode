import { describe, expect, it } from 'vitest';
import { analyzeModule } from '../src/analyzer';
import {
    buildVbaProjectIndex,
    projectAnalysisOptionsForModule,
    type VbaProjectModuleInput,
} from '../src/vbaProjectAnalysis';

// Issue #26: `EntryForm.NoSuchControl` compiles nowhere - the VBE refuses an
// unknown member on an early-bound form receiver - yet the analyzer said
// nothing. A form's member surface is code-behind plus designer controls plus
// the MSForms UserForm base, and it proves absence exactly when the control
// list is authoritative: supplied with the module by a designer-reading host,
// or parsed from a `.frm` designer header the source itself carries. A form
// whose designer nobody has read stays out of absence claims.

const FORM_CODE = [
    'Public Sub Accept()',
    'End Sub',
    '',
].join('\n');

const FORM_HEADER_SOURCE = [
    'VERSION 5.00',
    'Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} EntryForm',
    '   Caption         =   "Entry"',
    '   Begin {8BD21D10-EC42-11CE-9E0D-00AA006002F3} NameBox',
    '   End',
    'End',
    FORM_CODE,
].join('\n');

function analyzeCaller(
    callerSource: string,
    form: Partial<VbaProjectModuleInput> & { source: string },
): ReturnType<typeof analyzeModule> {
    const modules: VbaProjectModuleInput[] = [
        { moduleName: 'Caller', type: 'standard', source: callerSource },
        { moduleName: 'EntryForm', type: 'userform', ...form },
    ];
    const project = buildVbaProjectIndex(modules);
    const options = projectAnalysisOptionsForModule(project, 'Caller');
    return analyzeModule(callerSource, { moduleName: 'Caller', ...options });
}

function memberNotFound(diagnostics: ReturnType<typeof analyzeModule>): string[] {
    return diagnostics
        .filter((hit) => hit.code === 'member-not-found')
        .map((hit) => hit.message);
}

describe('member-not-found on form receivers (issue #26)', () => {
    const CALLER = [
        'Option Explicit',
        '',
        'Public Sub MissingControl()',
        '    EntryForm.NoSuchControl.Text = "?"',
        'End Sub',
        '',
    ].join('\n');

    it('flags a member the form does not have when the host supplied the controls', () => {
        const hits = memberNotFound(analyzeCaller(CALLER, {
            source: FORM_CODE,
            implicitMembers: [{ name: 'NameBox', type: 'MSForms.TextBox' }],
        }));
        expect(hits).toEqual(["Method or data member not found: 'EntryForm.NoSuchControl'."]);
    });

    it('flags through a source-carried designer header with no host list', () => {
        const hits = memberNotFound(analyzeCaller(CALLER, { source: FORM_HEADER_SOURCE }));
        expect(hits).toEqual(["Method or data member not found: 'EntryForm.NoSuchControl'."]);
    });

    it('an authoritative EMPTY control list still proves absence', () => {
        const hits = memberNotFound(analyzeCaller(CALLER, {
            source: FORM_CODE,
            implicitMembers: [],
        }));
        expect(hits).toEqual(["Method or data member not found: 'EntryForm.NoSuchControl'."]);
    });

    it('stays silent when nobody has read the designer', () => {
        // No implicitMembers and no header: the control list is unknown, so
        // absence cannot be proven.
        const hits = memberNotFound(analyzeCaller(CALLER, { source: FORM_CODE }));
        expect(hits).toEqual([]);
    });

    it('stays silent behind a real export header that defers to an .frx blob', () => {
        // A genuine VBA .frm export keeps its controls in the binary .frx
        // behind an OleObjectBlob line (measured on an Excel-authored form),
        // so this header proves nothing about the control list.
        const realExportHeader = [
            'VERSION 5.00',
            'Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} EntryForm ',
            '   Caption         =   "Entry"',
            '   ClientHeight    =   3015',
            '   OleObjectBlob   =   "EntryForm.frx":0000',
            'End',
            FORM_CODE,
        ].join('\n');
        const hits = memberNotFound(analyzeCaller(CALLER, { source: realExportHeader }));
        expect(hits).toEqual([]);
    });

    it('never flags members the form does have', () => {
        const caller = [
            'Option Explicit',
            '',
            'Public Sub UsesRealMembers()',
            '    EntryForm.NameBox.Text = "x"',   // designer control
            '    EntryForm.Accept',               // code-behind member
            '    EntryForm.Show',                 // UserForm base (VBA extender)
            '    EntryForm.Caption = "t"',        // UserForm base (MSForms)
            '    EntryForm.Controls.Clear',       // UserForm base collection
            'End Sub',
            '',
        ].join('\n');
        const hits = memberNotFound(analyzeCaller(caller, {
            source: FORM_CODE,
            implicitMembers: [{ name: 'NameBox', type: 'MSForms.TextBox' }],
        }));
        expect(hits).toEqual([]);
    });

    it('leaves late-bound receivers alone', () => {
        const caller = [
            'Option Explicit',
            '',
            'Public Sub LateBound(ByVal anything As Object)',
            '    anything.NoSuchControl.Text = "?"',
            'End Sub',
            '',
        ].join('\n');
        const hits = memberNotFound(analyzeCaller(caller, {
            source: FORM_CODE,
            implicitMembers: [{ name: 'NameBox', type: 'MSForms.TextBox' }],
        }));
        expect(hits).toEqual([]);
    });
});
