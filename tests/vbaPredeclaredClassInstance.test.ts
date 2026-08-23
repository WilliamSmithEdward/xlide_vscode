import { describe, expect, it } from 'vitest';
import * as path from 'path';
import { analyzeModule } from '../src/analyzer';
import { classifyModuleType, readModules } from '../src/vba/workbookService';
import {
    buildVbaProjectIndex,
    projectAnalysisOptionsForModule,
    type VbaProjectModuleInput,
} from '../src/vbaProjectAnalysis';

// Issue #47: a VBA class module's name is a TYPE, not a value. `Ticket.Change`
// where Ticket is a plain class is the same "Variable not defined" the compiler
// raises for any other undeclared name - `Debug > Compile` refuses the project.
// The analyzer stayed silent because the qualifier skip accepted ANY project
// member surface as a legal bare receiver.
//
// What makes a module's own name usable as a value is a DEFAULT INSTANCE:
// standard modules are namespaces, documents and UserForms always have one, and
// a class has one only with `Attribute VB_PredeclaredId = True` - which is
// invisible in the code pane and lives in the exported header. So the bit has
// three states, and only a vouched-for FALSE reports. A host that never read
// the header leaves it unknown, and unknown must stay silent: fullBuild.xlsm
// alone carries 12 predeclared classes (the stdVBA library), every one of which
// would go red on a guess.

const TICKET_BODY = ['Option Explicit', 'Public Sub ChangeTest()', 'End Sub', ''].join('\n');

/** A `.cls` header as the VBE exports it, which standalone files carry. */
function clsSource(name: string, predeclared: boolean): string {
    return [
        `Attribute VB_Name = "${name}"`,
        'Attribute VB_GlobalNameSpace = False',
        'Attribute VB_Creatable = False',
        `Attribute VB_PredeclaredId = ${predeclared ? 'True' : 'False'}`,
        'Attribute VB_Exposed = False',
        TICKET_BODY,
    ].join('\n');
}

function undeclaredIn(
    callerBody: string,
    ticket: Partial<VbaProjectModuleInput> = {},
): string[] {
    const caller = [
        'Option Explicit',
        '',
        'Public Sub Probe()',
        '    Dim o As Object',
        '    Dim t As Object',
        callerBody,
        'End Sub',
        '',
    ].join('\n');
    const modules: VbaProjectModuleInput[] = [
        { moduleName: 'Caller', type: 'standard', source: caller },
        { moduleName: 'Ticket', type: 'class', source: TICKET_BODY, ...ticket },
        {
            moduleName: 'Helper',
            type: 'standard',
            source: 'Option Explicit\nPublic Sub Go()\nEnd Sub\n',
        },
        {
            moduleName: 'Sheet1',
            type: 'document',
            source: 'Option Explicit\nPublic Sub Run()\nEnd Sub\n',
        },
    ];
    const project = buildVbaProjectIndex(modules);
    const options = projectAnalysisOptionsForModule(project, 'Caller');
    return analyzeModule(caller, { moduleName: 'Caller', ...options })
        .filter((hit) => hit.code === 'undeclared-variable')
        .map((hit) => hit.message);
}

describe('a class name used as an instance (issue #47)', () => {
    it('reports a plain class module used as a qualifier', () => {
        expect(undeclaredIn('    Ticket.ChangeTest', { predeclared: false }))
            .toEqual([
                "Variable not defined: 'Ticket'. Declare it before using it, or remove Option Explicit.",
            ]);
    });

    it('reports it case-insensitively, as VBA resolves it', () => {
        expect(undeclaredIn('    ticket.ChangeTest', { predeclared: false })).toHaveLength(1);
    });

    it('stays silent for a class with a default instance', () => {
        expect(undeclaredIn('    Ticket.ChangeTest', { predeclared: true })).toEqual([]);
    });

    it('stays silent when nobody read the attribute header', () => {
        // The safety net. Absent is not "no": a host that cannot see the
        // attribute must not turn every predeclared class red.
        expect(undeclaredIn('    Ticket.ChangeTest', {})).toEqual([]);
    });

    it('leaves the legal uses of a class NAME alone in every state', () => {
        for (const predeclared of [false, true, undefined]) {
            const at = { predeclared } as Partial<VbaProjectModuleInput>;
            expect(undeclaredIn('    Dim d As Ticket', at), `As, ${predeclared}`).toEqual([]);
            expect(undeclaredIn('    Set t = New Ticket', at), `New, ${predeclared}`).toEqual([]);
            expect(
                undeclaredIn('    If TypeOf o Is Ticket Then\n    End If', at),
                `TypeOf, ${predeclared}`,
            ).toEqual([]);
        }
    });

    it('leaves standard-module and document qualifiers alone', () => {
        expect(undeclaredIn('    Helper.Go', { predeclared: false })).toEqual([]);
        expect(undeclaredIn('    Sheet1.Run', { predeclared: false })).toEqual([]);
    });

    it('no longer calls a predeclared class undefined when read bare', () => {
        // The same bit fixes the mirror defect: `Set x = stdArray` names a real
        // value when the class has a default instance.
        expect(undeclaredIn('    Set t = Ticket', { predeclared: true })).toEqual([]);
        expect(undeclaredIn('    Set t = Ticket', { predeclared: false })).toHaveLength(1);
    });

    it('reads the attribute from a standalone export when no host answers', () => {
        // A `.cls` on disk carries its own header, so the index can answer
        // without a designer-reading host in the loop.
        expect(undeclaredIn('    Ticket.ChangeTest', { source: clsSource('Ticket', true) }))
            .toEqual([]);
        expect(undeclaredIn('    Ticket.ChangeTest', { source: clsSource('Ticket', false) }))
            .toHaveLength(1);
    });

    it('prefers the host answer over the module text', () => {
        expect(
            undeclaredIn('    Ticket.ChangeTest', {
                source: clsSource('Ticket', false),
                predeclared: true,
            }),
        ).toEqual([]);
    });
});

describe('the attribute header itself', () => {
    it('still classifies a predeclared+exposed module as a document', () => {
        // Guards the reader this bit shares: booleans are written UNQUOTED, and
        // reading only the quoted form made every one of them answer "".
        const header = [
            'Attribute VB_Name = "Sheet1"',
            'Attribute VB_PredeclaredId = True',
            'Attribute VB_Exposed = True',
        ].join('\r\n');
        expect(classifyModuleType('Anything', header)).toBe('document');
    });

    it('does not mistake a plain class for a document', () => {
        const header = [
            'Attribute VB_Name = "Ticket"',
            'Attribute VB_PredeclaredId = False',
            'Attribute VB_Exposed = False',
        ].join('\r\n');
        expect(classifyModuleType('Ticket', header)).toBe('standard');
    });
});

describe('reading the bit out of a real workbook', () => {
    // FormFixture.xlsm was authored by live Excel, so its headers are the ones
    // the VBE actually writes rather than any shape this repo invented.
    const FIXTURE = path.join(__dirname, 'fixtures', 'binaries', 'FormFixture.xlsm');

    it('answers for every module the VBE gives a default instance', () => {
        const byName = new Map(readModules(FIXTURE, false).map((m) => [m.name, m]));
        expect(byName.get('ThisWorkbook')?.predeclared).toBe(true);
        expect(byName.get('Sheet1')?.predeclared).toBe(true);
        expect(byName.get('FrmPicker')?.predeclared).toBe(true);
    });

    it('leaves a standard module unanswered rather than guessing', () => {
        // A `.bas` header carries no VB_PredeclaredId at all. Nothing reads the
        // bit for a standard module, and inventing `false` for one would be a
        // claim the header never made.
        const byName = new Map(readModules(FIXTURE, false).map((m) => [m.name, m]));
        expect(byName.get('XlideFormProbe')?.predeclared).toBeUndefined();
    });
});
