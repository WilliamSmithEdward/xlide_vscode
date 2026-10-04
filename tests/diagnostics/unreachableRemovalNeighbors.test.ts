import { describe, expect, it } from 'vitest';
import { analyzeModule, resolveDiagnosticCodeActions } from '../../src/analyzer';
import { applyVbaTextEdits } from '../../src/analyzer/refactor/refactorTypes';
function remove(source: string): string {
    const found = analyzeModule(source).filter(d => d.code === 'unreachable-code');
    expect(found).toHaveLength(1);
    const diagnostic = found[0];
    const edit = diagnostic.data?.removeUnreachableCode?.edit;
    expect(edit).toBeDefined();
    const actions = resolveDiagnosticCodeActions(source, diagnostic);
    const action = actions.find(action => action.title === 'Remove unreachable code');
    expect(action).toBeDefined();
    expect(action!.edits).toEqual([edit]);
    return applyVbaTextEdits(source, action!.edits);
}
describe('unreachable removal preserves live neighbors', () => {
    for (const eol of ['\n', '\r\n', '\r']) {
        it('keeps Exit Do on a shared line with ' + JSON.stringify(eol), () => {
            const source = ['Sub P()', 'Do', 'Exit Do: Debug.Print "dead"', 'Loop', 'Debug.Print "live"', 'End Sub', ''].join(eol);
            expect(remove(source)).toBe(['Sub P()', 'Do', 'Exit Do', 'Loop', 'Debug.Print "live"', 'End Sub', ''].join(eol));
        });
        it('keeps a live terminator before a multi-line dead range with ' + JSON.stringify(eol), () => {
            const source = ['Sub P()', 'Exit Sub: Debug.Print "one"', 'Debug.Print "two"', 'End Sub', ''].join(eol);
            expect(remove(source)).toBe(['Sub P()', 'Exit Sub', 'End Sub', ''].join(eol));
        });
        it('preserves trailing and interior comments with ' + JSON.stringify(eol), () => {
            const source = ['Sub P()', 'Exit Sub: Debug.Print "one" \'first note', '  \'interior note', 'Debug.Print "two" \'last note', 'End Sub', ''].join(eol);
            expect(remove(source)).toBe(['Sub P()', "Exit Sub \'first note", "  \'interior note", " \'last note", 'End Sub', ''].join(eol));
        });
        it('keeps a landing label after dead statements with ' + JSON.stringify(eol), () => {
            const source = ['Sub P()', 'GoTo Live: Debug.Print "dead"', 'Live:', 'Debug.Print "live"', 'End Sub', ''].join(eol);
            expect(remove(source)).toBe(['Sub P()', 'GoTo Live', 'Live:', 'Debug.Print "live"', 'End Sub', ''].join(eol));
        });
        it('removes only whole standalone dead lines with ' + JSON.stringify(eol), () => {
            const source = ['Sub P()', 'Exit Sub', 'Debug.Print "one"', 'Debug.Print "two"', 'End Sub', ''].join(eol);
            expect(remove(source)).toBe(['Sub P()', 'Exit Sub', 'End Sub', ''].join(eol));
        });
        it('retains the refusal to delete a dead block declaring a local with ' + JSON.stringify(eol), () => {
            const source = ['Sub P()', 'Exit Sub', 'If True Then', 'Dim x As Long', 'Debug.Print x', 'End If', 'End Sub', ''].join(eol);
            const found = analyzeModule(source).filter(d => d.code === 'unreachable-code');
            expect(found).toHaveLength(1);
            expect(found[0].data?.removeUnreachableCode).toBeUndefined();
        });
    }
});
