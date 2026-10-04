import { describe, expect, it } from 'vitest';
import { analyzeModule, extractLeadingDoc, extractModuleHeaderDoc, resolveDiagnosticCodeActions } from '../src/analyzer';
import { attachedCommentsStart, leadingDocLines, scanDocTags } from '../src/analyzer/docs/docComment';
import { moveToModule } from '../src/analyzer/refactor/moveToModule';
import { encapsulateField } from '../src/analyzer/refactor/encapsulateField';
import { applyVbaTextEdits } from '../src/analyzer/refactor/refactorTypes';

const doc = "''' <summary>Preserved description.</summary>";
const directive = "' @xlide-analysis-disable-next-member unused-procedure";
for (const endings of [['\n'], ['\r\n'], ['\r'], ['\r', '\n', '\r\n']]) {
    const join = (lines: string[]) => lines.map((line, i) => line + (i < lines.length - 1 ? endings[i % endings.length] : '')).join('');
    describe('documentation physical lines ' + JSON.stringify(endings), () => {
        it('removes a module declaration without preceding code', () => {
            const prefix = join(['Option Explicit', '']);
            const removed = join([doc, directive, 'Private unused As Long', '']);
            const suffix = join(['Public Sub P()', 'Debug.Print 1', 'End Sub', '']);
            const source = prefix + removed + suffix;
            const diagnostic = analyzeModule(source).find(d => d.code === 'unused-variable');
            expect(diagnostic).toBeDefined();
            const action = resolveDiagnosticCodeActions(source, diagnostic!).find(a => a.title === "Remove unused declaration of 'unused'");
            expect(action).toBeDefined();
            expect(applyVbaTextEdits(source, action!.edits)).toBe(prefix + suffix);
        });
        it('returns the own line start without attached comments', () => {
            const source = join(['Option Explicit', 'Private unused As Long', '']);
            expect(attachedCommentsStart(source, source.indexOf('Private'))).toBe(source.indexOf('Private'));
        });
        it('extracts docs through directives with exact source spans', () => {
            const source = join(['Option Explicit', directive, doc, directive, 'Public Sub P()', 'End Sub', '']);
            const offset = source.indexOf('Public');
            expect(extractLeadingDoc(source, offset)?.summary).toBe('Preserved description.');
            const lines = leadingDocLines(source, offset);
            expect(lines).toEqual([{start:source.indexOf(doc), directivesStart:source.indexOf(directive), textStart:source.indexOf('<summary>'), text:doc.slice(4)}]);
            const tags = scanDocTags(lines)!;
            expect(tags).toHaveLength(1);
            expect(source.slice(tags[0].open.start, tags[0].open.end)).toBe('<summary>');
            expect(attachedCommentsStart(source, offset)).toBe(source.indexOf(directive));
        });
        it('stops documentation at an ordinary comment or blank line', () => {
            for (const separator of ["' ordinary", ' ']) {
                const source = join([doc, separator, 'Public Sub P()', 'End Sub', '']);
                expect(extractLeadingDoc(source, source.indexOf('Public'))).toBeUndefined();
                expect(attachedCommentsStart(source, source.indexOf('Public'))).toBe(source.indexOf('Public'));
            }
        });
        it('reads a module header without claiming member documentation', () => {
            expect(extractModuleHeaderDoc(join([doc, 'Option Explicit', 'Public Sub P()', 'End Sub']))?.summary).toBe('Preserved description.');
            expect(extractModuleHeaderDoc(join([doc, 'Public Sub P()', 'End Sub']))).toBeUndefined();
        });
        it('moves only the documented procedure and preserves its neighbors', () => {
            const prefix = join(['Option Explicit', '']);
            const moved = join([doc, directive, 'Public Sub Build()', 'Debug.Print 1', 'End Sub']);
            const suffix = join(['Public Sub Other()', 'End Sub', '']);
            const source = prefix + moved + endings[0] + suffix;
            const target = 'Option Explicit' + endings[0];
            const result = moveToModule({source, offset:source.indexOf('Public Sub Build'), moduleName:'Reports', targetModuleName:'Helpers', otherModuleSources:{Helpers:target}});
            if (!result.ok) { throw new Error(result.reason); }
            expect(applyVbaTextEdits(source, result.edits)).toBe(prefix + suffix);
            const targetEdits = result.otherModules!.find(m => m.moduleName === 'Helpers')!.edits;
            const output = applyVbaTextEdits(target, targetEdits);
            expect(output).toContain(moved);
            if (endings.length === 1) { expect(output).toBe(target + endings[0] + moved + endings[0]); }
            expect(output.match(/Option Explicit/g)).toHaveLength(1);
        });
        it('repoints code calls without rewriting strings or comments', () => {
            const source = join(['Option Explicit', 'Public Sub Build()', 'Debug.Print "Reports.Build"', "' Reports.Build", 'End Sub', 'Public Sub Other()', 'Reports.Build', 'End Sub', '']);
            const target = 'Option Explicit' + endings[0];
            const result = moveToModule({source, offset:source.indexOf('Public Sub Build'), moduleName:'Reports', targetModuleName:'Helpers', otherModuleSources:{Helpers:target}});
            if (!result.ok) { throw new Error(result.reason); }
            expect(applyVbaTextEdits(source, result.edits)).toContain('Helpers.Build');
            const targetEdits = result.otherModules!.find(m => m.moduleName === 'Helpers')!.edits;
            const output = applyVbaTextEdits(target, targetEdits);
            expect(output).toContain('"Reports.Build"');
            expect(output).toContain("' Reports.Build");
        });
        it('runs documentation diagnostics on attached CR comments', () => {
            const source = join(['Option Explicit', doc, 'Public Sub P(ByVal x As Long)', 'Debug.Print x', 'End Sub', '']);
            const diagnostic = analyzeModule(source).find(d => d.code === 'doc-param-missing');
            expect(diagnostic).toBeDefined();
            expect(source.slice(diagnostic!.span.start, diagnostic!.span.end)).toBe('x');
            const action = resolveDiagnosticCodeActions(source, diagnostic!).find(a => a.title === "Add a <param> for 'x'")!;
            expect(action).toBeDefined();
            const output = applyVbaTextEdits(source, action.edits);
            expect(extractLeadingDoc(output, output.indexOf('Public Sub'))?.params.map(p => p.name)).toEqual(['x']);
            if (endings[0] === '\r' && endings.length === 1) { expect(output).not.toContain('\n'); }
        });
        it('removes an unknown tag with only its next-line directive', () => {
            const nextLine = "' @xlide-analysis-disable-next-line unused-variable";
            const tag = "''' <param name=\"y\">Gone.</param>";
            const source = join(['Option Explicit', doc, nextLine, tag, 'Public Sub P()', 'End Sub', '']);
            const diagnostic = analyzeModule(source).find(d => d.code === 'doc-param-unknown');
            expect(diagnostic).toBeDefined();
            const action = resolveDiagnosticCodeActions(source, diagnostic!).find(a => a.title === "Remove the <param> for 'y'")!;
            expect(action).toBeDefined();
            const output = applyVbaTextEdits(source, action.edits);
            const from = source.indexOf(nextLine), to = source.indexOf('Public');
            expect(output).toBe(source.slice(0, from) + source.slice(to));
        });
        it('moves field documentation once to the generated property', () => {
            const source = join(['Option Explicit', doc, 'Public Total As Long', '']);
            const result = encapsulateField({source, offset:source.indexOf('Total')});
            if (!result.ok) { throw new Error(result.reason); }
            const output = applyVbaTextEdits(source, result.edits);
            expect(output.startsWith('Option Explicit' + endings[0])).toBe(true);
            expect(output.match(/Preserved description/g)).toHaveLength(1);
            expect(extractLeadingDoc(output, output.indexOf('Public Property Get'))?.summary).toBe('Preserved description.');
            expect(extractLeadingDoc(output, output.indexOf('Private m_Total'))).toBeUndefined();
        });
    });
}
