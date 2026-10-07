import { describe, expect, it } from 'vitest';
import { readAttributeAnnotations } from '../src/analyzer/annotations/attributeAnnotations';
import { applyAttributeAnnotations } from '../src/analyzer/annotations/attributeRewriter';
import { joinVbaSource } from '../src/vba/moduleSource';

const endings = [['LF', ['\n']], ['CRLF', ['\r\n']], ['CR', ['\r']], ['mixed', ['\r\n', '\r', '\n']]] as const;
function text(lines: string[], separators: readonly string[]): string {
    return lines.map((line, i) => line + separators[i % separators.length]).join('');
}
describe('annotation physical lines', () => {
    it.each(endings)('binds module, variable and property legs on %s lines', (_, separators) => {
        const lines = ['Attribute VB_Name = "M"', "'@ModuleDescription(\"module\")", "'@VariableDescription(\"var\")", 'Private x As Long', "'@Description(\"get\")", 'Property Get P() As Long', 'End Property', "'@Description(\"let\")", 'Property Let P(ByVal value As Long)', 'End Property'];
        const source = text(lines, separators);
        const found = readAttributeAnnotations(source);
        expect(found).toEqual({ annotations: [
            { kind: 'ModuleDescription', line: 2, argument: 'module' },
            { kind: 'VariableDescription', line: 3, argument: 'var', target: 'x', targetLine: 4 },
            { kind: 'Description', line: 5, argument: 'get', target: 'P', targetLine: 6, targetOccurrence: 0 },
            { kind: 'Description', line: 8, argument: 'let', target: 'P', targetLine: 9, targetOccurrence: 1 },
        ], problems: [] });
        const insertions = new Map([[0, 'Attribute VB_Description = "module"'], [3, 'Attribute x.VB_VarDescription = "var"'], [5, 'Attribute P.VB_Description = "get"'], [8, 'Attribute P.VB_Description = "let"']]);
        const expected = lines.map((line, i) => {
            const ending = separators[i % separators.length];
            const inserted = insertions.get(i);
            return line + ending + (inserted ? inserted + ending : '');
        }).join('');
        const result = applyAttributeAnnotations(source, found);
        expect(result.text).toBe(expected);
        expect(result.changes).toHaveLength(4);
        expect(result.skipped).toEqual([]);
        expect(applyAttributeAnnotations(result.text, readAttributeAnnotations(result.text))).toEqual({ text: expected, changes: [], skipped: [] });
    });

    it.each(endings)('reports physical lines for placement and duplicate errors on %s', (_, separators) => {
        const source = text(["'@ModuleDescription", "'@ModuleDescription(\"valid\")", "'@ModuleDescription(\"duplicate\")", "'@Description(\"wrong\")", 'Dim x As Long', 'Sub P()', "'@Exposed", 'End Sub', "'@Description(\"dangling\")"], separators);
        const result = readAttributeAnnotations(source);
        expect(result.annotations).toEqual([{ kind: 'ModuleDescription', line: 2, argument: 'valid' }]);
        expect(result.problems.map(problem => problem.line)).toEqual([1, 3, 4, 7, 9]);
        expect(result.problems.map(problem => problem.message)).toEqual([
            "'@ModuleDescription needs the text to write, in brackets: '@ModuleDescription(\"...\").",
            "'@ModuleDescription appears more than once; the first one counts.",
            "'@Description describes a procedure, and 'x' is a variable. Use '@VariableDescription for a variable.",
            "'@Exposed is a module annotation and belongs in the declarations section, above the first procedure.",
            "'@Description is above nothing, so there is nothing to bind it to.",
        ]);
    });

    it('recognizes a CR-only body assembled with the actual save-path join helper', () => {
        const body = "'@Description(\"saved\")\rSub P()\rEnd Sub\r";
        const source = joinVbaSource('Attribute VB_Name = "M"', body);
        const annotations = readAttributeAnnotations(source);
        expect(annotations.annotations).toEqual([{ kind: 'Description', line: 2, argument: 'saved', target: 'P', targetLine: 3, targetOccurrence: 0 }]);
        expect(annotations.problems).toEqual([]);
        const result = applyAttributeAnnotations(source, annotations);
        expect(result.text).toBe(source.replace('Sub P()\r', 'Sub P()\rAttribute P.VB_Description = "saved"\r'));
        expect(result.changes).toEqual([{ target: 'P', attribute: 'VB_Description', to: '"saved"' }]);
        expect(result.skipped).toEqual([]);
    });
});
