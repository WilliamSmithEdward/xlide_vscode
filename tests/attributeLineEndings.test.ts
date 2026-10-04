import { describe, expect, it } from 'vitest';
import { readAttributeAnnotations } from '../src/analyzer/annotations/attributeAnnotations';
import { applyAttributeAnnotations } from '../src/analyzer/annotations/attributeRewriter';

const mixed = 'Attribute VB_Name = "M"\r\nAttribute VB_Description = "old"\n\'@ModuleDescription("old")\r\nSub P()\nEnd Sub\r\n';

describe('attribute writer line endings', () => {
    it('leaves an already aligned accepted annotation byte for byte unchanged', () => {
        const annotations = readAttributeAnnotations(mixed);
        expect(annotations.annotations).toHaveLength(1);
        expect(annotations.problems).toEqual([]);
        expect(applyAttributeAnnotations(mixed, annotations)).toEqual({ text: mixed, changes: [], skipped: [] });
    });

    it('changes only the requested attribute value in mixed source', () => {
        const source = mixed.replace('@ModuleDescription("old")', '@ModuleDescription("new")');
        expect(applyAttributeAnnotations(source, readAttributeAnnotations(source))).toEqual({
            text: source.replace('VB_Description = "old"', 'VB_Description = "new"'),
            changes: [{ target: 'module', attribute: 'VB_Description', from: '"old"', to: '"new"' }], skipped: [],
        });
    });

    it('preserves unrelated endings when inserting module, variable and continued procedure attributes', () => {
        const source = 'Attribute VB_Name = "M"\r\n\'@ModuleDescription("module")\n\'@VariableDescription("var")\r\nPrivate x As Long\n\'@Description("proc")\r\nSub P( _\n ByVal n As Long)\r\nAttribute Other.VB_Description = "leave"\nEnd Sub';
        const expected = source.replace('Attribute VB_Name = "M"\r\n', 'Attribute VB_Name = "M"\r\nAttribute VB_Description = "module"\r\n')
            .replace('Private x As Long\n', 'Private x As Long\nAttribute x.VB_VarDescription = "var"\n')
            .replace(' ByVal n As Long)\r\n', ' ByVal n As Long)\r\nAttribute P.VB_Description = "proc"\r\n');
        const result = applyAttributeAnnotations(source, readAttributeAnnotations(source));
        expect(result.text).toBe(expected);
        expect(result.changes).toHaveLength(3);
        expect(result.skipped).toEqual([]);
        expect(applyAttributeAnnotations(result.text, readAttributeAnnotations(result.text))).toEqual({ text: expected, changes: [], skipped: [] });
    });

    it.each(['\n', '\r\n', '\r'])('preserves %j with and without a final newline', eol => {
        for (const final of ['', eol]) {
            const source = ['Attribute VB_Name = "M"', 'Sub P()', 'End Sub'].join(eol) + final;
            const result = applyAttributeAnnotations(source, { problems: [], annotations: [{ kind: 'Description', line: 1, target: 'P', argument: 'doc' }] });
            expect(result.text).toBe(source.replace('Sub P()' + eol, 'Sub P()' + eol + 'Attribute P.VB_Description = "doc"' + eol));
            expect(result.skipped).toEqual([]);
            expect(result.changes).toHaveLength(1);
        }
    });

    it.each(['Attribute VB_Name = "M"', 'Attribute VB_Name = "M"\r\nSub P()'])('inserts after the last unterminated line: %j', source => {
        const member = source.endsWith('Sub P()');
        const result = applyAttributeAnnotations(source, { problems: [], annotations: [member
            ? { kind: 'Description', line: 1, target: 'P', argument: 'doc' }
            : { kind: 'ModuleDescription', line: 1, argument: 'doc' }] });
        expect(result.text).toBe(source + (member ? '\r\nAttribute P.VB_Description = "doc"' : '\nAttribute VB_Description = "doc"'));
        expect(result.changes).toHaveLength(1);
        expect(result.skipped).toEqual([]);
    });

    it('preserves mixed source when all targets are skipped or no annotations exist', () => {
        const result = applyAttributeAnnotations(mixed, { problems: [], annotations: [{ kind: 'Description', line: 1, target: 'Missing', argument: 'doc' }] });
        expect(result.text).toBe(mixed);
        expect(result.changes).toEqual([]);
        expect(result.skipped).toHaveLength(1);
        expect(applyAttributeAnnotations(mixed, { problems: [], annotations: [] })).toEqual({ text: mixed, changes: [], skipped: [] });
    });
});
