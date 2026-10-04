import { expect, it } from 'vitest';
import { readAttributeAnnotations } from '../src/analyzer/annotations/attributeAnnotations';

function counted(source: string) {
    const push = Array.prototype.push;
    let reads = 0;
    Array.prototype.push = function(...items) {
        for (const item of items) {
            if (item && typeof item === 'object' && typeof item.kind === 'string' && typeof item.line === 'number'
                && (typeof item.target === 'string' || ['ModuleDescription', 'PredeclaredId', 'Exposed'].includes(item.kind))) {
                const kind = item.kind;
                // Accepted annotations are never mutated by the reader. Observe
                // their stable kind fields during duplicate checks.
                Object.defineProperty(item, 'kind', { get() { reads++; return kind; } });
                Object.freeze(item);
            }
        }
        return push.apply(this, items);
    };
    try {
        const result = readAttributeAnnotations(source);
        return { result, reads };
    } finally { Array.prototype.push = push; }
}

it.each(['procedures', 'variables'])('bounds accepted-kind reads for distinct %s', shape => {
    const count = 1000;
    const source = Array.from({ length: count }, (_, i) => shape === 'procedures'
        ? "'@Description(\"desc\")\nPublic Sub P" + i + '()\nEnd Sub\n'
        : "'@VariableDescription(\"desc\")\nPublic v" + i + ' As Long\n').join('');
    const measured = counted(source);
    const annotations = Array.from({ length: count }, (_, i) => ({
        kind: shape === 'procedures' ? 'Description' : 'VariableDescription',
        line: i * (shape === 'procedures' ? 3 : 2) + 1,
        argument: 'desc', target: (shape === 'procedures' ? 'P' : 'v') + i,
        targetLine: i * (shape === 'procedures' ? 3 : 2) + 2,
        ...(shape === 'procedures' ? { targetOccurrence: 0 } : {}),
    }));
    expect(measured.result).toEqual({ annotations, problems: [] });
    expect(measured.reads).toBeLessThan(20 * count);
});

it('does not scan unrelated variable annotations for module duplicates', () => {
    const count = 1000;
    const source = Array.from({ length: count }, (_, i) => "'@VariableDescription(\"desc\")\nPublic v" + i + ' As Long\n').join('')
        + Array(100).fill("'@ModuleDescription(\"module\")\n").join('');
    const measured = counted(source);
    expect(measured.result.annotations).toHaveLength(count + 1);
    expect(measured.result.annotations[count]).toEqual({ kind: 'ModuleDescription', line: 2001, argument: 'module' });
    expect(measured.result.problems).toEqual(Array.from({ length: 99 }, (_, i) => ({
        line: 2002 + i, message: "'@ModuleDescription appears more than once; the first one counts.",
    })));
    expect(measured.reads).toBeLessThan(20 * count);
});

it('keeps variable duplicate spelling exact across declarations', () => {
    const source = "'@VariableDescription(\"one\")\nDim x As Long\n'@VariableDescription(\"two\")\nDim x As Long\n'@VariableDescription(\"three\")\nDim X As Long\n";
    const result = readAttributeAnnotations(source);
    expect(result.annotations.map(annotation => [annotation.target, annotation.argument])).toEqual([['x', 'one'], ['X', 'three']]);
    expect(result.problems).toEqual([{ line: 3, message: "'@VariableDescription appears more than once above 'x'; the first one counts." }]);
});

it('does not reserve a procedure kind for an invalid annotation', () => {
    const result = readAttributeAnnotations("'@Description\n'@Description(\"ok\")\nSub P()\nEnd Sub\n");
    expect(result.annotations).toEqual([{ kind: 'Description', line: 2, argument: 'ok', target: 'P', targetLine: 3, targetOccurrence: 0 }]);
    expect(result.problems).toHaveLength(1);
    expect(result.problems[0].line).toBe(1);
    expect(result.problems[0].message).toContain('needs the text');
});

it('keeps property legs and global DefaultMember validation distinct', () => {
    const source = "'@Description(\"get\")\nProperty Get P() As Long\nEnd Property\n'@Description(\"let\")\nProperty Let P(ByVal v As Long)\nEnd Property\n'@DefaultMember\n'@DefaultMember\nSub D()\nEnd Sub\n";
    const result = readAttributeAnnotations(source);
    expect(result.annotations.map(annotation => [annotation.kind, annotation.argument, annotation.targetOccurrence])).toEqual([['Description', 'get', 0], ['Description', 'let', 1], ['DefaultMember', undefined, 0]]);
    expect(result.problems).toEqual([{ line: 8, message: "'@DefaultMember appears again; a class has one default member, and line 7 already names it." }]);
});
