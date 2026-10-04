import { beforeEach, describe, expect, it, vi } from 'vitest';
const calls = vi.hoisted(() => ({ declarations: vi.fn(), references: vi.fn() }));
vi.mock('../src/analyzer/flow/procedureLabels', async (importOriginal) => {
    const real = await importOriginal<typeof import('../src/analyzer/flow/procedureLabels')>();
    return { ...real, statementLabelDeclarations: calls.declarations.mockImplementation(real.statementLabelDeclarations),
        statementLabelReferences: calls.references.mockImplementation(real.statementLabelReferences) };
});
import { checkStatementForms } from '../src/analyzer/diagnostics/rules/statementForms';
import { parseModule } from '../src/analyzer/parser/parseModule';
import { buildModuleSymbols } from '../src/analyzer/symbols/buildModuleSymbols';
import type { MemberCompletionContext } from '../src/analyzer/completion/memberAccess';
const project: MemberCompletionContext = { projectClassMembers: [{ name: 'Foo', moduleName: 'Foo', kind: 'standardModule', members: [] }] };
function run(body: string, context: MemberCompletionContext = {}) {
    const source = 'Option Explicit\nSub Main()\nDim total As Long\n' + body + '\nEnd Sub';
    const mod = parseModule(source), symbols = buildModuleSymbols('M', 'standard', source, { parsedModule: mod });
    const push = vi.fn();
    checkStatementForms(source, mod, symbols, undefined, undefined, push, context);
    return push.mock.calls;
}
function counts() { return [calls.declarations.mock.calls.length, calls.references.mock.calls.length]; }
beforeEach(() => { calls.declarations.mockClear(); calls.references.mockClear(); });

describe('statement-form label collection', () => {
    it.each([false, true])('skips 1000 unrelated assignments with project metadata=%s', withProject => {
        const diagnostics = run(Array.from({ length: 1000 }, (_, index) => 'total=' + index).join('\n'), withProject ? project : {});
        expect(diagnostics).toEqual([]);
        expect(counts()).toEqual([0, 0]);
    });

    it('skips qualified names that cannot trigger the bare-module rule', () => {
        run('Call Foo.Bar\ntotal=Foo.Bar()', project);
        expect(counts()).toEqual([0, 0]);
    });

    it('skips a name shadowed by a local', () => {
        const diagnostics = run('Dim Foo As Long\nFoo=1\ntotal=Foo', project);
        expect(diagnostics).toEqual([]);
        expect(counts()).toEqual([0, 0]);
    });

    it('shares one label collection when the same span has two module-name hits', () => {
        const diagnostics = run('Foo Foo', project);
        expect(diagnostics.filter(([rule]) => rule === 'malformedStatement')).toHaveLength(2);
        expect(counts()).toEqual([1, 1]);
    });

    it('preserves label declarations and references in their own namespace', () => {
        const diagnostics = run('GoTo Foo\nFoo:', project);
        expect(diagnostics).toEqual([]);
        expect(counts()).toEqual([2, 2]);
    });
});
