import { beforeEach, describe, expect, it, vi } from 'vitest';
const calls = vi.hoisted(() => ({ resolve: vi.fn() }));
vi.mock('../src/analyzer/completion/memberAccess', async (importOriginal) => {
    const real = await importOriginal<typeof import('../src/analyzer/completion/memberAccess')>();
    return { ...real, projectClassMemberAt: calls.resolve.mockImplementation(real.projectClassMemberAt) };
});
import { checkStatementForms } from '../src/analyzer/diagnostics/rules/statementForms';
import { parseModule } from '../src/analyzer/parser/parseModule';
import { buildModuleSymbols } from '../src/analyzer/symbols/buildModuleSymbols';
import type { MemberCompletionContext } from '../src/analyzer/completion/memberAccess';
import type { VbaProjectClassMembers } from '../src/analyzer/symbols/symbolModel';
function run(body: string, context: MemberCompletionContext = {}) {
    const source = 'Option Explicit\nSub Main()\nDim c As C\nDim result As Long\n' + body + '\nEnd Sub';
    const mod = parseModule(source), symbols = buildModuleSymbols('M', 'standard', source, { parsedModule: mod });
    const push = vi.fn();
    checkStatementForms(source, mod, symbols, undefined, undefined, push, context);
    return push.mock.calls;
}
const type = (members: VbaProjectClassMembers['members'], kind: VbaProjectClassMembers['kind'] = 'class'): VbaProjectClassMembers => ({ name: 'C', moduleName: 'C', kind, members });
const sub = { name: 'DoIt', kind: 'method' as const, moduleName: 'C', sub: true };
beforeEach(() => { calls.resolve.mockClear(); });

describe('statement-form class Sub candidate lookup', () => {
    it.each([
        {},
        { projectClassMembers: [type([{ name: 'Value', moduleName: 'C', kind: 'property', returns: 'Long' }])] },
        { projectClassMembers: [type([sub])] },
        { projectClassMembers: [type([sub], 'standardModule')] },
    ] satisfies MemberCompletionContext[])('skips 1000 member names absent from class Subs (%j)', context => {
        expect(run(Array.from({ length: 1000 }, () => 'result=c.Value').join('\n'), context)).toEqual([]);
        expect(calls.resolve.mock.calls.length).toBe(0);
    });

    it('resolves matching Sub names and reports the actual class member', () => {
        const diagnostics = run('result=c.DoIt()\nresult=c.dOiT()', { projectClassMembers: [type([sub])] });
        expect(diagnostics.map(args => args[0])).toEqual(['subUsedAsValue', 'subUsedAsValue']);
        expect(calls.resolve.mock.calls.length).toBe(2);
    });

    it('preserves first member matching and does not flag a Function before a same-name Sub', () => {
        const diagnostics = run('result=c.DoIt()', { projectClassMembers: [type([{ ...sub, sub: false, returns: 'Long' }, sub])] });
        expect(diagnostics).toEqual([]);
        expect(calls.resolve.mock.calls.length).toBe(1);
    });

    it('does not scan unrelated class metadata for an already found Sub name', () => {
        let reads = 0;
        const unrelated = Array.from({ length: 1000 }, (_, i) => ({ name: 'Other' + i, moduleName: 'Other' + i, kind: 'class' as const,
            members: [{ name: 'Unused', kind: 'method' as const, moduleName: 'Other' + i, get sub() { reads++; return true; } }] }));
        const diagnostics = run('result=c.DoIt()\nresult=c.DoIt()', { projectClassMembers: [type([sub]), ...unrelated] });
        expect(diagnostics.map(args => args[0])).toEqual(['subUsedAsValue', 'subUsedAsValue']);
        expect(reads).toBe(0);
    });

    it('finds later candidates after a hit and retains names after a negative lookup', () => {
        const diagnostics = run('result=c.DoIt()\nresult=c.Finish()\nresult=c.Value\nresult=c.DoIt()',
            { projectClassMembers: [type([sub, { ...sub, name: 'Finish' }])] });
        expect(diagnostics.map(args => args[0])).toEqual(['subUsedAsValue', 'subUsedAsValue', 'subUsedAsValue']);
        expect(calls.resolve.mock.calls.length).toBe(3);
    });

    it('uses current metadata on each rule invocation', () => {
        const members: VbaProjectClassMembers['members'] = [];
        const context = { projectClassMembers: [type(members)] };
        expect(run('result=c.DoIt()', context)).toEqual([]);
        members.push(sub);
        expect(run('result=c.DoIt()', context).map(args => args[0])).toEqual(['subUsedAsValue']);
    });
});
