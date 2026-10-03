import { describe, expect, it, vi } from 'vitest';
import { checkDeclarationOrder } from '../../src/analyzer/diagnostics/rules/declarationOrder';
import { parseModule } from '../../src/analyzer/parser/parseModule';
import type { VbaProjectClassMembers } from '../../src/analyzer/symbols/symbolModel';

function surface(name: string, references: string[], moduleName = 'Other'): VbaProjectClassMembers {
    return { name, moduleName, kind: 'userType', members: references.map((returns,i)=>({ name:'field'+i, kind:'property', moduleName, returns })) };
}
const source = 'Public Type Root\n    first As Entry\n    second As Entry\nEnd Type';
function check(project: VbaProjectClassMembers[]) {
    const push=vi.fn();
    checkDeclarationOrder(source,parseModule(source),'Module',undefined,project,undefined,push);
    return push;
}

describe('cross-module type dependency queue', () => {
    it('does not enqueue shared descendant types for every incoming edge', () => {
        const width=60;
        const leaves=Array.from({length:width},(_,i)=>'Leaf'+i);
        const branches=Array.from({length:width},(_,i)=>'Branch'+i);
        const project=[surface('Entry',branches), ...branches.map(name=>surface(name,leaves)), ...leaves.map(name=>surface(name,['Long']))];
        const module=parseModule(source);
        let queued=0;
        const original=Array.prototype.push;
        Array.prototype.push=function(this: unknown[], ...items: unknown[]) {
            queued+=items.filter(item=>typeof item==='object' && item!==null && 'kind' in item && item.kind==='userType').length;
            return original.apply(this,items);
        };
        const push=vi.fn();
        try {
            checkDeclarationOrder(source,module,'Module',undefined,project,undefined,push);
        } finally { Array.prototype.push=original; }
        expect(push).not.toHaveBeenCalled();
        // Two fields traverse the same graph: queue work must follow node count,
        // rather than repeat every shared leaf for each of its incoming edges.
        expect(queued).toBeLessThan(project.length*4);
    });

    it('preserves one cycle diagnostic per field, including its first external type and span', () => {
        const project=[surface('Entry',['Left','Right']),surface('Left',['Shared']),surface('Right',['Shared']),surface('Shared',['Entry','rOoT'])];
        const push=check(project);
        expect(push.mock.calls.map(call=>call[0])).toEqual(['circularDeclarationDependency','circularDeclarationDependency']);
        expect(push.mock.calls.every(call=>call[1].includes('Other.Entry'))).toBe(true);
        expect(push.mock.calls.map(call=>source.slice(call[2].start,call[2].end))).toEqual(['Entry','Entry']);
    });

    it('does not traverse ambiguous names or a type owned by the current module', () => {
        expect(check([surface('Entry',['Leaf']),surface('Leaf',['Root']),surface('LEAF',['Root'],'Third')])).not.toHaveBeenCalled();
        expect(check([surface('Entry',['Local']),surface('Local',['Root'],'Module')])).not.toHaveBeenCalled();
    });

    it('uses changed project metadata on the next rule pass', () => {
        const leaf=surface('Leaf',['Long']);
        const project=[surface('Entry',['Leaf']),leaf];
        expect(check(project)).not.toHaveBeenCalled();
        leaf.members[0].returns='Root';
        expect(check(project)).toHaveBeenCalledTimes(2);
    });
});
