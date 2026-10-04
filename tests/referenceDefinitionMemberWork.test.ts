import { expect, it } from 'vitest';
import { collectSymbolReferences } from '../src/vbaReferenceResolution';
import { buildVbaProjectIndex } from '../src/vbaProjectAnalysis';
import { resolveMemberDefinitionsAt } from '../src/analyzer/completion/memberAccess';
import type { HostMember, HostObjectModel } from '../src/analyzer/host/excelObjectModel';

it.each(['first', 'last'])('bounds metadata reads in actual collection (%s member)', position => {
    const count = 1000, name = position === 'first' ? 'P0' : 'P999';
    const modules = [
        { moduleName: 'Library', source: Array.from({ length: count }, (_, i) => 'Public Sub P' + i + '()\nEnd Sub\n').join('') },
        { moduleName: 'Caller', source: 'Sub UseIt()\n' + Array(count).fill('    Library.' + name).join('\n') + '\nEnd Sub\n' },
    ];
    const project = buildVbaProjectIndex(modules);
    const byModule = new Map(modules.map(mod => [mod.moduleName.toLowerCase(), mod]));
    let reads = 0;
    const surfaces = new Set(modules.flatMap(mod => project.projectMemberSurfaces(mod.moduleName)));
    for (const surface of surfaces) {
        for (const member of surface.members) {
            const name = member.name;
            Object.defineProperty(member, 'name', { get() { reads++; return name; } });
            Object.freeze(member);
        }
        Object.freeze(surface.members);
        Object.freeze(surface);
    }
    const at = modules[0].source.indexOf(name + '()');
    const result = collectSymbolReferences(byModule, project, modules, modules[0].source,
        'Library', modules[0], name, at + name.length, at, true);
    expect(result).toEqual({ references: [
        { moduleName: 'Library', line: position === 'first' ? 0 : 1998, column: 11, length: name.length, kind: 'write' },
        ...Array.from({ length: count }, (_, i) => ({ moduleName: 'Caller', line: i + 1, column: 12, length: name.length, kind: 'read' })),
    ], hasSymbol: true, ambiguous: [] });
    expect(reads).toBeLessThan(30 * count);
});

function hostFixture() {
    let reads = 0;
    const definition = (name: string) => ({ moduleName: name, nameSpan: { start: 0, end: 1 }, fullSpan: { start: 0, end: 2 } });
    const members: Array<HostMember & { definitions: ReturnType<typeof definition>[] }> = Array.from({ length: 1000 }, (_, i) => Object.freeze({
        get name() { reads++; return 'Member' + i; }, kind: 'method' as const, definitions: [definition('Home' + i)],
    }));
    Object.freeze(members);
    const model: HostObjectModel = { source: 'Frozen lookup work model', types: { 'Excel.Range': { displayName: 'Range', members } }, aliases: { range: 'Excel.Range' }, globals: {} };
    const source = 'Sub UseIt()\nDim r As Range\nr.Member0\nEnd Sub\n';
    const offset = source.indexOf('r.Member0') + 'r.Member0'.length;
    resolveMemberDefinitionsAt(source, offset, 'Member0', { model });
    reads = 0;
    return { model, source, offset, reads: () => reads, definition };
}

it('preserves the cheap first-hit uncached definition API', () => {
    const f = hostFixture();
    expect(resolveMemberDefinitionsAt(f.source, f.offset, 'mEmBeR0', { model: f.model })).toEqual([f.definition('Home0')]);
    expect(f.reads()).toBe(1);
});

it('indexes a supplied cached surface once across repeated definition queries', () => {
    const f = hostFixture(), ctx = { model: f.model, memberSurfaceCache: new Map() };
    const source = f.source.replace('r.Member0', 'r.Member999');
    const offset = source.indexOf('r.Member999') + 'r.Member999'.length;
    for (let i = 0; i < 1000; i++) {
        expect(resolveMemberDefinitionsAt(source, offset, 'Member999', ctx)).toEqual([f.definition('Home999')]);
    }
    expect(f.reads()).toBe(1000);
});

it('preserves first case-insensitive definitions on duplicate members', () => {
    const f = hostFixture();
    const first = { name: 'Member0', kind: 'method' as const, definitions: [f.definition('First')] };
    const second = { name: 'MEMBER0', kind: 'method' as const, definitions: [f.definition('Second')] };
    const model = { ...f.model, types: { 'Excel.Range': { displayName: 'Range', members: [first, second] } } };
    expect(resolveMemberDefinitionsAt(f.source, f.offset, 'mEmBeR0', { model, memberSurfaceCache: new Map() })).toEqual(first.definitions);
});

it('keeps changed definition metadata fresh between contexts', () => {
    const f = hostFixture();
    const run = (home: string) => {
        const model = { ...f.model, types: { 'Excel.Range': { displayName: 'Range', members: [{ name: 'Member0', kind: 'method' as const, definitions: [f.definition(home)] }] } } };
        return resolveMemberDefinitionsAt(f.source, f.offset, 'Member0', { model, memberSurfaceCache: new Map() });
    };
    expect(run('First')).toEqual([f.definition('First')]);
    expect(run('Second')).toEqual([f.definition('Second')]);
});
