import { describe, expect, it } from 'vitest';
import { buildLiveVbaProjectIndex, projectEditorSymbolContextForModule } from '../src/vbaProjectAnalysis';

const caller = Array.from({ length: 1200 }, (_, i) =>
    'Public Sub Local' + i + '(Optional value As Long = 1)\nEnd Sub\n').join('') +
    'Private Const PrivateValue As Long = 1\nPublic Const SharedValue As Long = 2\n';

const modules = () => [
    { moduleName: 'Caller', moduleKind: 'standard' as const, source: caller },
    { moduleName: 'Api', moduleKind: 'standard' as const, source: 'Public Sub Work()\nEnd Sub\nPrivate Sub Hidden()\nEnd Sub\nPublic Const ApiValue As Long = 3\n' },
    { moduleName: 'Person', moduleKind: 'class' as const, source: 'Public Sub Save()\nEnd Sub\nPublic Name As String\n' },
];

describe('external editor symbol queries', () => {
    it('skips current-module declarations before scanning or formatting them', () => {
        const project = buildLiveVbaProjectIndex(modules());
        // Local types and qualified Caller.member surfaces remain necessary.
        // Warm these before counting work specific to the bare external lists.
        project.visibleTypeNames('Caller');
        project.projectMemberSurfaces('Caller');
        const root = project.getModule('Caller')!.root;
        let reads = 0;
        root.children = new Proxy(root.children!, {
            get(target, property, receiver) {
                if (typeof property === 'string' && /^\d+$/.test(property)) { reads++; }
                return Reflect.get(target, property, receiver);
            },
        });
        const context = projectEditorSymbolContextForModule(project, 'cAlLeR');
        expect(context.externalProjectProcedures.map(item => item.name)).toEqual(['Work']);
        expect(context.externalProjectSymbols.map(item => item.name)).toEqual(['Work', 'ApiValue']);
        expect(reads).toBe(0);
    });

    it('matches full visible queries filtered to external modules in either cache order', () => {
        for (const externalFirst of [false, true]) {
            const project = buildLiveVbaProjectIndex(modules());
            const full = () => ({
                procedures: project.visibleProcedureSignatures('Caller').filter(item => item.moduleName.toLowerCase() !== 'caller'),
                symbols: project.visibleIdentifierSymbols('Caller').filter(item => item.moduleName.toLowerCase() !== 'caller'),
            });
            const before = externalFirst ? undefined : full();
            const context = projectEditorSymbolContextForModule(project, 'Caller');
            const expected = before ?? full();
            expect(context.externalProjectProcedures).toEqual(expected.procedures);
            expect(context.externalProjectSymbols).toEqual(expected.symbols);
            expect(project.visibleProcedureSignatures('Caller').some(item => item.name === 'Local0')).toBe(true);
            expect(project.visibleIdentifierSymbols('Caller').some(item => item.name === 'PrivateValue')).toBe(true);
            expect(context.analysisOptions.projectClassMembers?.find(item => item.name === 'Caller')?.members
                .some(item => item.name === 'Local0')).toBe(true);
        }
    });

    it('refreshes visibility after edits and module removal without mixing own/external caches', () => {
        const project = buildLiveVbaProjectIndex(modules());
        projectEditorSymbolContextForModule(project, 'Caller');
        project.setModule({ moduleName: 'Api', moduleKind: 'standard', source: 'Private Sub Work()\nEnd Sub\nPublic Sub Fresh()\nEnd Sub' });
        expect(projectEditorSymbolContextForModule(project, 'Caller').externalProjectProcedures.map(item => item.name)).toEqual(['Fresh']);
        expect(project.visibleProcedureSignatures('Api').filter(item => item.moduleName === 'Api').map(item => item.name)).toEqual(['Work', 'Fresh']);
        project.removeModule('Api');
        expect(projectEditorSymbolContextForModule(project, 'Caller').externalProjectProcedures).toEqual([]);
        const only = buildLiveVbaProjectIndex([modules()[0]]);
        expect(projectEditorSymbolContextForModule(only, 'Caller').externalProjectSymbols).toEqual([]);
    });
});

it.skipIf(process.env.XLIDE_EXTERNAL_EDITOR_BENCH !== '1')('measures fresh editor projections without formatting discarded local signatures', () => {
    const times = Array.from({ length: 31 }, () => {
        const project = buildLiveVbaProjectIndex(modules());
        const before = performance.now();
        projectEditorSymbolContextForModule(project, 'Caller');
        return performance.now() - before;
    }).sort((a, b) => a-b);
    process.stdout.write('Fresh editor context projection benchmark: ' + JSON.stringify({
        procedures: 1201, bytes: caller.length, medianMs: times[15], slowestMs: times[30],
    }) + '\n');
});
