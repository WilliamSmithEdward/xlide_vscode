import { describe, expect, it } from 'vitest';
import { analyzeVbaModuleSource } from '../../../src/vbaModuleAnalysis';
import { referencedProjectNames, referencedProjectSurface, type ReferencedProjectModule, type ReferencedProjectSurface } from '../../../src/analyzer/symbols/referencedProjects';
import { oracleProjectReferences } from '../../helpers/oracleModuleInput';

const provider = () => referencedProjectSurface('ProbeLibrary', [
    { name: 'Service', kind: 'standard', source: 'Public Function Value() As Long\r\nEnd Function\r\nPublic Type Record\r\n    Data As Long\r\nEnd Type\r\n' },
    { name: 'Secret', kind: 'standard', source: 'Option Private Module\r\nPublic Function Hidden() As Long\r\nEnd Function\r\n' },
    { name: 'Widget', kind: 'class', source: 'Attribute VB_Exposed = True\r\nPublic Value As Long\r\n' },
    { name: 'Thing', kind: 'class', source: 'Public Value As Long\r\n' },
]);

function errors(host: string, code: string, surface: ReferencedProjectSurface) {
    return analyzeVbaModuleSource({
        source: `Option Explicit\r\nPublic Sub Main()\r\n    ${code}\r\nEnd Sub\r\n`,
        moduleName: 'Module1', moduleKind: 'standard', host,
        ...oracleProjectReferences({ host }),
        referencedLibraries: [...oracleProjectReferences({ host }).referencedLibraries, 'ProbeLibrary'],
        referencedProjects: ['ProbeLibrary'], referencedProjectSurfaces: [surface],
    }).diagnostics.filter(diagnostic => diagnostic.severity === 'error');
}

describe.each(['excel', 'word', 'powerpoint', 'access'])('indexed referenced names in %s', host => {
    it('preserves qualified, bare and case-insensitive standard-module lookups', () => {
        expect(errors(host, 'Dim w As pRoBeLiBrArY.wIdGeT\r\n    Dim x As Long\r\n    x = ProbeLibrary.Value() + Service.Value() + Value()', provider())).toEqual([]);
        expect(errors(host, 'Dim x As Long\r\n    x = ProbeLibrary.Secret.Hidden()', provider()).map(d => d.message)).toEqual(["Method or data member not found: 'ProbeLibrary.Secret'."]);
        expect(errors(host, 'Dim x As Long\r\n    x = ProbeLibrary.Service.Nope()', provider()).map(d => d.message)).toEqual(["Method or data member not found: 'ProbeLibrary.Service.Nope'."]);
        expect(errors(host, 'Dim w As ProbeLibrary.Thing', provider())).toHaveLength(1);
    });

    it('bounds module-list reads across many qualified types and project calls', () => {
        const modules: ReferencedProjectModule[] = Array.from({ length: 150 }, (_, i) => ({
            name: `Part${i}`, kind: 'class', visible: true, types: [`Part${i}`], members: [],
        }));
        modules.push({ name: 'Service', kind: 'standard', visible: true, types: [], members: ['Value'] });
        let reads = 0;
        const surface: ReferencedProjectSurface = { name: 'ProbeLibrary', modules: new Proxy(modules, {
            get(target, key, receiver) {
                if (typeof key === 'string' && /^\d+$/.test(key)) reads++;
                return Reflect.get(target, key, receiver);
            },
        }) };
        const code = ['Dim x As Long', ...modules.slice(0, -1).map((module, i) =>
            `Dim w${i} As ProbeLibrary.${module.name}\r\n    x = ProbeLibrary.Value() + ProbeLibrary.Service.Value()`)].join('\r\n    ');
        expect(errors(host, code, surface)).toEqual([]);
        // One surface index replaces per-type and per-call module scans.
        // The bound allows fixed passes but fails quadratic lookup work.
        expect(reads).toBeLessThan(modules.length * 6);
    });
});

describe('referenced-project snapshot lifetime', () => {
    it('shares one index for a snapshot and rebuilds for its replacement', () => {
        const first = provider();
        expect(referencedProjectNames(first)).toBe(referencedProjectNames(first));
        const replacement = referencedProjectSurface('ProbeLibrary', [{ name: 'Updated', kind: 'class', source: 'Attribute VB_Exposed = True\r\n' }]);
        expect(referencedProjectNames(replacement).types.has('updated')).toBe(true);
        expect(referencedProjectNames(replacement).types.has('widget')).toBe(false);
        expect(referencedProjectNames(first).types.has('widget')).toBe(true);
    });
});
