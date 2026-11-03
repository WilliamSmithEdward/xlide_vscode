import { describe, expect, it } from 'vitest';
import { analyzeVbaModuleSource } from '../../../src/vbaModuleAnalysis';
import { referencedProjectSurface } from '../../../src/analyzer/symbols/referencedProjects';
import { oracleProjectReferences } from '../../helpers/oracleModuleInput';

// Isolate type lookup: member diagnostics for referenced objects are separate work.
const provider = referencedProjectSurface('ProbeLibrary', [
    { name: 'Widget', kind: 'class', source: 'Attribute VB_Exposed = True\r\nOption Explicit\r\n' },
    { name: 'Thing', kind: 'class', source: 'Option Explicit\r\n' },
]);

function diagnostics(host: string, type: string, eol: string) {
    const source = `Option Explicit\nPublic Sub Main()\n    Dim w As ${type}\nEnd Sub\n`.replace(/\n/g, eol);
    return analyzeVbaModuleSource({
        source, moduleName: 'Module1', moduleKind: 'standard', host,
        ...oracleProjectReferences({ host }),
        referencedLibraries: [...oracleProjectReferences({ host }).referencedLibraries, 'ProbeLibrary'],
        referencedProjects: ['ProbeLibrary'], referencedProjectSurfaces: [provider],
    }).diagnostics.filter(diagnostic => diagnostic.severity === 'error');
}

describe.each(['excel', 'word', 'powerpoint', 'access'])('continued referenced types in %s', host => {
    it.each(['\n', '\r\n', '\r'])('accepts an exposed type with %j line endings', eol => {
        for (const type of ['ProbeLibrary.Widget', 'ProbeLibrary. _\n        Widget', 'ProbeLibrary. _\u3000\n        Widget', 'ProbeLibrary. _\u0019\n        Widget', '[ProbeLibrary]. _\n        [Widget]', 'ProbeLibrary _\n        . _\n        Widget']) {
            expect(diagnostics(host, type, eol), type).toEqual([]);
        }
    });

    it.each(['Thing', 'Missing'])('still rejects the inaccessible type %s', type => {
        for (const eol of ['\n', '\r\n', '\r']) {
            const errors = diagnostics(host, `ProbeLibrary. _\n        ${type}`, eol);
            expect(errors).toHaveLength(1);
            expect(errors[0].code).toBe('missing-library-reference');
            expect(errors[0].message).toContain(`offers no type '${type}'`);
        }
    });
});
