// Tripwire: no identifier XLIDE declares in an injected VBA module may differ
// in CASING from a name the host object models or the VBA runtime already
// canonicalize (issue #38).
//
// VBA cases identifiers project-wide to the latest declaration it sees, so a
// module declaring `ByVal value` re-spells every `.Value` in the user's whole
// project to `.value` the moment it is injected - measured live, turning a
// fixture's `Err.Number` into `Err.number`. XlideAssert is installed
// permanently, so its contagion is permanent and user-visible.
//
// The rule is mechanical rather than a hand-kept list: any declared name that
// matches a canonical name case-insensitively must match it exactly.
import { describe, expect, it } from 'vitest';
import {
    buildModuleSymbols,
    getHostConstants,
    getHostGlobals,
    getHostMembers,
    VBA_RUNTIME_CONSTANTS,
    VBA_RUNTIME_FUNCTIONS,
    VBA_RUNTIME_OBJECTS,
    type VbaSymbol,
} from '../src/analyzer';
import { getExcelObjectModel, type HostObjectModel } from '../src/analyzer/host/excelObjectModel';
import { getWordObjectModel } from '../src/analyzer/host/wordObjectModel';
import { getPowerPointObjectModel } from '../src/analyzer/host/powerpointObjectModel';
import { getAccessObjectModel } from '../src/analyzer/host/accessObjectModel';
import { XLIDE_ASSERT_MODULE_SOURCE } from '../src/vbaTestSupportModule';
import {
    buildVbaTestDirectRunnerModule,
    buildVbaTestDispatchModule,
} from '../src/vbaTestRunnerModuleCodegen';
import type { VbaTestCase } from '../src/vbaTestRunner';

/** Lowercased name -> every canonical spelling the models and runtime carry. */
function canonicalSpellings(): Map<string, Set<string>> {
    const byLower = new Map<string, Set<string>>();
    const add = (name: string): void => {
        const key = name.toLowerCase();
        const spellings = byLower.get(key) ?? new Set<string>();
        spellings.add(name);
        byLower.set(key, spellings);
    };
    const models: HostObjectModel[] = [
        getExcelObjectModel(),
        getWordObjectModel(),
        getPowerPointObjectModel(),
        getAccessObjectModel(),
    ];
    for (const model of models) {
        for (const qualified of Object.keys(model.types)) {
            for (const member of getHostMembers(qualified, model)) {
                add(member.name);
            }
        }
        for (const global of getHostGlobals(model)) {
            add(global.name);
        }
        for (const constant of getHostConstants(model)) {
            add(constant.name);
        }
    }
    for (const object of VBA_RUNTIME_OBJECTS) {
        add(object.name);
        for (const member of object.members ?? []) {
            add(member.name);
        }
    }
    for (const fn of VBA_RUNTIME_FUNCTIONS) {
        add(fn.name);
    }
    for (const constant of VBA_RUNTIME_CONSTANTS) {
        add(constant.name);
    }
    return byLower;
}

/** Every name the module's own text declares: procedures, params, locals, consts. */
function declaredNames(source: string): string[] {
    const names: string[] = [];
    const walk = (symbols: readonly VbaSymbol[] | undefined): void => {
        for (const symbol of symbols ?? []) {
            names.push(symbol.name);
            walk(symbol.children);
        }
    };
    walk(buildModuleSymbols('Probe', 'standard', source).root.children);
    return names;
}

const SAMPLE_TESTS: VbaTestCase[] = [{
    id: 'TestModule.Scenario',
    moduleName: 'TestModule',
    moduleType: 'standard',
    procedureName: 'Scenario',
    qualifiedName: 'TestModule.Scenario',
    line: 1,
    column: 0,
    annotationLine: 0,
    metadata: { tags: [] },
} as VbaTestCase];

const INJECTED_MODULES: ReadonlyArray<[string, string]> = [
    ['XlideAssert', XLIDE_ASSERT_MODULE_SOURCE],
    ['XlideTestRuntime', buildVbaTestDirectRunnerModule(SAMPLE_TESTS)],
    ['XlideTestDispatch', buildVbaTestDispatchModule([
        { name: 'TestModule', type: 'standard', source: 'Public Sub Scenario()\r\nEnd Sub\r\n' },
    ])],
];

describe.each(INJECTED_MODULES)('%s declares no re-casing identifier', (_label, source) => {
    it('spells every shadowing name in the canonical casing', () => {
        const canonical = canonicalSpellings();
        const violations: string[] = [];
        for (const name of new Set(declaredNames(source))) {
            const spellings = canonical.get(name.toLowerCase());
            if (spellings && !spellings.has(name)) {
                violations.push(`${name} (canonical: ${[...spellings].sort().join(' | ')})`);
            }
        }
        expect(violations).toEqual([]);
    });
});

describe('the casing rule itself', () => {
    it('catches a lowercase declaration of a canonical host member name', () => {
        // The rule has teeth: the pre-fix spelling is exactly what it rejects.
        const canonical = canonicalSpellings();
        expect(canonical.get('value')?.has('value')).toBe(false);
        expect(canonical.get('value')?.has('Value')).toBe(true);
        expect(canonical.get('number')?.has('Number')).toBe(true);
        expect(canonical.get('source')?.has('Source')).toBe(true);
        expect(canonical.get('message')?.has('Message')).toBe(true);
    });
});
