import { afterEach, expect, it, vi } from 'vitest';
import { collectSymbolReferences, projectClassMemberAtDefinition } from '../src/vbaReferenceResolution';
import { buildVbaProjectIndex } from '../src/vbaProjectAnalysis';
import * as lexer from '../src/analyzer/lexer/tokenize';
import * as facts from '../src/analyzer/symbols/classMemberFacts';

afterEach(() => vi.restoreAllMocks());

function fixture(modules: { moduleName: string; source: string; type?: string }[]) {
    const project = buildVbaProjectIndex(modules);
    return { modules, project, byModule: new Map(modules.map((mod) => [mod.moduleName.toLowerCase(), mod])) };
}

it('bounds local lexer and stripping input on fresh large-module revisions', () => {
    const padding = Array.from({ length: 1300 }, (_, index) =>
        `Public Function Pad${index}() As Object\nSet Pad${index} = Nothing\nEnd Function\n`).join('');
    for (let revision = 0; revision < 6; revision++) {
        const source = padding + '\nPublic Sub LocalProbe()\nDim Target As Long\nTarget = Target + 1 '
            + `' revision ${revision}\nEnd Sub\n`;
        const { project, modules, byModule } = fixture([{ moduleName: 'Home', type: 'class', source }]);
        const members = vi.spyOn(project, 'projectMemberSurfaces');
        const values = vi.spyOn(facts, 'classMemberValues');
        const originalTokenize = lexer.tokenizeCached;
        const lexerInputs: number[] = [];
        const tokens = vi.spyOn(lexer, 'tokenizeCached').mockImplementation((text) => {
            lexerInputs.push(text.length);
            return originalTokenize(text);
        });
        const originalSplit = String.prototype.split;
        const splitInputs: number[] = [];
        String.prototype.split = function (...args: Parameters<typeof originalSplit>) {
            splitInputs.push(String(this).length);
            return Reflect.apply(originalSplit, this, args);
        };
        const at = source.lastIndexOf('Target +');
        let result;
        try {
            result = collectSymbolReferences(byModule, project, modules, source, 'Home', modules[0], 'Target', at + 6, at, true);
            expect(collectSymbolReferences(byModule, project, modules, source, 'Home', modules[0], 'Target', at + 6, at, true))
                .toEqual(result);
        } finally {
            String.prototype.split = originalSplit;
            tokens.mockRestore();
        }
        expect(result?.references.map((row) => row.kind)).toEqual(['write', 'write', 'read']);
        expect(result?.references.map((row) => row.line)).toEqual([3902, 3903, 3903]);
        expect(members.mock.calls.length).toBe(0);
        expect(values.mock.calls.length).toBe(0);
        expect(lexerInputs.length).toBeGreaterThan(0);
        expect(Math.max(...lexerInputs)).toBeLessThan(300);
        expect(Math.max(...splitInputs)).toBeLessThan(300);
        // Both requests strip only the small procedure. The scoped substring
        // never becomes a key in the long-lived module stripping cache.
        expect(splitInputs.filter((length) => length === Math.max(...splitInputs))).toHaveLength(2);
        members.mockRestore();
        values.mockRestore();
    }
});

it('omits diagnostic value scans for unmatched qualified members through the whole collector', () => {
    const source = 'Public Value As Object\n' + Array.from({ length: 1300 }, (_, index) =>
        `Public Function MissPad${index}() As Object\nSet MissPad${index} = Nothing\nEnd Function\n`).join('')
        + '\nSub P()\nThisWorkbook.Sheets(1).cez\nEnd Sub\n';
    const { project, modules, byModule } = fixture([{ moduleName: 'Home', type: 'class', source }]);
    const values = vi.spyOn(facts, 'classMemberValues');
    const at = source.lastIndexOf('.cez') + 1;
    expect(collectSymbolReferences(byModule, project, modules, source, 'Home', modules[0], 'cez', at + 3, at, true))
        .toEqual({ references: [], hasSymbol: false, ambiguous: [] });
    expect(values.mock.calls.length).toBe(0);
    // The exported declaration lookup keeps its existing diagnostic metadata
    // unless a reference-only caller explicitly opts out of those facts.
    const member = projectClassMemberAtDefinition(project, 'Home', 'Value', source.indexOf('Value'));
    expect(member?.knownValue).toBe('nothing');
    expect(values).toHaveBeenCalledTimes(1);
});

it('retains function-result declaration and qualified callable references', () => {
    const source = 'Public Function Result() As Object\n Set Result = Nothing\nEnd Function\n';
    const caller = 'Sub P()\n    Dim home As Home\n    Set value = home.Result\nEnd Sub\n';
    const { project, modules, byModule } = fixture([
        { moduleName: 'Home', source, type: 'class' }, { moduleName: 'Caller', source: caller },
    ]);
    const at = source.indexOf('Set Result') + 4;
    expect(collectSymbolReferences(byModule, project, modules, source, 'Home', modules[0], 'Result', at + 6, at, true))
        .toEqual({
            references: [
                { moduleName: 'Home', line: 0, column: 16, length: 6, kind: 'write' },
                { moduleName: 'Caller', line: 2, column: 21, length: 6, kind: 'read' },
                { moduleName: 'Home', line: 1, column: 5, length: 6, kind: 'write' },
            ], hasSymbol: true, ambiguous: [],
        });
});

it('preserves first-line columns and mixed line endings for an indented parameter', () => {
    const word = 'Δείγμα';
    const source = 'Option Explicit\r\n   Public Sub P(ByVal Δείγμα As Long)\n    Δείγμα = Δείγμα + 1\rEnd Sub\n';
    const { project, modules, byModule } = fixture([{ moduleName: 'Home', source, type: 'class' }]);
    const at = source.indexOf(word);
    expect(collectSymbolReferences(byModule, project, modules, source, 'Home', modules[0], word, at + word.length, at, true))
        .toEqual({ references: [
            { moduleName: 'Home', line: 1, column: 22, length: word.length, kind: 'write' },
            { moduleName: 'Home', line: 2, column: 4, length: word.length, kind: 'write' },
            { moduleName: 'Home', line: 2, column: 13, length: word.length, kind: 'read' },
        ], hasSymbol: true, ambiguous: [] });
});
