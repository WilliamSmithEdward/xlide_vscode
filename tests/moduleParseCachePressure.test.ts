import { describe, expect, it } from 'vitest';
import { parseModule, parseModuleFreshForTests } from '../src/analyzer/parser/parseModule';
import { buildModuleSymbols } from '../src/analyzer/symbols/buildModuleSymbols';
import { moduleHasConditionalDirectives } from '../src/analyzer/conditional/conditionalCompilation';

function fixture(name: string, eol = '\n') {
    return [
        `Attribute VB_Name = "${name}"`, 'Option Explicit',
        'Public Type Point', '    Value As Long', 'End Type',
        'Public Sub Before()', '    Dim stable As Long', 'End Sub', '',
    ].join(eol) + eol + ("' Unchanged parser pressure prefix" + eol).repeat(700) +
        ['Public Sub Tail()', '    Dim target As Long', '    target = 1', 'End Sub', ''].join(eol);
}

function pressure(tag: string) {
    for (let index = 0; index < 20; index++) {
        parseModule(`Sub ${tag}${index}()\nDim value As Long\nvalue = value + ${index}\nEnd Sub\n`);
    }
}

describe('module parser snapshots under short lookup pressure', () => {
    it.each(['\n', '\r\n', '\r'])('retains the exact large snapshot after short lookups (%j)', eol => {
        const source = fixture('Exact' + eol.length + eol.charCodeAt(0), eol);
        const before = parseModule(source);
        pressure('ExactPressure');
        expect(parseModule(source)).toBe(before);
        expect(before).toStrictEqual(parseModuleFreshForTests(source));
    });

    it.each(['\n', '\r\n', '\r'])('reuses unchanged members and preserves full AST metadata after an edit (%j)', eol => {
        let source = fixture('Edited' + eol.length + eol.charCodeAt(0), eol);
        for (let revision = 2; revision <= 6; revision++) {
            const before = parseModule(source);
            const snapshot = structuredClone(before);
            const changed = source.replace(/target = \d+/, 'target = ' + revision);
            pressure('EditPressure' + revision);
            const after = parseModule(changed);
            expect(after).toStrictEqual(parseModuleFreshForTests(changed));
            expect(before).toStrictEqual(snapshot);
            expect(after.members[0]).toBe(before.members[0]);
            expect(after.members.at(-1)).not.toBe(before.members.at(-1));
            source = changed;
        }
    });

    it('keeps branch facts and compiler environments independent after lookup pressure', () => {
        const source = '#If Feature Then\nPublic Sub Enabled()\nEnd Sub\n' +
            '#Else\nPublic Sub Disabled()\nEnd Sub\n#End If\n' + fixture('Branches');
        const before = parseModule(source);
        const snapshot = structuredClone(before);
        pressure('BranchPressure');
        expect(parseModule(source)).toBe(before);
        expect(moduleHasConditionalDirectives(before)).toBe(true);
        const names = (feature: boolean) => buildModuleSymbols('Branches', 'class', source,
            { conditionalCompilation: { projectConstants: { Feature: feature } } }).root.children?.map(symbol => symbol.name);
        expect(names(true)).toContain('Enabled');
        expect(names(true)).not.toContain('Disabled');
        expect(names(false)).toContain('Disabled');
        expect(names(false)).not.toContain('Enabled');
        expect(before).toStrictEqual(snapshot);
    });

    it('keeps eight large records as the upper bound', () => {
        const source = fixture('LargeBound0');
        const before = parseModule(source);
        for (let index = 1; index <= 8; index++) { parseModule(fixture('LargeBound' + index)); }
        expect(parseModule(source)).not.toBe(before);
    });

    it('keeps eight short records as the upper bound', () => {
        const source = 'Sub SmallBound0()\nEnd Sub\n';
        const before = parseModule(source);
        for (let index = 1; index <= 8; index++) { parseModule(`Sub SmallBound${index}()\nEnd Sub\n`); }
        expect(parseModule(source)).not.toBe(before);
    });

    it('retains short snapshots while large records turn over', () => {
        const source = 'Sub SmallIndependent()\nEnd Sub\n';
        const before = parseModule(source);
        for (let index = 0; index < 8; index++) { parseModule(fixture('IndependentLarge' + index)); }
        expect(parseModule(source)).toBe(before);
    });

    it.each([4095, 4096])('uses the intended cache at the %i-character boundary', length => {
        const head = `Sub Boundary${length}()\nEnd Sub\n`;
        const source = head + "'" + 'x'.repeat(length - head.length - 1);
        const before = parseModule(source);
        pressure('BoundaryPressure' + length);
        if (length === 4096) { expect(parseModule(source)).toBe(before); }
        else { expect(parseModule(source)).not.toBe(before); }
        expect(parseModule(source)).toStrictEqual(parseModuleFreshForTests(source));
    });
});
