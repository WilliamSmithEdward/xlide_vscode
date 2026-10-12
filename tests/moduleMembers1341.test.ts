import { describe, expect, it } from 'vitest';
import { listModuleMembers, listProcedures } from '../src/vba/moduleSource';
import { memberThemeIconName, orderModuleMembers, type ModuleMemberEntry } from '../src/moduleDisplay';

// Issue #1341: the explorer lists a module's Enums and Types beside its
// procedures, in an order the user picks.

const BODY = [
    'Option Explicit',
    'Private Type Point',
    '    X As Long',
    'End Type',
    'Public Enum Colors \' the palette',
    '    Red',
    'End Enum',
    'Property Let Total(ByVal v As Long)',
    'End Property',
    'Sub Run()',
    '    Type = 1',
    'End Sub',
    'Property Get Total() As Long',
    'End Property',
    'Enum Answer',
    '    Yes',
    'End Enum',
    'Function Area() As Long',
    'End Function',
    'Property Set Total(ByVal v As Object)',
    'End Property',
    '\' Type Commented',
].join('\r\n');

const labels = (members: readonly ModuleMemberEntry[]): string[] => members.map((m) => `${m.kind} ${m.name}`);

describe('listModuleMembers', () => {
    it('lists Enums and Types beside the procedures, in the order they are written', () => {
        expect(listModuleMembers(BODY)).toEqual([
            { name: 'Point', kind: 'Type', line: 2 },
            { name: 'Colors', kind: 'Enum', line: 5 },
            { name: 'Total', kind: 'Property Let', line: 8 },
            { name: 'Run', kind: 'Sub', line: 10 },
            { name: 'Total', kind: 'Property Get', line: 13 },
            { name: 'Answer', kind: 'Enum', line: 15 },
            { name: 'Area', kind: 'Function', line: 18 },
            { name: 'Total', kind: 'Property Set', line: 20 },
        ]);
    });

    it('leaves the procedure list alone', () => {
        expect(listProcedures(BODY).map((m) => m.kind)).toEqual(['Property Let', 'Sub', 'Property Get', 'Function', 'Property Set']);
    });
});

describe('orderModuleMembers', () => {
    const members = listModuleMembers(BODY);

    it('keeps the code order, with a property\'s Get, Let and Set together where the first sits', () => {
        expect(labels(orderModuleMembers(members, 'code'))).toEqual([
            'Type Point', 'Enum Colors',
            'Property Get Total', 'Property Let Total', 'Property Set Total',
            'Sub Run', 'Enum Answer', 'Function Area',
        ]);
    });

    it('sorts by name, case aside', () => {
        const mixed = [...members, { name: 'apply', kind: 'Sub', line: 30 }];
        expect(labels(orderModuleMembers(mixed, 'name'))).toEqual([
            'Enum Answer', 'Sub apply', 'Function Area', 'Enum Colors', 'Type Point', 'Sub Run',
            'Property Get Total', 'Property Let Total', 'Property Set Total',
        ]);
    });

    it('sorts by kind: Enums, Types, Subs, Functions, then Properties, each by name', () => {
        expect(labels(orderModuleMembers(members, 'kind'))).toEqual([
            'Enum Answer', 'Enum Colors', 'Type Point', 'Sub Run', 'Function Area',
            'Property Get Total', 'Property Let Total', 'Property Set Total',
        ]);
    });

    it('keeps two rows of one label in the order they are written', () => {
        const twice = [{ name: 'Post', kind: 'Sub', line: 3 }, { name: 'Post', kind: 'Sub', line: 9 }];
        expect(orderModuleMembers(twice, 'name').map((m) => m.line)).toEqual([3, 9]);
    });
});

describe('memberThemeIconName', () => {
    it('gives an Enum and a Type their own icons', () => {
        expect(['Enum Colors', 'Type Point', 'Sub Run', 'Property Get Total'].map(memberThemeIconName))
            .toEqual(['symbol-enum', 'symbol-struct', 'symbol-method', 'symbol-method']);
    });
});
