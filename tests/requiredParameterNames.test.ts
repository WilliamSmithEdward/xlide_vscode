import { describe, it, expect, vi } from 'vitest';

vi.mock('vscode', async () => (await import('./helpers/vscodeMock')).vscodeMock());

import { parseModule } from '../src/analyzer/parser/parseModule';
import { requiredParameterNames } from '../src/commands/shared';

function firstProcedure(src: string) {
    const proc = parseModule(src).members.find((member) => member.kind === 'Procedure');
    if (!proc || proc.kind !== 'Procedure') {
        throw new Error('no procedure parsed');
    }
    return proc;
}

describe('requiredParameterNames', () => {
    it('returns none for a parameterless Sub', () => {
        expect(requiredParameterNames(firstProcedure('Sub Foo()\nEnd Sub\n'))).toEqual([]);
    });

    it('lists the non-Optional, non-ParamArray parameters', () => {
        const proc = firstProcedure('Sub Foo(a As String, b As Long, Optional c As String)\nEnd Sub\n');
        expect(requiredParameterNames(proc)).toEqual(['a', 'b']);
    });

    it('treats Optional and ParamArray parameters as not required', () => {
        const proc = firstProcedure('Sub Foo(Optional a As String, ParamArray b() As Variant)\nEnd Sub\n');
        expect(requiredParameterNames(proc)).toEqual([]);
    });

    it('handles a line-continued signature (the reported F5 case)', () => {
        const src =
            'Sub mySub2Test(myParam1 As String, _\n' +
            '    myParam2 As String, _\n' +
            '    myParam3 As String, _\n' +
            '    Optional myParam4 As String)\n' +
            '    MsgBox "Hello World!"\n' +
            'End Sub\n';
        expect(requiredParameterNames(firstProcedure(src))).toEqual(['myParam1', 'myParam2', 'myParam3']);
    });
});
