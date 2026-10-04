import { describe, expect, it } from 'vitest';
import { parseModule } from '../src/analyzer/parser/parseModule';
let eviction = 0;
function cold(source: string) {
    for (let i = 0; i < 12; i++) { parseModule("' cache eviction " + eviction++); }
    return parseModule(source);
}
const prefix = Array.from({ length: 140 }, (_, i) => 'Public Sub Pad' + i + '()\r\nDebug.Print 1\r\nEnd Sub\r\n').join('');
const tail = [
    'Private value As Object',
    'Private Sub Probe(Optional n As Long = 1)',
    ' Dim text As String: text = "a ""quoted"" string"',
    ' With ThisWorkbook.Sheets(1)',
    '  .Name = text',
    ' End With',
    ' If n Then',
    '  result = n + _',
    '   1',
    ' End If',
    'End Sub',
    '   ',
].join('\r\n');

describe('unchanged parser prefix reuse', () => {
    it('matches a cold full parse after insertions and deletions at every fixture offset', () => {
        const original = prefix + tail;
        for (let offset = prefix.length; offset <= original.length; offset++) {
            for (const inserted of ['x', "'", '"', ':', '_', '\r', '\n', 'ก้']) {
                const before = cold(original);
                const snapshot = structuredClone(before);
                const changed = original.slice(0, offset) + inserted + original.slice(offset);
                const result = parseModule(changed);
                expect(result, 'insert ' + JSON.stringify(inserted) + ' at ' + offset).toEqual(cold(changed));
                expect(before).toEqual(snapshot);
            }
            if (offset < original.length) {
                cold(original);
                const changed = original.slice(0, offset) + original.slice(offset + 1);
                expect(parseModule(changed), 'delete at ' + offset).toEqual(cold(changed));
            }
        }
    }, 60000);

    it('shares unchanged members but reparses the edited procedure with new absolute spans', () => {
        const original = prefix + tail;
        const before = cold(original);
        const snapshot = structuredClone(before);
        const changed = original.replace('.Name = text', '.Name = text & "more"');
        const result = parseModule(changed);
        expect(result.members.filter((member, i) => member === before.members[i]).length).toBeGreaterThan(130);
        expect(result.members[result.members.length - 1]).not.toBe(before.members[before.members.length - 1]);
        expect(result).toEqual(cold(changed));
        expect(before).toEqual(snapshot);
    });

    it('preserves module kind, declarations and diagnostics when a tail gains an error', () => {
        const original = 'Attribute VB_PredeclaredId = True\r\n' + prefix + tail;
        const before = cold(original);
        const changed = original.replace('End With', 'End If');
        const result = parseModule(changed);
        expect(result.moduleKind).toBe('class');
        expect(result.diagnostics.length).toBeGreaterThan(0);
        expect(result.members[0]).toBe(before.members[0]);
        expect(result).toEqual(cold(changed));
        const next = changed.replace('result = n', 'result = x');
        parseModule(changed);
        expect(parseModule(next)).toEqual(cold(next));
    });

    it('retains a full parse for conditional-directive state and large replacements', () => {
        for (const original of [
            prefix + '#If VBA7 Then\r\n' + tail + '\r\n#End If\r\n',
            prefix + tail.replace(' Dim text', '#If VBA7 Then\r\n#End If\r\n Dim text'),
        ]) {
            const before = cold(original);
            const changed = original.replace('.Name = text', '.Name = n');
            const result = parseModule(changed);
            expect(result.members[0]).not.toBe(before.members[0]);
            expect(result).toEqual(cold(changed));
        }
        const original = prefix + tail;
        const before = cold(original);
        const changed = original + "' " + 'x'.repeat(200);
        const result = parseModule(changed);
        expect(result.members[0]).not.toBe(before.members[0]);
        expect(result).toEqual(cold(changed));
    });

    it('does not reuse a closing word that is extended, a partial colon line, or a CRLF split', () => {
        for (const [original, changed] of [
            [prefix, prefix.trimEnd() + 'Extra\r\n'],
            [prefix + 'Sub A():End Sub:Sub B():End Sub\r\n', prefix + 'Sub A():End Sub:Sub B():End SubExtra\r\n'],
            [prefix + 'Sub A()\r\nEnd Sub\r\n', prefix + 'Sub A()\rx\nEnd Sub\r\n'],
            [prefix + 'If n Then a: Dim b As Long\r\n', prefix + 'If n Then a: Dim b As Object\r\n'],
        ]) {
            cold(original);
            expect(parseModule(changed)).toEqual(cold(changed));
        }
    });

    it('uses safe earlier boundaries for long colon lines and prior diagnostics', () => {
        const line = Array.from({ length: 1000 }, (_, i) => 'Dim x' + i + ' As Long').join(':') + '\r\n';
        const original = prefix + line;
        const before = cold(original);
        const changed = original.replace('Dim x999 As Long', 'Dim x999 As Object');
        const result = parseModule(changed);
        expect(result.members[0]).toBe(before.members[0]);
        expect(result).toEqual(cold(changed));
        const broken = 'Sub MissingEnd()\r\n' + prefix + tail;
        cold(broken);
        expect(parseModule(broken.replace('.Name = text', '.Name = n'))).toEqual(cold(broken.replace('.Name = text', '.Name = n')));
    });
});
