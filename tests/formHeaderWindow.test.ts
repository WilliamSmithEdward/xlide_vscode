import { describe, it, expect, vi, afterEach } from 'vitest';
import { parseUserFormControls, hasAuthoritativeDesignerHeader } from '../src/vbaUserFormControls';

afterEach(() => vi.restoreAllMocks());
describe('form header allocation work', () => {
    for (const form of [false, true]) {
        it(form ? 'stops allocating after a small designer header' : 'rejects a large code module without allocating all its lines', () => {
            const header = form ? 'VERSION 5.00\nBegin {GUID} ProbeForm\n Begin Forms.ComboBox.1 Picker\n End\nEnd\n' : 'Option Explicit\n';
            const source = header + Array.from({ length: 5000 }, (_, i) => 'Sub HeaderWork' + i + '()\nEnd Sub\n').join('');
            const split = String.prototype.split;
            let wholeSplits = 0;
            let allocatedLines = 0;
            vi.spyOn(String.prototype, 'split').mockImplementation(function (this: string, separator: string | RegExp, limit?: number) {
                const result = split.call(this, separator, limit);
                if (String(this) === source) {
                    wholeSplits++;
                    allocatedLines += result.length;
                }
                return result;
            });
            expect(parseUserFormControls(source).map(control => control.name)).toEqual(form ? ['Picker'] : []);
            expect(hasAuthoritativeDesignerHeader(source)).toBe(form);
            expect(allocatedLines).toBe(0);
            expect(wholeSplits).toBe(0);
        });
    }
});

describe('designer header line boundaries', () => {
    const control = { name: 'Picker', progId: 'Forms.ComboBox.1', type: 'MSForms.ComboBox' };
    const lines = ['VERSION 5.00', 'Begin {GUID} ProbeForm', ' Begin Forms.ComboBox.1 Picker', ' End', 'End'];
    it.each(['\n', '\r\n'])('accepts BOM/blank prefixes and a final line without a newline for %j', ending => {
        const source = '\uFEFF \t' + ending + ending + lines.join(ending);
        expect(parseUserFormControls(source)).toEqual([control]);
        expect(hasAuthoritativeDesignerHeader(source)).toBe(true);
    });
    it('preserves the legacy treatment of a bare CR inside a physical line', () => {
        const source = lines.join('\r');
        expect(parseUserFormControls(source)).toEqual([]);
        expect(hasAuthoritativeDesignerHeader(source)).toBe(false);
    });
    it.each(['', ' \t\r\n', 'VERSION 5.00', 'VERSION_5.00\nEnd', "' comment\n" + lines.join('\n')])('rejects a missing or later opening VERSION line for case %#', source => {
        expect(parseUserFormControls(source)).toEqual([]);
        expect(hasAuthoritativeDesignerHeader(source)).toBe(false);
    });
    it('does not make an OleObjectBlob header authoritative', () => {
        const source = 'VERSION 5.00\nBegin {GUID} ProbeForm\n OleObjectBlob = "probe.frx":0000\nEnd\n';
        expect(parseUserFormControls(source)).toEqual([]);
        expect(hasAuthoritativeDesignerHeader(source)).toBe(false);
    });
    it('stops at code after an unterminated designer rather than finding fake controls later', () => {
        const source = 'VERSION 5.00\nBegin {GUID} ProbeForm\nSub Code()\n Begin Forms.ComboBox.1 Fake\n End\nEnd\n';
        expect(parseUserFormControls(source)).toEqual([]);
        expect(hasAuthoritativeDesignerHeader(source)).toBe(false);
    });
});
