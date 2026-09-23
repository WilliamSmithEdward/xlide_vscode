import { describe, expect, it } from 'vitest';
import {
    capturePlaceLines,
    parsePlace,
    placeAssignmentLine,
    PLACE_REPORT_LINE,
    restorePlaceLines,
} from '../src/officeViewPlace';

describe('a place in a copy of the file', () => {
    it('reads what each application calls a place', () => {
        const excel = capturePlaceLines('excel').join('\n');
        expect(excel).toContain('$place.sheet = [string]$window.ActiveSheet.Name');
        expect(excel).toContain('$place.scrollColumn = [int]$window.ScrollColumn');
        const word = capturePlaceLines('word').join('\n');
        expect(word).toContain('$place.scrolled = [int]$window.VerticalPercentScrolled');
        const deck = capturePlaceLines('powerpoint').join('\n');
        expect(deck).toContain('$place.slide = [int]$file.Windows.Item(1).View.Slide.SlideIndex');
        // Which of the application's other files was in front, by the
        // property each one names it with.
        expect(excel).toContain('$app.ActiveWorkbook');
        expect(word).toContain('$app.ActiveDocument');
        expect(deck).toContain('$app.ActivePresentation');
    });

    it('has none in Access, which reopens a database as it was', () => {
        expect(capturePlaceLines('access')).toEqual([]);
        expect(restorePlaceLines('access')).toEqual([]);
    });

    it('survives the trip from the close script to the reopen', () => {
        const place = parsePlace([
            'noise',
            'XLIDE_PLACE|{"sheet":"Q3 \'Actuals\'","selection":"$B$2:$C$4","activeCell":"$C$3","scrollRow":40,"scrollColumn":2,"otherActive":"C:\\\\w\\\\Other.xlsx"}',
        ]);
        expect(place).toEqual({
            sheet: "Q3 'Actuals'", selection: '$B$2:$C$4', activeCell: '$C$3', scrollRow: 40, scrollColumn: 2,
            otherActive: 'C:\\w\\Other.xlsx',
        });
        // Quoted for PowerShell: a sheet name may hold an apostrophe.
        expect(placeAssignmentLine(place!)).toBe(
            "$place = @{ scrollRow = 40; scrollColumn = 2; sheet = 'Q3 ''Actuals'''; selection = '$B$2:$C$4'; activeCell = '$C$3'; otherActive = 'C:\\w\\Other.xlsx' }",
        );
    });

    it('keeps a zero, which is a place too, and drops what is not one', () => {
        expect(parsePlace(['XLIDE_PLACE|{"start":0,"end":0,"scrolled":0}'])).toEqual({ start: 0, end: 0, scrolled: 0 });
        expect(parsePlace(['XLIDE_PLACE|{"slide":"3","sheet":42,"bogus":1}'])).toBeUndefined();
        expect(parsePlace(['XLIDE_PLACE|{}'])).toBeUndefined();
        expect(parsePlace(['XLIDE_PLACE|nope'])).toBeUndefined();
        expect(parsePlace([])).toBeUndefined();
    });

    it('is printed only when there is one', () => {
        expect(PLACE_REPORT_LINE).toBe('if ($place) { [Console]::Out.WriteLine("XLIDE_PLACE|" + (ConvertTo-Json -InputObject $place -Compress)) }');
    });

    it('is put back with every step on its own, so one that fails costs only itself', () => {
        for (const host of ['excel', 'word', 'powerpoint'] as const) {
            const restore = restorePlaceLines(host);
            expect(restore[0]).toBe('if ($place) {');
            expect(restore.join('\n')).toContain('$place.ContainsKey(');
            expect(restore.filter((line) => line.trim().startsWith('try')).length, host).toBeGreaterThan(0);
        }
    });
});
