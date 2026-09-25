import * as path from 'path';
import { describe, expect, it } from 'vitest';
import { listWorkbookSheets } from '../src/vba/projectService';

// SheetsFixture.xlsm, .xlsb and .xls were saved by Excel 16.0 from one
// workbook: Budget, whose module is Sheet1; Drawn, a sheet added through
// automation, with a shape and no module; Trend, a chart sheet; Later, with
// nothing; and Hidden, hidden. Excel reported the code names as it saved:
// only Budget has one.
const FIXTURES = path.join(__dirname, 'fixtures', 'binaries');

const SHEETS = [
    { name: 'Budget', codeName: 'Sheet1', kind: 'worksheet' },
    { name: 'Drawn', kind: 'worksheet' },
    { name: 'Trend', kind: 'chartsheet' },
    { name: 'Later', kind: 'worksheet' },
    { name: 'Hidden', kind: 'worksheet', state: 'hidden' },
];

describe('the sheets of a workbook', () => {
    it.each(['xlsm', 'xlsb', 'xls'])('read the same from a .%s: name, code name, kind and state, in tab order', (extension) => {
        expect(listWorkbookSheets(path.join(FIXTURES, `SheetsFixture.${extension}`)).sheets).toEqual(SHEETS);
    });

    it.each(['xltm', 'xlt', 'xla'])('read from a template or add-in (.%s)', (extension) => {
        expect(listWorkbookSheets(path.join(FIXTURES, `ExcelFixture.${extension}`)).sheets)
            .toEqual([{ name: 'Sheet1', codeName: 'Sheet1', kind: 'worksheet' }]);
    });

    it('read from a workbook with no VBA project, where no sheet has a code name, chart sheet included', () => {
        for (const extension of ['xlsm', 'xlsb']) {
            expect(listWorkbookSheets(path.join(FIXTURES, `NoVbaSheetsFixture.${extension}`)).sheets, extension).toEqual([
                { name: 'Trend', kind: 'chartsheet' },
                { name: 'Report', kind: 'worksheet' },
                { name: 'Data', kind: 'worksheet' },
                { name: 'Sheet4', kind: 'worksheet' },
            ]);
        }
        const legacy = listWorkbookSheets(path.join(FIXTURES, 'NoVbaFixture.xls')).sheets;
        expect(legacy.length).toBeGreaterThan(0);
        expect(legacy.map((sheet) => sheet.codeName)).toEqual(legacy.map(() => undefined));
    });

    it('give a sheet with shapes and no module no code name, as the shape listing does', () => {
        const sheets = listWorkbookSheets(path.join(FIXTURES, 'ShapesFixture.xlsm')).sheets;

        expect(sheets).toEqual([{ name: 'Sheet1', codeName: 'Sheet1', kind: 'worksheet' }, { name: 'Sheet2', kind: 'worksheet' }]);
    });

    it('are refused for a file that is not a workbook', () => {
        expect(() => listWorkbookSheets(path.join(FIXTURES, 'WordFixture.docm'))).toThrow(/has no sheets/);
        expect(() => listWorkbookSheets(path.join(FIXTURES, 'PowerPointFixture.pptm'))).toThrow(/has no sheets/);
    });
});
