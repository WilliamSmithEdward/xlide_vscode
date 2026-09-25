import * as fs from 'fs';
import * as path from 'path';
import { describe, expect, it } from 'vitest';
import { ZipArchive } from '../src/vba/zip';
import { sheetsOfOoxml } from '../src/vba/workbookSheets';

// A sheet part far larger than the head that is read for its properties:
// the prefix read decodes only the first bytes of the deflate stream, and
// the sheet list still finds the code name at the top of the part.
const FIXTURES = path.join(__dirname, 'fixtures', 'binaries');

function withLargeBudgetSheet(): ZipArchive {
    const zip = ZipArchive.read(fs.readFileSync(path.join(FIXTURES, 'SheetsFixture.xlsm')));
    const part = 'xl/worksheets/sheet1.xml';
    const xml = zip.read(part).toString('utf8');
    // Rows that do not repeat, so the part does not compress to a few bytes.
    const rows = Array.from({ length: 60000 }, (_, i) =>
        `<row r="${i + 2}"><c r="A${i + 2}"><v>${(i * 7919) % 100003}</v></c><c r="B${i + 2}" t="str"><v>item ${i.toString(36)}</v></c></row>`);
    const large = xml.replace(/<sheetData\b[^>]*\/>|<sheetData>[\s\S]*?<\/sheetData>/, `<sheetData>${rows.join('')}</sheetData>`);
    expect(large.length).toBeGreaterThan(2_000_000);
    zip.write(part, Buffer.from(large, 'utf8'));
    return ZipArchive.read(zip.toBytes());
}

describe('reading the head of a zip entry', () => {
    it('decodes only the start of a large deflated part, which holds what the sheet list needs', () => {
        const zip = withLargeBudgetSheet();
        const whole = zip.read('xl/worksheets/sheet1.xml');
        const head = zip.readPrefix('xl/worksheets/sheet1.xml', 16 * 1024);

        expect(head.length).toBeLessThan(whole.length);
        expect(head.length).toBeGreaterThan(16 * 1024);
        expect(whole.subarray(0, head.length).equals(head)).toBe(true);
        expect(head.toString('utf8')).toContain('<sheetPr codeName="Sheet1"/>');
        expect(sheetsOfOoxml(zip)[0]).toEqual({ name: 'Budget', codeName: 'Sheet1', kind: 'worksheet' });
    });

    it('gives the whole entry when it is no larger than the head', () => {
        const zip = ZipArchive.read(fs.readFileSync(path.join(FIXTURES, 'SheetsFixture.xlsm')));
        expect(zip.readPrefix('xl/workbook.xml', 16 * 1024).equals(zip.read('xl/workbook.xml'))).toBe(true);
    });
});
