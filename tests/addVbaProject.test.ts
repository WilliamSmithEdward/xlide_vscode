import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { afterEach, describe, expect, it } from 'vitest';
import { AddVbaProjectError, addVbaProjectToPackage } from '../src/vba/addVbaProject';
import { addVbaProject, hasVbaProject, listModules, readModule, writeModule } from '../src/vba/projectService';
import { Cfb } from '../src/vba/cfb';
import { ZipArchive } from '../src/vba/zip';

// "Add VBA Project" on a file saved in a macro-enabled format before its
// first macro. Every expectation below is what the application itself wrote
// when it added VBA to the same file (measured on Office 16.0, build 20326):
// Excel's code names, module order, module headers and PROJECT lines, and
// each application's content type and relationship for the part.

const FIXTURES = path.join(__dirname, 'fixtures', 'binaries');
const TEMPLATES = path.join(__dirname, '..', 'assets', 'templates');

const dirs: string[] = [];
function scratch(fixture: string, name = fixture): string {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'xlide-add-vba-'));
    dirs.push(dir);
    const file = path.join(dir, name);
    fs.copyFileSync(path.join(FIXTURES, fixture), file);
    return file;
}

afterEach(() => {
    for (const dir of dirs.splice(0)) {
        fs.rmSync(dir, { recursive: true, force: true });
    }
});

const part = (file: string, name: string): string => ZipArchive.read(fs.readFileSync(file)).read(name).toString('utf8');
const projectStream = (file: string, bin: string): string =>
    Cfb.fromBytes(ZipArchive.read(fs.readFileSync(file)).read(bin)).getStream('PROJECT').toString('latin1');

describe('adding a VBA project to a workbook', () => {
    it('makes the modules Excel makes, named the way Excel names them', () => {
        // A chart sheet in front of three worksheets, renamed and reordered.
        // Excel numbers a new code name by the sheet's position among the
        // worksheets and chart sheets, not by when it was made.
        const file = scratch('NoVbaSheetsFixture.xlsm');
        expect(hasVbaProject(file)).toBe(false);

        const added = addVbaProject(file, path.join(TEMPLATES, 'blank.xlsm'));

        expect(added.modules).toEqual(['ThisWorkbook', 'Chart1', 'Sheet2', 'Sheet3', 'Sheet4']);
        expect(listModules(file).map((module) => [module.name, module.type, (module as { documentType?: string }).documentType])).toEqual([
            ['ThisWorkbook', 'document', 'workbook'],
            ['Chart1', 'document', 'chart'],
            ['Sheet2', 'document', 'worksheet'],
            ['Sheet3', 'document', 'worksheet'],
            ['Sheet4', 'document', 'worksheet'],
        ]);
        // The template's starter module is not carried over.
        expect(listModules(file).some((module) => module.name === 'Module1')).toBe(false);
    });

    it('writes each code name into the package, as Excel does', () => {
        const file = scratch('NoVbaSheetsFixture.xlsm');
        addVbaProject(file, path.join(TEMPLATES, 'blank.xlsm'));

        expect(part(file, 'xl/workbook.xml')).toContain('<workbookPr codeName="ThisWorkbook" defaultThemeVersion="202300"/>');
        // An empty sheetPr takes the attribute; a sheet with none gets one
        // as its first element.
        expect(part(file, 'xl/chartsheets/sheet1.xml')).toContain('<sheetPr codeName="Chart1"/>');
        for (const [sheet, codeName] of [['sheet1', 'Sheet2'], ['sheet2', 'Sheet3'], ['sheet3', 'Sheet4']]) {
            expect(part(file, `xl/worksheets/${sheet}.xml`)).toMatch(new RegExp(`<worksheet\\b[^>]*><sheetPr codeName="${codeName}"/><dimension`));
        }
    });

    it('declares each module the way Excel does, a chart sheet with its own base', () => {
        const file = scratch('NoVbaSheetsFixture.xlsm');
        addVbaProject(file, path.join(TEMPLATES, 'blank.xlsm'));

        const project = projectStream(file, 'xl/vbaProject.bin');
        for (const name of ['ThisWorkbook', 'Chart1', 'Sheet2', 'Sheet3', 'Sheet4']) {
            expect(project).toContain(`Document=${name}/&H00000000\r\n`);
        }
        expect(project).not.toMatch(/^Module=/m);
        const header = (name: string): string => readModule(file, name, true).source;
        expect(header('Chart1')).toContain('Attribute VB_Base = "0{00020821-0000-0000-C000-000000000046}"');
        expect(header('Sheet2')).toContain('Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"');
        expect(header('ThisWorkbook')).toContain('Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"');
        expect(header('Sheet4')).toMatch(/^Attribute VB_Name = "Sheet4"\r\n/);
    });

    it('relates the part and declares its type the way Excel does', () => {
        const file = scratch('NoVbaSheetsFixture.xlsm');
        addVbaProject(file, path.join(TEMPLATES, 'blank.xlsm'));

        expect(part(file, '[Content_Types].xml')).toContain('<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/>');
        // The workbook's last relationship was rId7.
        expect(part(file, 'xl/_rels/workbook.xml.rels'))
            .toContain('<Relationship Id="rId8" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/>');
    });

    it('takes a module afterwards like any project', () => {
        const file = scratch('NoVbaSheetsFixture.xlsm');
        addVbaProject(file, path.join(TEMPLATES, 'blank.xlsm'));
        writeModule(file, 'Module1', 'Public Sub Go()\r\nEnd Sub\r\n', 'standard');
        expect(listModules(file).map((module) => module.name)).toEqual(['ThisWorkbook', 'Chart1', 'Sheet2', 'Sheet3', 'Sheet4', 'Module1']);
    });

    it('gives dialog sheets and Excel 4 macro sheets neither a code name nor a module', () => {
        const file = scratch('NoVbaOtherSheetsFixture.xlsm');
        addVbaProject(file, path.join(TEMPLATES, 'blank.xlsm'));
        expect(listModules(file).map((module) => module.name)).toEqual(['ThisWorkbook', 'Sheet1']);
        expect(part(file, 'xl/dialogsheets/sheet1.xml')).not.toContain('codeName');
        expect(part(file, 'xl/macrosheets/sheet1.xml')).not.toContain('codeName');
    });

    it('keeps a code name a sheet already has, and numbers the rest around it', () => {
        // Excel writes code names when the VBE is opened, and drops the
        // project again if no code was written: a file can have both.
        const file = scratch('NoVbaSheetsFixture.xlsm');
        const zip = ZipArchive.read(fs.readFileSync(file));
        zip.write('xl/worksheets/sheet3.xml', Buffer.from(part(file, 'xl/worksheets/sheet3.xml')
            .replace(/(<worksheet\b[^>]*>)/, '$1<sheetPr codeName="Sheet2"/>'), 'utf8'));
        fs.writeFileSync(file, zip.toBytes());

        addVbaProject(file, path.join(TEMPLATES, 'blank.xlsm'));

        // Report, at position 2, would have been Sheet2, which Sheet4 holds.
        expect(listModules(file).map((module) => module.name)).toEqual(['ThisWorkbook', 'Chart1', 'Sheet3', 'Sheet4', 'Sheet2']);
        expect(part(file, 'xl/worksheets/sheet3.xml').match(/codeName="/g)).toHaveLength(1);
    });

    it('gives the workbook the macro-enabled type its extension needs', () => {
        // A workbook saved as .xlsx and renamed .xlsm: Excel refuses it
        // under that extension, VBA or not.
        const file = scratch('NoVbaFixture.xlsm');
        const zip = ZipArchive.read(fs.readFileSync(file));
        zip.write('[Content_Types].xml', Buffer.from(part(file, '[Content_Types].xml').replace(
            'application/vnd.ms-excel.sheet.macroEnabled.main+xml',
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
        ), 'utf8'));
        fs.writeFileSync(file, zip.toBytes());

        addVbaProject(file, path.join(TEMPLATES, 'blank.xlsm'));

        expect(part(file, '[Content_Types].xml'))
            .toContain('<Override PartName="/xl/workbook.xml" ContentType="application/vnd.ms-excel.sheet.macroEnabled.main+xml"/>');
    });
});

describe('adding a VBA project to a binary workbook', () => {
    // The same workbook saved as .xlsb, and Excel's own copy of it with VBA
    // added. Excel filled the empty code name in BrtWbProp, each BrtWsProp
    // and the BrtCsProp, and moved the one binary-index offset that pointed
    // past a record that grew.

    /** A binary part as its records, type and payload ([MS-XLSB] 2.1.4). */
    function records(data: Buffer): Array<[number, string]> {
        const out: Array<[number, string]> = [];
        let at = 0;
        const varint = (max: number): number => {
            let value = 0;
            for (let i = 0; i < max; i++) {
                const byte = data[at++];
                value += (byte & 0x7f) * 2 ** (7 * i);
                if (!(byte & 0x80)) { return value; }
            }
            throw new Error('bad varint');
        };
        while (at < data.length) {
            const type = varint(2);
            const size = varint(4);
            out.push([type, data.subarray(at, at + size).toString('hex')]);
            at += size;
        }
        expect(at).toBe(data.length);
        return out;
    }

    it('writes the records Excel writes', () => {
        const file = scratch('NoVbaSheetsFixture.xlsb');
        const added = addVbaProject(file, path.join(TEMPLATES, 'blank.xlsb'));
        expect(added.modules).toEqual(['ThisWorkbook', 'Chart1', 'Sheet2', 'Sheet3', 'Sheet4']);

        const ours = ZipArchive.read(fs.readFileSync(file));
        const excels = ZipArchive.read(fs.readFileSync(path.join(FIXTURES, 'VbaAddedByExcelFixture.xlsb')));
        const binaryParts = excels.names().filter((name) => name.endsWith('.bin') && !name.endsWith('vbaProject.bin'));
        expect(binaryParts).toContain('xl/worksheets/binaryIndex1.bin');
        for (const name of binaryParts) {
            const mine = records(ours.read(name));
            const theirs = records(excels.read(name));
            if (name === 'xl/workbook.bin') {
                // BrtFileVersion (128) is the one difference: Excel stamps a
                // GUID for the project there, and opens a workbook without one.
                const firstDiffering = mine.findIndex(([type, payload], i) => type !== theirs[i][0] || payload !== theirs[i][1]);
                expect(mine[firstDiffering][0]).toBe(128);
                mine.splice(firstDiffering, 1);
                theirs.splice(firstDiffering, 1);
            }
            expect(mine, name).toEqual(theirs);
        }
    });

    it('declares the part with an override, since .bin already means the workbook', () => {
        const file = scratch('NoVbaSheetsFixture.xlsb');
        addVbaProject(file, path.join(TEMPLATES, 'blank.xlsb'));
        const types = part(file, '[Content_Types].xml');
        expect(types).toContain('<Default Extension="bin" ContentType="application/vnd.ms-excel.sheet.binary.macroEnabled.main"/>');
        expect(types).toContain('<Override PartName="/xl/vbaProject.bin" ContentType="application/vnd.ms-office.vbaProject"/>');
        expect(part(file, 'xl/_rels/workbook.bin.rels')).toContain('Target="vbaProject.bin"');
    });
});

describe('adding a VBA project to a document or a presentation', () => {
    it('gives a Word document ThisDocument, as Word does', () => {
        const file = scratch('NoVbaFixture.docm');
        const added = addVbaProject(file, path.join(TEMPLATES, 'blank.docm'));
        expect(added.modules).toEqual(['ThisDocument']);
        expect(readModule(file, 'ThisDocument', true).source).toContain('Attribute VB_Base = "1Normal.ThisDocument"');
        expect(part(file, '[Content_Types].xml')).toContain('<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/>');
        expect(part(file, 'word/_rels/document.xml.rels')).toContain('Target="vbaProject.bin"');
        expect(projectStream(file, 'word/vbaProject.bin')).toContain('Name="Project"');
    });

    it('gives a presentation an empty project, as PowerPoint does', () => {
        const file = scratch('NoVbaFixture.pptm');
        const added = addVbaProject(file, path.join(TEMPLATES, 'blank.pptm'));
        expect(added.modules).toEqual([]);
        expect(hasVbaProject(file)).toBe(true);
        expect(listModules(file)).toEqual([]);
        expect(part(file, 'ppt/_rels/presentation.xml.rels')).toContain('Target="vbaProject.bin"');
        writeModule(file, 'Module1', 'Public Sub Go()\r\nEnd Sub\r\n', 'standard');
        expect(listModules(file).map((module) => module.name)).toEqual(['Module1']);
    });
});

describe('what cannot take a new VBA project', () => {
    it('a file that already has one', () => {
        const file = scratch('FormFixture.xlsm');
        expect(() => addVbaProject(file, path.join(TEMPLATES, 'blank.xlsm'))).toThrow('The file already has a VBA project.');
    });

    it('a package that is not an Office document', () => {
        const zip = ZipArchive.read(fs.readFileSync(path.join(FIXTURES, 'NoVbaFixture.xlsm')));
        zip.delete('_rels/.rels');
        expect(() => addVbaProjectToPackage(zip.toBytes(), 'xlsm', fs.readFileSync(path.join(TEMPLATES, 'blank.xlsm'))))
            .toThrow(AddVbaProjectError);
    });

    it('a legacy file, whose project lives in records XLIDE does not write', () => {
        const file = scratch('NoVbaFixture.xls');
        expect(() => addVbaProject(file, path.join(TEMPLATES, 'blank.xlsm'))).toThrow(/not an Office Open XML file/);
        expect(hasVbaProject(file)).toBe(false);
    });
});
