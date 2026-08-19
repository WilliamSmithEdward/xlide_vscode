import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';

import {
    listModules,
    readModule,
    renameModule,
    validateWorkbook,
    writeModule,
} from '../src/vba/workbookService';
import { openMacroContainer } from '../src/vba/macroContainer';
import { VbaProject } from '../src/vba/vbaProject';

// Module names beyond the project's ANSI code page are fully supported: the
// unicode dir records and the CFB stream name carry the real name while the
// ANSI records and the PROJECT stream hold its '?'-folded projection - the
// same shape Office produces. Verified against live Excel (2026-08-18): the
// VBE lists the unicode name, Application.Run executes the module, Excel
// re-saves it, and the engine reads the Excel-authored result back intact.
// The one refusal left is a REAL hazard: two names whose folded projections
// collide would declare the same name twice in the PROJECT stream, which
// Excel treats as corruption.

const FIXTURE = path.join(__dirname, 'fixtures', 'binaries', 'FormFixture.xlsm');

let tempDir: string;
let workbook: string;

beforeEach(() => {
    tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'xlide-name-codepage-'));
    workbook = path.join(tempDir, 'Probe.xlsm');
    fs.copyFileSync(FIXTURE, workbook);
});

afterEach(() => {
    fs.rmSync(tempDir, { recursive: true, force: true });
});

describe('module names beyond the project code page', () => {
    it('renames a module to a unicode name a cp1252 project cannot fold, and back', () => {
        const standard = listModules(workbook).find((m) => m.type === 'standard');
        expect(standard).toBeDefined();

        renameModule(workbook, standard!.name, 'Модуль1');

        expect(listModules(workbook).map((m) => m.name)).toContain('Модуль1');
        expect(readModule(workbook, 'Модуль1').source.length).toBeGreaterThan(0);
        expect(validateWorkbook(workbook).issues).toEqual([]);

        renameModule(workbook, 'Модуль1', standard!.name);
        expect(listModules(workbook).map((m) => m.name)).toContain(standard!.name);
        expect(validateWorkbook(workbook).issues).toEqual([]);
    });

    it('creates a new module under a unicode name and round-trips its source', () => {
        writeModule(workbook, 'Модель1', 'Sub A()\r\nEnd Sub\r\n', 'standard');

        expect(listModules(workbook).map((m) => m.name)).toContain('Модель1');
        expect(readModule(workbook, 'Модель1').source).toContain('Sub A()');
        expect(validateWorkbook(workbook).issues).toEqual([]);
    });

    it('refuses a second name whose folded projection collides', () => {
        writeModule(workbook, 'Модуль1', 'Sub A()\r\nEnd Sub\r\n', 'standard');

        // 'Модель1' folds to the same '??????1' projection as 'Модуль1'.
        expect(() => writeModule(workbook, 'Модель1', 'Sub B()\r\nEnd Sub\r\n', 'standard'))
            .toThrow(/stores both as "\?\?\?\?\?\?1"/);

        // A distinct projection coexists - but renaming it INTO the taken
        // projection is refused the same way.
        writeModule(workbook, 'Модель2', 'Sub C()\r\nEnd Sub\r\n', 'standard');
        expect(() => renameModule(workbook, 'Модель2', 'Записъ1'))
            .toThrow(/stores both as/);
        expect(listModules(workbook).map((m) => m.name))
            .toEqual(expect.arrayContaining(['Модуль1', 'Модель2']));
        expect(validateWorkbook(workbook).issues).toEqual([]);
    });

    it('accepts a non-ASCII name the code page does store, round-tripping it', () => {
        const standard = listModules(workbook).find((m) => m.type === 'standard');
        expect(standard).toBeDefined();

        renameModule(workbook, standard!.name, 'Módulo1');

        const names = listModules(workbook).map((m) => m.name);
        expect(names).toContain('Módulo1');
        expect(readModule(workbook, 'Módulo1').source.length).toBeGreaterThan(0);
        expect(validateWorkbook(workbook).issues).toEqual([]);
    });
});

describe('cross-locale projects with unicode stream names', () => {
    /**
     * Builds the on-disk shape Office produces for a module name beyond the
     * project's code page: the real UTF-16 name in the unicode dir records
     * and the CFB directory, and the '?'-folded projection in the ANSI
     * records. Driving the project layer directly, exactly like a file
     * authored elsewhere.
     */
    function synthesizeCrossLocaleWorkbook(target: string): void {
        const container = openMacroContainer(fs.readFileSync(FIXTURE));
        const cfb = container.vbaCfb();
        const project = VbaProject.parse(cfb);
        project.renameModule('XlideFormProbe', 'Модуль1');
        project.save(cfb);
        fs.writeFileSync(target, container.toFileBytes(cfb));
    }

    it('surfaces the real unicode name and reads the module source', () => {
        synthesizeCrossLocaleWorkbook(workbook);

        const names = listModules(workbook).map((m) => m.name);
        expect(names).toContain('Модуль1');
        expect(names.some((n) => n.includes('?'))).toBe(false);
        expect(readModule(workbook, 'Модуль1').source).toContain('Option Explicit');
        expect(validateWorkbook(workbook).issues).toEqual([]);
    });

    it('writes back through the resolved stream name without detaching it', () => {
        synthesizeCrossLocaleWorkbook(workbook);

        writeModule(workbook, 'Модуль1', 'Sub Replaced()\r\nEnd Sub\r\n', 'standard');

        expect(readModule(workbook, 'Модуль1').source).toContain('Sub Replaced()');
        expect(validateWorkbook(workbook).issues).toEqual([]);
    });

    it('renames a cross-locale module back to ASCII through the folded PROJECT line', () => {
        synthesizeCrossLocaleWorkbook(workbook);

        renameModule(workbook, 'Модуль1', 'RenamedBack');

        const names = listModules(workbook).map((m) => m.name);
        expect(names).toContain('RenamedBack');
        expect(names.some((n) => n.includes('?'))).toBe(false);
        expect(readModule(workbook, 'RenamedBack').source).toContain('Option Explicit');
        expect(validateWorkbook(workbook).issues).toEqual([]);
    });
});
