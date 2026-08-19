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

// [MS-OVBA] stores module names in the project's ANSI code page. Before the
// guard, renaming a module to a name the page cannot represent was accepted
// and '?'-folded on save, which detached the module from its source stream -
// the code was simply gone on the next open - and let two distinct names
// collide into one. The engine now refuses such names up front, while names
// the project's own code page can store keep working, non-ASCII included.

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

describe('module names vs the project code page', () => {
    it('refuses a rename the cp1252 project cannot store, leaving the module intact', () => {
        const standard = listModules(workbook).find((m) => m.type === 'standard');
        expect(standard).toBeDefined();

        expect(() => renameModule(workbook, standard!.name, 'Модуль1'))
            .toThrow(/cannot be stored in this project's code page \(1252\)/);

        const names = listModules(workbook).map((m) => m.name);
        expect(names).toContain(standard!.name);
        expect(names.some((n) => n.includes('?'))).toBe(false);
        expect(readModule(workbook, standard!.name).source.length).toBeGreaterThan(0);
        expect(validateWorkbook(workbook).issues).toEqual([]);
    });

    it('refuses creating a new module whose name the code page cannot store', () => {
        expect(() => writeModule(workbook, 'Модель1', 'Sub A()\r\nEnd Sub\r\n', 'standard'))
            .toThrow(/cannot be stored in this project's code page/);

        expect(listModules(workbook).map((m) => m.name).some((n) => n.includes('?'))).toBe(false);
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
     * records. Driving the project layer directly sits below the service
     * guard, exactly like a file authored elsewhere.
     */
    function synthesizeCrossLocaleWorkbook(target: string): void {
        const container = openMacroContainer(fs.readFileSync(FIXTURE));
        const cfb = container.vbaCfb();
        const project = VbaProject.parse(cfb);
        project.renameModule('XlideFormProbe', 'Модуль1');
        project.save(cfb);
        fs.writeFileSync(target, container.toFileBytes(cfb));
    }

    it('reads the module source through the unicode stream name', () => {
        synthesizeCrossLocaleWorkbook(workbook);

        // The outward name is still the ANSI projection; what matters is the
        // module's code resolves instead of reading as an empty missing
        // stream, and validation agrees the project is whole.
        const names = listModules(workbook).map((m) => m.name);
        expect(names).toContain('??????1');
        expect(readModule(workbook, '??????1').source).toContain('Option Explicit');
        expect(validateWorkbook(workbook).issues).toEqual([]);
    });

    it('writes back through the resolved stream name without detaching it', () => {
        synthesizeCrossLocaleWorkbook(workbook);

        writeModule(workbook, '??????1', 'Sub Replaced()\r\nEnd Sub\r\n', 'standard');

        expect(readModule(workbook, '??????1').source).toContain('Sub Replaced()');
        expect(validateWorkbook(workbook).issues).toEqual([]);
    });
});
