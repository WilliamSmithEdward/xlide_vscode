import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';

import {
    listModules,
    readModule,
    renameModule,
    validateProject,
    writeModule,
} from '../src/vba/projectService';
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
let project: string;

beforeEach(() => {
    tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'xlide-name-codepage-'));
    project = path.join(tempDir, 'Probe.xlsm');
    fs.copyFileSync(FIXTURE, project);
});

afterEach(() => {
    fs.rmSync(tempDir, { recursive: true, force: true });
});

describe('module names beyond the project code page', () => {
    it('renames a module to a unicode name a cp1252 project cannot fold, and back', () => {
        const standard = listModules(project).find((m) => m.type === 'standard');
        expect(standard).toBeDefined();

        renameModule(project, standard!.name, 'Модуль1');

        expect(listModules(project).map((m) => m.name)).toContain('Модуль1');
        expect(readModule(project, 'Модуль1').source.length).toBeGreaterThan(0);
        expect(validateProject(project).issues).toEqual([]);

        renameModule(project, 'Модуль1', standard!.name);
        expect(listModules(project).map((m) => m.name)).toContain(standard!.name);
        expect(validateProject(project).issues).toEqual([]);
    });

    it('creates a new module under a unicode name and round-trips its source', () => {
        writeModule(project, 'Модель1', 'Sub A()\r\nEnd Sub\r\n', 'standard');

        expect(listModules(project).map((m) => m.name)).toContain('Модель1');
        expect(readModule(project, 'Модель1').source).toContain('Sub A()');
        expect(validateProject(project).issues).toEqual([]);
    });

    it('refuses a second name whose folded projection collides', () => {
        writeModule(project, 'Модуль1', 'Sub A()\r\nEnd Sub\r\n', 'standard');

        // 'Модель1' folds to the same '??????1' projection as 'Модуль1'.
        expect(() => writeModule(project, 'Модель1', 'Sub B()\r\nEnd Sub\r\n', 'standard'))
            .toThrow(/stores both as "\?\?\?\?\?\?1"/);

        // A distinct projection coexists - but renaming it INTO the taken
        // projection is refused the same way.
        writeModule(project, 'Модель2', 'Sub C()\r\nEnd Sub\r\n', 'standard');
        expect(() => renameModule(project, 'Модель2', 'Записъ1'))
            .toThrow(/stores both as/);
        expect(listModules(project).map((m) => m.name))
            .toEqual(expect.arrayContaining(['Модуль1', 'Модель2']));
        expect(validateProject(project).issues).toEqual([]);
    });

    it('accepts a non-ASCII name the code page does store, round-tripping it', () => {
        const standard = listModules(project).find((m) => m.type === 'standard');
        expect(standard).toBeDefined();

        renameModule(project, standard!.name, 'Módulo1');

        const names = listModules(project).map((m) => m.name);
        expect(names).toContain('Módulo1');
        expect(readModule(project, 'Módulo1').source.length).toBeGreaterThan(0);
        expect(validateProject(project).issues).toEqual([]);
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
        synthesizeCrossLocaleWorkbook(project);

        const names = listModules(project).map((m) => m.name);
        expect(names).toContain('Модуль1');
        expect(names.some((n) => n.includes('?'))).toBe(false);
        expect(readModule(project, 'Модуль1').source).toContain('Option Explicit');
        expect(validateProject(project).issues).toEqual([]);
    });

    it('writes back through the resolved stream name without detaching it', () => {
        synthesizeCrossLocaleWorkbook(project);

        writeModule(project, 'Модуль1', 'Sub Replaced()\r\nEnd Sub\r\n', 'standard');

        expect(readModule(project, 'Модуль1').source).toContain('Sub Replaced()');
        expect(validateProject(project).issues).toEqual([]);
    });

    it('renames a cross-locale module back to ASCII through the folded PROJECT line', () => {
        synthesizeCrossLocaleWorkbook(project);

        renameModule(project, 'Модуль1', 'RenamedBack');

        const names = listModules(project).map((m) => m.name);
        expect(names).toContain('RenamedBack');
        expect(names.some((n) => n.includes('?'))).toBe(false);
        expect(readModule(project, 'RenamedBack').source).toContain('Option Explicit');
        expect(validateProject(project).issues).toEqual([]);
    });
});
