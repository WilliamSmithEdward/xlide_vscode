import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
vi.mock('vscode', async () => (await import('./helpers/vscodeMock')).vscodeMock({ window: { showInputBox: vi.fn() } }));
import * as vscode from 'vscode';
import { openFileInHost, runHostMacro, showAccessDesign } from '../src/officeHostLauncher';
import { runPowerShell } from '../src/util/powershell';
vi.mock('../src/util/powershell', async original => ({ ...(await original<typeof import('../src/util/powershell')>()), runPowerShell: vi.fn() }));
import { VbaProjectProtection, decodeProtectionRecord, clearVbaProjectAuthorizations } from '../src/vba/projectProtection';
import * as svc from '../src/vba/projectService';
import { ProjectEngine } from '../src/projectEngine';
import { ensureProjectPassword } from '../src/projectPasswordPrompt';

import { hashed, record, raw, protectFixture, FIXTURE_PASSWORD } from './helpers/projectProtectionFixture';
let dir: string;
let file: string;
let moduleName: string;
function protect(password: string, target = file): void { protectFixture(target, password); }
beforeEach(() => {
    clearVbaProjectAuthorizations(); vi.clearAllMocks();
    dir = fs.mkdtempSync(path.join(os.tmpdir(), 'xlide-protected-'));
    file = path.join(dir, 'Book.xlsm');
    fs.copyFileSync(path.join(__dirname, '../assets/templates/blank.xlsm'), file);
    moduleName = svc.listModules(file)[0].name;
});
afterEach(() => { clearVbaProjectAuthorizations(); fs.rmSync(dir, { recursive: true, force: true }); });

describe('MS-OVBA password verification', () => {
    it('decodes the published Office no-password example', () => {
        expect(decodeProtectionRecord('0E0CD1ECDFF4E7F5E7F5E7')).toEqual(Buffer.from([0]));
        expect(new VbaProjectProtection(Buffer.from('CMG="0705D8E3D8EDDBF1DBF1DBF1DBF1"\r\nDPB="0E0CD1ECDFF4E7F5E7F5E7"\r\nGC="1517CAF1D6F9D7F9D706"'), 1252).requiresPassword).toBe(false);
    });
    it.each([0, 2, 4, 6])('handles filler length from seed %i', seed => {
        const data = hashed('Correct!');
        expect(decodeProtectionRecord(record(data, seed))).toEqual(data);
    });
    it.each([[1252, 'café'], [1251, 'Пароль'], [932, '日本語']])('uses project code page %i', (cp, password) => {
        const protection = new VbaProjectProtection(raw(password as string, cp as number), cp as number);
        expect(protection.verify(password as string)).toBe(true);
        expect(protection.verify('incorrect')).toBe(false);
    });
    it('verifies legacy short passwords and rejects lossy encoding', () => {
        const protection = new VbaProjectProtection(raw('?', 1252, true), 1252);
        expect(protection.hasPassword).toBe(true);
        expect(protection.verify('?')).toBe(true);
        expect(protection.verify('😀')).toBe(false);
    });
    it('fails closed when the protection stream is unavailable', () => {
        const protection = new VbaProjectProtection(undefined, 1252);
        expect(protection.requiresPassword).toBe(true);
        expect(() => protection.verify('anything')).toThrow(/malformed or unsupported/);
    });
    it.each(['DPB="00"', 'DPB="bad"', 'DPB=""', 'CMG="00"', 'GC="00"'])('fails closed for %s', text => {
        const protection = new VbaProjectProtection(Buffer.from(text), 1252);
        expect(protection.requiresPassword).toBe(true);
        expect(() => protection.verify('anything')).toThrow(/malformed or unsupported/);
    });
});

describe('protected file access', () => {
    it('uses the permanent protected workbook fixture', () => {
        const permanent = path.join(__dirname, 'fixtures/binaries/PasswordProtectedFixture.xlsm');
        expect(svc.getProtectionInfo(permanent)).toMatchObject({ isPasswordProtected: true, isAccessLocked: true });
        expect(() => svc.readModule(permanent, 'Runner')).toThrow(/password-protected/);
        expect(svc.unlockProject(permanent, FIXTURE_PASSWORD)).toBe(true);
        expect(svc.getProtectionInfo(permanent)).toMatchObject({ isPasswordProtected: true, isAccessLocked: false });
        expect(svc.readModule(permanent, 'Runner').source).toContain('counter = 1');
        const form = path.join(__dirname, 'fixtures/binaries/PasswordProtectedFormFixture.xlsm');
        expect(svc.unlockProject(form, FIXTURE_PASSWORD)).toBe(true);
        expect(svc.readFormPreview(form, 'FrmPicker')).toBeDefined();
    });
    it.each(['SheetsFixture.xlsb', 'XlsFixture.xls', 'WordFixture.docm', 'WordFixture.doc', 'PowerPointFixture.pptm', 'PowerPointFixture.ppt', 'AccessFixture.accdb', 'AccessFixture.mdb'])('enforces the password for %s', name => {
        const target = path.join(dir, name);
        fs.copyFileSync(path.join(__dirname, 'fixtures/binaries', name), target);
        protect('Correct!', target);
        expect(() => svc.listModules(target)).toThrow(/password-protected/);
        expect(svc.unlockProject(target, 'wrong')).toBe(false);
        expect(svc.unlockProject(target, 'Correct!')).toBe(true);
        expect(svc.listModules(target).length).toBeGreaterThan(0);
    });
    it('refuses Office launches and macro/form execution after cancellation', async () => {
        protect('Correct!');
        vi.mocked(vscode.window.showInputBox).mockResolvedValue(undefined);
        const calls = [
            () => openFileInHost(file, { attachToRunning: true, readOnly: false }, vi.fn()),
            () => runHostMacro(file, { moduleName, procedureName: 'Run' }, { attachToRunning: true }, vi.fn()),
            () => showAccessDesign(file, { kind: 'form', name: 'Form1' }, { attachToRunning: true }, vi.fn()),
        ];
        for (const call of calls) { await expect(call()).rejects.toBeInstanceOf(vscode.CancellationError); }
        expect(runPowerShell).not.toHaveBeenCalled();
    });

    it('refuses service reads, writes, exports, analysis and workbook interactions without altering bytes', () => {
        protect('Correct!'); const before = fs.readFileSync(file);
        expect(svc.getProtectionInfo(file).isPasswordProtected).toBe(true);
        const operations = [
            () => svc.listModules(file), () => svc.readModule(file, moduleName),
            () => svc.readModules(file), () => svc.readModulesFromBuffer(before),
            () => svc.getProjectInfo(file), () => svc.getModulesAndProtectionInfo(file),
            () => svc.listSubs(file, moduleName), () => svc.listReferences(file),
            () => svc.validateProject(file), () => svc.readFormExport(file, moduleName),
            () => svc.readFormPreview(file, moduleName), () => svc.readFormMarkup(file, moduleName),
            () => svc.writeModule(file, moduleName, 'Option Explicit'),
            () => svc.renameModule(file, moduleName, 'Changed'), () => svc.deleteModule(file, moduleName),
            () => svc.readCells(file, 'Sheet1', 'A1'), () => svc.listSheets(file),
            () => svc.writeCells(file, 'Sheet1', 'A1', [[1]]),
            () => svc.listShapes(file), () => svc.shapeMacros(file),
            () => svc.createProject(file, path.join(__dirname, '../assets/templates/blank.xlsm')),
            () => svc.addReference(file, 'word'), () => svc.removeReference(file, 'Excel'),
        ];
        for (const operation of operations) { expect(operation).toThrow(/password-protected/); }
        expect(svc.unlockProject(file, 'wrong')).toBe(false);
        expect(fs.readFileSync(file)).toEqual(before);
        expect(svc.unlockProject(file, 'Correct!')).toBe(true);
        expect(svc.readModule(file, moduleName).source).toBeTypeOf('string');
        expect(svc.readModulesFromBuffer(before, false, file).length).toBeGreaterThan(0);
        expect(() => svc.readModulesFromBuffer(before)).toThrow(/password-protected/);
        svc.writeModule(file, moduleName, svc.readModule(file, moduleName).source + '\r\n\' verified write\r\n');
        expect(svc.getProtectionInfo(file).isPasswordProtected).toBe(true);
        expect(svc.isProjectAccessLocked(file)).toBe(false);
    });
    it('historical unprotected buffers do not revoke the live file authorization', () => {
        protect('Correct!'); expect(svc.unlockProject(file, 'Correct!')).toBe(true);
        const history = fs.readFileSync(path.join(__dirname, '../assets/templates/blank.xlsm'));
        svc.readModulesFromBuffer(history, false, file);
        expect(svc.isProjectAccessLocked(file)).toBe(false);
    });
    it('refusing differently protected history preserves current authorization', () => {
        protect('Old'); const history = fs.readFileSync(file);
        protect('Current'); expect(svc.unlockProject(file, 'Current')).toBe(true);
        expect(() => svc.readModulesFromBuffer(history, false, file)).toThrow(/password-protected/);
        expect(svc.isProjectAccessLocked(file)).toBe(false);
    });
    it('relocks changed protection and does not authorize another file', () => {
        protect('First'); expect(svc.unlockProject(file, 'First')).toBe(true);
        const other = path.join(dir, 'Other.xlsm'); fs.copyFileSync(file, other);
        expect(() => svc.readModule(other, moduleName)).toThrow(/password-protected/);
        protect('Second'); expect(() => svc.readModule(file, moduleName)).toThrow(/password-protected/);
        expect(svc.unlockProject(file, 'First')).toBe(false);
        expect(svc.unlockProject(file, 'Second')).toBe(true);
        clearVbaProjectAuthorizations(); expect(svc.isProjectAccessLocked(file)).toBe(true);
    });
    it('prompts privately, retries an incorrect password and reuses authorization', async () => {
        protect('Correct!');
        vi.mocked(vscode.window.showInputBox).mockResolvedValueOnce('wrong').mockResolvedValueOnce('Correct!');
        const engine = new ProjectEngine({} as never);
        await engine.call('readModule', { path: file, module: moduleName });
        await engine.call('getProjectInfo', { path: file });
        expect(vscode.window.showInputBox).toHaveBeenCalledTimes(2);
        expect(vscode.window.showInputBox).toHaveBeenCalledWith(expect.objectContaining({ password: true, ignoreFocusOut: true }));
        engine.dispose(); expect(svc.isProjectAccessLocked(file)).toBe(true);
    });
    it('cancellation refuses an agent-style write and leaves the file intact', async () => {
        protect('Correct!'); const before = fs.readFileSync(file);
        vi.mocked(vscode.window.showInputBox).mockResolvedValue(undefined);
        const engine = new ProjectEngine({} as never);
        await expect(engine.call('writeModule', { path: file, module: moduleName, source: 'Changed' })).rejects.toBeInstanceOf(vscode.CancellationError);
        expect(fs.readFileSync(file)).toEqual(before);
        expect(svc.isProjectAccessLocked(file)).toBe(true);
    });
    it('shares concurrent prompts and does not prompt for unprotected files', async () => {
        await ensureProjectPassword(file); expect(vscode.window.showInputBox).not.toHaveBeenCalled();
        protect('Correct!');
        let respond!: (password: string) => void;
        vi.mocked(vscode.window.showInputBox).mockImplementationOnce(() => new Promise(resolve => { respond = resolve; }));
        const first = ensureProjectPassword(file); const second = ensureProjectPassword(file);
        await vi.waitFor(() => expect(vscode.window.showInputBox).toHaveBeenCalledTimes(1));
        respond('Correct!'); await Promise.all([first, second]);
    });
});
