// The code page VB6 files are read and written in (issue #205). VB6 saved in
// the ANSI code page of the Windows it ran on; a Russian Windows writes
// cp1251, where the Cyrillic capital Che is byte 0xD7, which cp1252 reads as
// the multiplication sign.

import { afterEach, describe, expect, it } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { decodeCodePage, encodeCodePage } from '../src/vba/codePages';
import { hostPlatform, setHostPlatform, type HostPlatform } from '../src/vba/hostPlatform';
import {
	readVb6Modules,
	resetVb6ProjectCacheForTests,
	setVb6CodePage,
	vb6CodePage,
	writeVb6Module,
} from '../src/vba/vb6/vb6Project';

const CRLF = '\r\n';
const MODULE = [
	'Attribute VB_Name = "\u041c\u043e\u0434\u0443\u043b\u044c1"',
	'Option Explicit',
	'Public Function \u0418\u0442\u043e\u0433() As Long',
	'    Dim \u0427\u0438\u0441\u043b\u043e As Long',
	'    \u0427\u0438\u0441\u043b\u043e = 2',
	'    \u0418\u0442\u043e\u0433 = \u0427\u0438\u0441\u043b\u043e * 3',
	'End Function',
	'',
].join(CRLF);
const MANIFEST = ['Type=Exe', 'Module=\u041c\u043e\u0434\u0443\u043b\u044c1; \u041c\u043e\u0434\u0443\u043b\u044c1.bas', ''].join(CRLF);

const realPlatform = hostPlatform();
const tempDirs: string[] = [];
afterEach(() => {
	setVb6CodePage(undefined);
	setHostPlatform(realPlatform);
	resetVb6ProjectCacheForTests();
	for (const dir of tempDirs.splice(0)) {
		fs.rmSync(dir, { recursive: true, force: true });
	}
});

/** A project whose files are the given texts, each written in `codePage`. */
function project(codePage: number, files: Record<string, string | Buffer> = { 'Project1.vbp': MANIFEST, '\u041c\u043e\u0434\u0443\u043b\u044c1.bas': MODULE }): string {
	const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'xlide-vb6-cp-'));
	tempDirs.push(dir);
	for (const [name, text] of Object.entries(files)) {
		fs.writeFileSync(path.join(dir, name), typeof text === 'string' ? encodeCodePage(text, codePage) : text);
	}
	return path.join(dir, 'Project1.vbp');
}

/** The real platform, answering `page` for the machine's ANSI code page. */
function machineWithCodePage(page: number | undefined): HostPlatform {
	return { ...realPlatform, ansiCodePage: () => page };
}

describe('VB6 files in the code page they were saved in (issue #205)', () => {
	it('reads a cp1251 project in the page the setting names', () => {
		setVb6CodePage(1251);
		const [module] = readVb6Modules(project(1251), true);
		expect(module.name).toBe('\u041c\u043e\u0434\u0443\u043b\u044c1');
		expect(module.source).toContain('Dim \u0427\u0438\u0441\u043b\u043e As Long');
		expect(module.filePath.endsWith('\u041c\u043e\u0434\u0443\u043b\u044c1.bas')).toBe(true);
	});

	it("reads it in the machine's own page when no setting names one", () => {
		setHostPlatform(machineWithCodePage(1251));
		expect(vb6CodePage()).toBe(1251);
		const [module] = readVb6Modules(project(1251), true);
		expect(module.source).toContain('\u0418\u0442\u043e\u0433 = \u0427\u0438\u0441\u043b\u043e * 3');
	});

	it("reads the machine's page from Windows, and has none elsewhere", () => {
		const page = realPlatform.ansiCodePage?.();
		if (process.platform === 'win32') {
			expect(Number.isInteger(page) && page! > 0).toBe(true);
		} else {
			expect(page).toBeUndefined();
		}
	});

	it("prefers the setting to the machine's page", () => {
		setHostPlatform(machineWithCodePage(932));
		setVb6CodePage(1251);
		expect(vb6CodePage()).toBe(1251);
	});

	it('falls back to Windows-1252 where there is no page, or one it cannot convert', () => {
		setHostPlatform(machineWithCodePage(undefined));
		expect(vb6CodePage()).toBe(1252);
		setHostPlatform(machineWithCodePage(1361));
		expect(vb6CodePage()).toBe(1252);
	});

	it('writes an edit back in the same page, keeping its Cyrillic', () => {
		setVb6CodePage(1251);
		const vbp = project(1251);
		writeVb6Module(vbp, '\u041c\u043e\u0434\u0443\u043b\u044c1', 'Option Explicit\r\nPublic Sub \u041f\u0440\u0438\u0432\u0435\u0442()\r\nEnd Sub\r\n');
		const bytes = fs.readFileSync(path.join(path.dirname(vbp), '\u041c\u043e\u0434\u0443\u043b\u044c1.bas'));
		expect(decodeCodePage(bytes, 1251)).toContain('Public Sub \u041f\u0440\u0438\u0432\u0435\u0442()');
		expect(bytes.includes(0x3f)).toBe(false); // no '?' stands in for a letter
	});

	it('reads a file with a UTF-8 byte-order mark as UTF-8, whatever the page', () => {
		setVb6CodePage(1251);
		const bom = Buffer.concat([Buffer.from([0xef, 0xbb, 0xbf]), Buffer.from(MODULE, 'utf8')]);
		const [module] = readVb6Modules(project(1251, { 'Project1.vbp': MANIFEST, '\u041c\u043e\u0434\u0443\u043b\u044c1.bas': bom }), true);
		expect(module.source).toContain('Dim \u0427\u0438\u0441\u043b\u043e As Long');
	});

	it('rereads a project when the setting changes', () => {
		const vbp = project(1251);
		setHostPlatform(machineWithCodePage(1252));
		const garbled = readVb6Modules(vbp, false);
		expect(garbled.map((module) => module.name)).not.toContain('\u041c\u043e\u0434\u0443\u043b\u044c1');
		setVb6CodePage(1251);
		expect(readVb6Modules(vbp, false).map((module) => module.name)).toEqual(['\u041c\u043e\u0434\u0443\u043b\u044c1']);
	});
});
