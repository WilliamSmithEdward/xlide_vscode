import { describe, expect, it } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { decodeCodePage, encodeCodePage } from '../src/vba/codePages';
import { Cfb } from '../src/vba/cfb';
import { compress, decompress } from '../src/vba/ovba';
import { VbaProject } from '../src/vba/vbaProject';
import { XlsxWorkbook } from '../src/vba/xlsx';
import * as svc from '../src/vba/workbookService';

// The regression from issue #6: cp1251 bytes for 'Модуль' read back as
// 'Ìîäóëü' because every non-1252 page fell back to latin1.
const CYRILLIC = 'Модуль: mdTest';
const CYRILLIC_CP1251 = Buffer.from([
	0xcc, 0xee, 0xe4, 0xf3, 0xeb, 0xfc, 0x3a, 0x20, 0x6d, 0x64, 0x54, 0x65, 0x73, 0x74,
]);

describe('code-page conversion', () => {
	it('decodes and encodes cp1251 (the issue #6 regression)', () => {
		expect(decodeCodePage(CYRILLIC_CP1251, 1251)).toBe(CYRILLIC);
		expect(decodeCodePage(CYRILLIC_CP1251, 1251)).not.toBe('Ìîäóëü: mdTest');
		expect(encodeCodePage(CYRILLIC, 1251).equals(CYRILLIC_CP1251)).toBe(true);
	});

	it('round-trips other single-byte languages', () => {
		const samples: Array<[number, string]> = [
			[1250, 'Zażółć gęślą jaźń'],          // Polish
			[1253, 'Δοκιμή ενότητας'],             // Greek
			[1254, 'Değişken tanımı'],             // Turkish
			[1251, 'Проверка модуля'],             // Russian
			[874, 'ทดสอบโมดูล'],                    // Thai
		];
		for (const [codePage, text] of samples) {
			const bytes = encodeCodePage(text, codePage);
			expect(decodeCodePage(bytes, codePage), `cp${codePage}`).toBe(text);
			// A meaningful encoding, not a wall of '?' substitutions.
			expect([...bytes].filter((b) => b === 0x3f).length, `cp${codePage}`).toBe(0);
		}
	});

	it('round-trips double-byte CJK pages', () => {
		const samples: Array<[number, string]> = [
			[932, 'モジュールのテスト'],   // Japanese Shift-JIS
			[936, '模块测试'],             // Simplified Chinese GBK
			[949, '모듈 테스트'],          // Korean
			[950, '模組測試'],             // Traditional Chinese Big5
		];
		for (const [codePage, text] of samples) {
			const bytes = encodeCodePage(text, codePage);
			expect(decodeCodePage(bytes, codePage), `cp${codePage}`).toBe(text);
			expect([...bytes].filter((b) => b === 0x3f).length, `cp${codePage}`).toBe(0);
		}
	});

	it('keeps cp1252 byte-exact with the original table for all 256 bytes', () => {
		for (let byte = 0; byte <= 0xff; byte++) {
			const decoded = decodeCodePage(Buffer.from([byte]), 1252);
			expect(decoded.length, `byte ${byte}`).toBe(1);
			const reencoded = encodeCodePage(decoded, 1252);
			expect(reencoded[0], `byte ${byte}`).toBe(byte);
		}
	});

	it('substitutes ? for characters a page cannot represent', () => {
		expect(encodeCodePage('Модуль', 1252).every((b) => b === 0x3f)).toBe(true);
		expect(decodeCodePage(encodeCodePage('日本語', 1251), 1251)).toBe('???');
	});

	it('falls back to cp1252 behavior for unknown code pages', () => {
		const text = 'plain ASCII stays plain';
		expect(decodeCodePage(Buffer.from(text, 'latin1'), 42_424)).toBe(text);
		expect(encodeCodePage(text, 42_424).toString('latin1')).toBe(text);
	});
});

describe('cp1251 workbook end to end', () => {
	const TEMPLATE = path.join(__dirname, '..', 'assets', 'templates', 'blank.xlsm');
	const REC_PROJECTCODEPAGE = 0x0003;

	/** Copy the template and rewrite its PROJECTCODEPAGE dir record to cp1251. */
	function makeCp1251Workbook(): string {
		const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'xlide-cp1251-'));
		const target = path.join(dir, 'Russian.xlsm');
		const xlsx = XlsxWorkbook.fromBuffer(fs.readFileSync(TEMPLATE));
		const cfb = Cfb.fromBytes(xlsx.readVbaProject());
		const dirRaw = Buffer.from(decompress(cfb.getStreamInStorage('VBA', 'dir'), 'VBA/dir'));

		// Walk [id u16][size u32][data] records to the code-page record.
		let pos = 0;
		let patched = false;
		while (pos + 6 <= dirRaw.length) {
			const id = dirRaw.readUInt16LE(pos);
			const size = dirRaw.readUInt32LE(pos + 2);
			if (id === REC_PROJECTCODEPAGE && size >= 2) {
				dirRaw.writeUInt16LE(1251, pos + 6);
				patched = true;
				break;
			}
			pos += 6 + size;
		}
		expect(patched).toBe(true);

		cfb.writeStreamInStorage('VBA', 'dir', compress(dirRaw));
		xlsx.writeVbaProject(cfb.toBytes());
		fs.writeFileSync(target, xlsx.toBytes());
		return target;
	}

	it('writes and reads Cyrillic source through the real engine', () => {
		const file = makeCp1251Workbook();
		const source = [
			"' Модуль: mdTest",
			'Public Sub Проверка()',
			'    Debug.Print "Русский текст"',
			'End Sub',
			'',
		].join('\r\n');

		svc.writeModule(file, 'mdTest', source, 'standard');
		expect(svc.readModule(file, 'mdTest', false).source).toBe(source);

		// The stored bytes are genuinely cp1251: the compressed module stream
		// decompresses to bytes containing 'Модуль' as CC EE E4 F3 EB FC.
		const cfb = Cfb.fromBytes(XlsxWorkbook.fromBuffer(fs.readFileSync(file)).readVbaProject());
		const stream = cfb.getStreamInStorage('VBA', 'mdTest');
		const bytes = decompress(stream, 'VBA/mdTest');
		expect(bytes.includes(Buffer.from([0xcc, 0xee, 0xe4, 0xf3, 0xeb, 0xfc]))).toBe(true);

		// And a fresh listModules still classifies + validates cleanly.
		expect(svc.listModules(file).map((m) => m.name)).toContain('mdTest');
		expect(svc.validateWorkbook(file).issues).toEqual([]);
	});
});

describe('non-ASCII module and procedure names (the wider bug class)', () => {
	const TEMPLATE = path.join(__dirname, '..', 'assets', 'templates', 'blank.xlsm');
	const REC_PROJECTCODEPAGE = 0x0003;

	function makeCp1251Workbook(): string {
		const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'xlide-cp1251n-'));
		const target = path.join(dir, 'Russian.xlsm');
		const xlsx = XlsxWorkbook.fromBuffer(fs.readFileSync(TEMPLATE));
		const cfb = Cfb.fromBytes(xlsx.readVbaProject());
		const dirRaw = Buffer.from(decompress(cfb.getStreamInStorage('VBA', 'dir'), 'VBA/dir'));
		let pos = 0;
		while (pos + 6 <= dirRaw.length) {
			const id = dirRaw.readUInt16LE(pos);
			const size = dirRaw.readUInt32LE(pos + 2);
			if (id === REC_PROJECTCODEPAGE && size >= 2) {
				dirRaw.writeUInt16LE(1251, pos + 6);
				break;
			}
			pos += 6 + size;
		}
		cfb.writeStreamInStorage('VBA', 'dir', compress(dirRaw));
		xlsx.writeVbaProject(cfb.toBytes());
		fs.writeFileSync(target, xlsx.toBytes());
		return target;
	}

	it('supports the full lifecycle of a Cyrillic-named module', () => {
		const file = makeCp1251Workbook();
		const source = [
			'Public Sub Проверка()',
			'    Dim счетчик As Long',
			'    счетчик = 1',
			'End Sub',
			'',
		].join('\r\n');

		// Create under a Cyrillic name, then rename to another Cyrillic name.
		svc.writeModule(file, 'МодульТест', source, 'standard');
		expect(svc.listModules(file).map((m) => m.name)).toContain('МодульТест');
		svc.renameModule(file, 'МодульТест', 'НовыйМодуль');

		const names = svc.listModules(file).map((m) => m.name);
		expect(names).toContain('НовыйМодуль');
		expect(names).not.toContain('МодульТест');
		expect(svc.readModule(file, 'НовыйМодуль', true).source)
			.toContain('Attribute VB_Name = "НовыйМодуль"');

		// The dir stream's Unicode name records agree with the ANSI ones.
		const cfb = Cfb.fromBytes(XlsxWorkbook.fromBuffer(fs.readFileSync(file)).readVbaProject());
		const project = VbaProject.parse(cfb);
		const module = project.getModule('НовыйМодуль');
		expect(module?.nameUnicode).toBe('НовыйМодуль');
		expect(module?.streamNameUnicode).toBe('НовыйМодуль');

		// Cyrillic procedure names appear in the tree's procedure list.
		const subs = svc.listSubs(file, 'НовыйМодуль');
		expect(subs.map((s) => s.name)).toContain('Проверка');

		expect(svc.validateWorkbook(file).issues).toEqual([]);
	});
});

describe('analyzer on non-ASCII identifiers', () => {
	it('reports nothing on clean Cyrillic and Greek code', async () => {
		// The structural engine used [A-Za-z_]\w* for procedure names, so
		// `Sub Проверка()` was invisible to it and its End Sub reported as
		// unmatched - a false error on every Russian user's code.
		const { analyzeVbaModuleSource } = await import('../src/vbaModuleAnalysis');
		const samples = [
			['Option Explicit', 'Public Sub Проверка()', '    Dim счетчик As Long',
				'    счетчик = 42', 'End Sub', ''].join('\r\n'),
			['Option Explicit', 'Public Function Δοκιμή() As Long',
				'    Δοκιμή = 1', 'End Function', ''].join('\r\n'),
		];
		for (const source of samples) {
			const result = analyzeVbaModuleSource({
				source, moduleName: 'M', moduleType: 'standard', moduleKind: 'standard',
			});
			expect(result.diagnostics, source.slice(0, 40)).toEqual([]);
		}
	});
});
