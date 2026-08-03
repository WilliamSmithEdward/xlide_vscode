import { describe, expect, it } from 'vitest';
import { TextDecoder } from 'util';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { codePageLabel, decodeCodePage, encodeCodePage, supportedCodePages } from '../src/vba/codePages';
import { Cfb } from '../src/vba/cfb';
import { compress, decompress } from '../src/vba/ovba';
import { XlsxWorkbook } from '../src/vba/xlsx';
import * as svc from '../src/vba/workbookService';

// One native-language sample per supported code page. Each runs the REAL
// engine end to end: a workbook whose PROJECTCODEPAGE is that page, a module
// written and read back through workbookService, and a clean validate. This is
// the CI gate for the issue-#6 bug class: an ASCII-only assumption anywhere in
// the pipeline fails one of these rows instead of reaching a user.
const LANGUAGE_MATRIX: Array<[number, string, string]> = [
	[874, 'Thai', 'ทดสอบภาษาไทย'],
	[932, 'Japanese (Shift-JIS)', 'テスト用モジュール'],
	[936, 'Chinese, Simplified (GBK)', '中文测试模块'],
	[949, 'Korean (EUC-KR)', '한국어 테스트'],
	[950, 'Chinese, Traditional (Big5)', '繁體中文測試'],
	[1250, 'Central European', 'Příliš žluťoučký kůň Zażółć gęślą'],
	[1251, 'Cyrillic', 'Проверка русского текста'],
	[1252, 'Western European', 'déjà vu € œuvre Straße'],
	[1253, 'Greek', 'Δοκιμή ελληνικού κειμένου'],
	[1254, 'Turkish', 'Türkçe deneme ğüşiöç İı'],
	[1255, 'Hebrew', 'בדיקת עברית'],
	[1256, 'Arabic', 'اختبار العربية'],
	[1257, 'Baltic', 'Lietuviškas tekstas ąčęėįšųū'],
	[1258, 'Vietnamese', 'Tiếng Việt thử nghiệm'],
	[10000, 'Mac Roman', 'déjà vu café œuvre'],
	[20866, 'Russian (KOI8-R)', 'Тест КОИ-8'],
	[21866, 'Ukrainian (KOI8-U)', 'Тест української ґї'],
	[28592, 'ISO-8859-2', 'Zažil žluťoučký Zażółć'],
	[28595, 'ISO-8859-5', 'Проверка ИСО'],
	[65001, 'UTF-8', 'любой текст 中文 déjà ทดสอบ'],
];

const REC_PROJECTCODEPAGE = 0x0003;
const TEMPLATE = path.join(__dirname, '..', 'assets', 'templates', 'blank.xlsm');

function makeWorkbookWithCodePage(codePage: number): string {
	const dir = fs.mkdtempSync(path.join(os.tmpdir(), `xlide-cp${codePage}-`));
	const target = path.join(dir, 'Book.xlsm');
	const xlsx = XlsxWorkbook.fromBuffer(fs.readFileSync(TEMPLATE));
	const cfb = Cfb.fromBytes(xlsx.readVbaProject());
	const dirRaw = Buffer.from(decompress(cfb.getStreamInStorage('VBA', 'dir'), 'VBA/dir'));
	let pos = 0;
	let patched = false;
	while (pos + 6 <= dirRaw.length) {
		const id = dirRaw.readUInt16LE(pos);
		const size = dirRaw.readUInt32LE(pos + 2);
		if (id === REC_PROJECTCODEPAGE && size >= 2) {
			dirRaw.writeUInt16LE(codePage, pos + 6);
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

/** cp1258 legitimately decodes to combining sequences; compare canonicalized. */
function nfc(text: string): string {
	return text.normalize('NFC');
}

describe('CI language matrix', () => {
	it('has a live decoder for every supported code page (full-ICU guard)', () => {
		// At runtime a missing decoder degrades gracefully to cp1252; in CI that
		// silence would hide exactly the bug class this matrix exists to catch,
		// so the build must prove the good path is actually exercised.
		const pages = supportedCodePages();
		expect(pages.length).toBeGreaterThanOrEqual(LANGUAGE_MATRIX.length);
		for (const codePage of pages) {
			const label = codePageLabel(codePage);
			expect(label, `cp${codePage} has a label`).toBeTruthy();
			expect(() => new TextDecoder(label!), `cp${codePage} (${label})`).not.toThrow();
		}
	});

	for (const [codePage, language, sample] of LANGUAGE_MATRIX) {
		it(`round-trips ${language} (cp${codePage}) through the real engine`, () => {
			// Codec level first: a meaningful encoding, not '?' substitutions.
			const bytes = encodeCodePage(sample, codePage);
			expect([...bytes].filter((b) => b === 0x3f).length, 'no ? substitutions').toBe(0);
			expect(nfc(decodeCodePage(bytes, codePage))).toBe(nfc(sample));

			// Then the whole engine: write, read, list, validate.
			const file = makeWorkbookWithCodePage(codePage);
			const source = [
				`' ${sample}`,
				'Public Sub LanguageProbe()',
				`    Debug.Print "${sample}"`,
				'End Sub',
				'',
			].join('\r\n');
			svc.writeModule(file, 'LangModule', source, 'standard');
			expect(nfc(svc.readModule(file, 'LangModule', false).source)).toBe(nfc(source));
			expect(svc.listModules(file).map((m) => m.name)).toContain('LangModule');
			expect(svc.validateWorkbook(file).issues).toEqual([]);
		});
	}

	it('supports native-language MODULE NAMES end to end for major scripts', () => {
		const cases: Array<[number, string, string]> = [
			[1251, 'МодульТест', 'НовыйМодуль'],
			[932, 'モジュール', 'テスト部品'],
			[936, '测试模块', '新模块'],
		];
		for (const [codePage, name, renamed] of cases) {
			const file = makeWorkbookWithCodePage(codePage);
			svc.writeModule(file, name, 'Public Sub P()\r\nEnd Sub\r\n', 'standard');
			expect(svc.listModules(file).map((m) => m.name), `cp${codePage}`).toContain(name);
			svc.renameModule(file, name, renamed);
			expect(svc.listModules(file).map((m) => m.name), `cp${codePage}`).toContain(renamed);
			expect(svc.readModule(file, renamed, true).source)
				.toContain(`Attribute VB_Name = "${renamed}"`);
			expect(svc.validateWorkbook(file).issues, `cp${codePage}`).toEqual([]);
		}
	});
});
