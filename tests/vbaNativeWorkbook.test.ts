import { describe, expect, it } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { Cfb } from '../src/vba/cfb';
import { compress, decompress } from '../src/vba/ovba';
import { VbaProject, detectSignature } from '../src/vba/vbaProject';
import { XlsxWorkbook } from '../src/vba/xlsx';
import * as svc from '../src/vba/workbookService';

// The bundled blank workbook is a real macro-enabled package (OOXML ZIP + a
// vbaProject.bin CFB), so these run the whole native stack without needing any
// workbook from outside the repo.
const TEMPLATE = path.join(__dirname, '..', 'assets', 'templates', 'blank.xlsm');

function tempCopy(): string {
	const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'xlide-native-'));
	const target = path.join(dir, 'Book.xlsm');
	fs.copyFileSync(TEMPLATE, target);
	return target;
}

describe('MS-OVBA compression', () => {
	it('round-trips text, empty input and highly repetitive input', () => {
		for (const sample of [
			'',
			'Attribute VB_Name = "M"\r\nPublic Sub A()\r\nEnd Sub\r\n',
			'x'.repeat(20000),
			Array.from({ length: 400 }, (_, i) => `Public Sub S${i}()\r\n    Debug.Print ${i}\r\nEnd Sub\r\n`).join(''),
		]) {
			const buf = Buffer.from(sample, 'latin1');
			expect(decompress(compress(buf)).equals(buf), `len ${sample.length}`).toBe(true);
		}
	});

	it('round-trips binary data across chunk boundaries', () => {
		const buf = Buffer.alloc(4096 * 3 + 17);
		for (let i = 0; i < buf.length; i++) { buf[i] = (i * 31 + (i >> 8)) & 0xff; }
		expect(decompress(compress(buf)).equals(buf)).toBe(true);
	});

	it('rejects a stream without the 0x01 signature byte', () => {
		expect(() => decompress(Buffer.from([0x02, 0x00]))).toThrow();
	});
});

describe('CFB container', () => {
	it('round-trips a real vbaProject.bin losslessly and idempotently', () => {
		const vba = XlsxWorkbook.fromBuffer(fs.readFileSync(TEMPLATE)).readVbaProject();
		const cfb = Cfb.fromBytes(vba);
		const streams = cfb.listStreamsInStorage('VBA');
		expect(streams.length).toBeGreaterThan(0);
		const before = streams.map((n) => cfb.getStreamInStorage('VBA', n).toString('base64'));

		const rebuilt = cfb.toBytes();
		const second = Cfb.fromBytes(rebuilt);
		expect(second.listStreamsInStorage('VBA').sort()).toEqual([...streams].sort());
		for (const [i, name] of streams.entries()) {
			expect(second.getStreamInStorage('VBA', name).toString('base64'), name).toBe(before[i]);
		}
		// Serializing an unchanged container twice must be byte-stable.
		expect(second.toBytes().equals(rebuilt)).toBe(true);
	});

	it('adds, renames and removes streams', () => {
		const cfb = Cfb.fromBytes(XlsxWorkbook.fromBuffer(fs.readFileSync(TEMPLATE)).readVbaProject());
		cfb.addStreamToStorage('VBA', 'ZZTemp', Buffer.from('hello'));
		expect(cfb.getStreamInStorage('VBA', 'ZZTemp').toString()).toBe('hello');
		cfb.renameStreamInStorage('VBA', 'ZZTemp', 'ZZRenamed');
		const reloaded = Cfb.fromBytes(cfb.toBytes());
		expect(reloaded.getStreamInStorage('VBA', 'ZZRenamed').toString()).toBe('hello');
		expect(reloaded.hasStreamInStorage('VBA', 'ZZTemp')).toBe(false);
		reloaded.removeStreamInStorage('VBA', 'ZZRenamed');
		expect(Cfb.fromBytes(reloaded.toBytes()).hasStreamInStorage('VBA', 'ZZRenamed')).toBe(false);
	});
});

describe('VBA project', () => {
	it('parses modules and survives a no-op save', () => {
		const cfb = Cfb.fromBytes(XlsxWorkbook.fromBuffer(fs.readFileSync(TEMPLATE)).readVbaProject());
		const project = VbaProject.parse(cfb);
		expect(project.modules.length).toBeGreaterThan(0);
		expect(project.codePage).toBeGreaterThan(0);
		const before = project.modules.map((m) => ({ name: m.name, source: m.source }));
		project.save(cfb);
		const reparsed = VbaProject.parse(Cfb.fromBytes(cfb.toBytes()));
		expect(reparsed.modules.map((m) => ({ name: m.name, source: m.source }))).toEqual(before);
	});

	it('never leaves __SRP_* performance-cache streams behind', () => {
		// Stale p-code caches make Excel follow compiled state that no longer
		// matches the module set, which it does not survive.
		const cfb = Cfb.fromBytes(XlsxWorkbook.fromBuffer(fs.readFileSync(TEMPLATE)).readVbaProject());
		cfb.addStreamToStorage('VBA', '__SRP_0', Buffer.alloc(64, 7));
		const project = VbaProject.parse(cfb);
		project.save(cfb);
		const reloaded = Cfb.fromBytes(cfb.toBytes());
		expect(reloaded.listStreamsInStorage('VBA').filter((n) => n.startsWith('__SRP_'))).toEqual([]);
	});

	it('reports signature state', () => {
		const cfb = Cfb.fromBytes(XlsxWorkbook.fromBuffer(fs.readFileSync(TEMPLATE)).readVbaProject());
		expect(detectSignature(cfb).present).toBe(false);
	});
});

describe('native workbook service', () => {
	it('lists modules with their types', () => {
		const modules = svc.listModules(TEMPLATE);
		expect(modules.length).toBeGreaterThan(0);
		for (const module of modules) {
			expect(['standard', 'class', 'document', 'userform']).toContain(module.type);
		}
	});

	it('writes, renames and deletes modules, leaving the rest untouched', () => {
		const target = tempCopy();
		const before = svc.readModules(target, true);

		svc.writeModule(target, 'NativeAdded', 'Public Sub Hi()\r\n    Debug.Print 1\r\nEnd Sub\r\n');
		expect(svc.readModule(target, 'NativeAdded', false).source).toContain('Public Sub Hi()');
		expect(svc.readModule(target, 'NativeAdded', true).source).toContain('Attribute VB_Name = "NativeAdded"');

		svc.renameModule(target, 'NativeAdded', 'NativeRenamed');
		expect(svc.listModules(target).map((m) => m.name)).toContain('NativeRenamed');
		expect(svc.readModule(target, 'NativeRenamed', true).source)
			.toContain('Attribute VB_Name = "NativeRenamed"');

		svc.writeModule(target, 'NativeDoomed', 'Public Sub Bye()\r\nEnd Sub\r\n');
		svc.deleteModule(target, 'NativeDoomed');
		expect(svc.listModules(target).map((m) => m.name)).not.toContain('NativeDoomed');

		// Every pre-existing module keeps its exact source.
		const after = svc.readModules(target, true);
		for (const module of before) {
			expect(after.find((m) => m.name === module.name)?.source, module.name).toBe(module.source);
		}
		expect(svc.validateWorkbook(target).issues).toEqual([]);
	});

	it('creates class modules with a class header', () => {
		const target = tempCopy();
		svc.writeModule(target, 'NativeClass', 'Public Sub Go()\r\nEnd Sub\r\n', 'class');
		expect(svc.listModules(target).find((m) => m.name === 'NativeClass')?.type).toBe('class');
		expect(svc.readModule(target, 'NativeClass', true).source).toContain('VERSION 1.0 CLASS');
	});

	it('lists procedures with body-relative line numbers', () => {
		const target = tempCopy();
		svc.writeModule(target, 'Procs', 'Public Sub First()\r\nEnd Sub\r\n\r\nPrivate Function Second() As Long\r\nEnd Function\r\n');
		expect(svc.listSubs(target, 'Procs')).toEqual([
			{ name: 'First', kind: 'Sub', line: 1 },
			{ name: 'Second', kind: 'Function', line: 4 },
		]);
	});

	it('reads and writes cells while preserving the VBA project', () => {
		const target = tempCopy();
		const modulesBefore = svc.readModules(target, true);
		const sheet = svc.listSheets(target).sheets[0].name;
		svc.writeCells(target, sheet, 'B2', [['text', 12.5, true], [null, 'x', 7]]);
		expect(svc.readCells(target, sheet, 'B2:D3').data).toEqual([
			['text', 12.5, true],
			[null, 'x', 7],
		]);
		expect(svc.readModules(target, true)).toEqual(modulesBefore);
	});

	it('reports workbook info in one pass', () => {
		const info = svc.getWorkbookInfo(TEMPLATE);
		expect(info.sheets.length).toBeGreaterThan(0);
		expect(Array.isArray(info.namedRanges)).toBe(true);
		expect(info.modules.length).toBeGreaterThan(0);
		expect(typeof info.isPasswordProtected).toBe('boolean');
		expect(typeof info.isSigned).toBe('boolean');
	});

	it('splits and rejoins the hidden attribute header', () => {
		const full = 'Attribute VB_Name = "M"\r\nAttribute VB_Exposed = False\r\n\r\nPublic Sub A()\r\nEnd Sub\r\n';
		const { header, body } = svc.splitVbaSource(full);
		expect(header).toBe('Attribute VB_Name = "M"\r\nAttribute VB_Exposed = False\r\n');
		expect(body).toBe('Public Sub A()\r\nEnd Sub\r\n');
		expect(svc.joinVbaSource(header, body)).toBe(
			'Attribute VB_Name = "M"\r\nAttribute VB_Exposed = False\r\nPublic Sub A()\r\nEnd Sub\r\n',
		);
	});

	it('classifies module and document types from source', () => {
		expect(svc.classifyModuleType('ThisWorkbook', '')).toBe('document');
		expect(svc.classifyModuleType('Sheet1', '')).toBe('document');
		expect(svc.classifyModuleType('Module1', '')).toBe('standard');
		expect(svc.classifyDocumentType('ThisWorkbook', '')).toBe('workbook');
		expect(svc.classifyDocumentType('Sheet3', '')).toBe('worksheet');
		// A UserForm carries two GUIDs in VB_Base (type library + instance).
		expect(svc.classifyModuleType(
			'UserForm1',
			'Attribute VB_Base = "0{C62A69F0-16DC-11CE-9E98-00AA00574A4F}{3A2C0F9E-1234-11CE-9E98-00AA00574A4F}"',
		)).toBe('userform');
		// A plain class module carries exactly one, and is not a document CLSID.
		expect(svc.classifyModuleType(
			'Class1',
			'Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"',
		)).toBe('standard');
	});
});
