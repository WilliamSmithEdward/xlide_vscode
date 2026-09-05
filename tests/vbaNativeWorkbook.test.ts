import { describe, expect, it } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { Cfb } from '../src/vba/cfb';
import { compress, decompress } from '../src/vba/ovba';
import { VbaProject, detectSignature } from '../src/vba/vbaProject';
import { ZipArchive } from '../src/vba/zip';
import { XlsxWorkbook } from '../src/vba/xlsx';
import * as svc from '../src/vba/projectService';

// The bundled blank workbook is a real macro-enabled package (OOXML ZIP + a
// vbaProject.bin CFB), so these run the whole native stack without needing any
// workbook from outside the repo.
const TEMPLATE_DIR = path.join(__dirname, '..', 'assets', 'templates');
const TEMPLATE = path.join(TEMPLATE_DIR, 'blank.xlsm');

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

	it('returns an exact prefix of the full output when capped with maxBytes', () => {
		// Lazy module loading inflates only a module's header, so a capped
		// decompression must agree byte for byte with the full one - never a
		// chunk resynchronized against a different starting point.
		const samples = [
			Buffer.from('Attribute VB_Name = "M"\r\n' + 'Public Sub A()\r\nEnd Sub\r\n'.repeat(600), 'latin1'),
			Buffer.from('ab'.repeat(30000), 'latin1'),
			(() => {
				const buf = Buffer.alloc(4096 * 5 + 123);
				for (let i = 0; i < buf.length; i++) { buf[i] = (i * 17 + (i >> 7)) & 0xff; }
				return buf;
			})(),
		];
		for (const raw of samples) {
			const packed = compress(raw);
			const full = decompress(packed);
			expect(full.equals(raw)).toBe(true);
			for (const cap of [1, 10, 4096, 4097, 8192, 20000]) {
				const part = decompress(packed, '<test>', cap);
				expect(full.subarray(0, part.length).equals(part), `cap ${cap}`).toBe(true);
				// A cap must never stop short of what it asked for while output remains.
				expect(part.length, `cap ${cap}`).toBeGreaterThanOrEqual(Math.min(cap, full.length));
			}
		}
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
		expect(svc.validateProject(target).issues).toEqual([]);
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
		const info = svc.getProjectInfo(TEMPLATE);
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

describe('lazy module sources', () => {
	it('exposes sourceHeader as an exact prefix of source, including past a chunk', () => {
		const file = tempCopy();
		// Well over one 4096-byte decompression chunk, so the header genuinely
		// stops short of the body rather than covering it by accident.
		const long = Array.from(
			{ length: 500 },
			(_, i) => `Public Sub Generated${i}()\r\n    Debug.Print ${i}\r\nEnd Sub\r\n`,
		).join('');
		svc.writeModule(file, 'BigModule', long, 'standard');

		const cfb = Cfb.fromBytes(XlsxWorkbook.fromBuffer(fs.readFileSync(file)).readVbaProject());
		const headerFirst = VbaProject.parse(cfb).getModule('BigModule');
		const sourceFirst = VbaProject.parse(cfb).getModule('BigModule');
		expect(headerFirst && sourceFirst).toBeTruthy();

		const header = headerFirst!.sourceHeader;
		const source = sourceFirst!.source;
		expect(source.length).toBeGreaterThan(header.length);
		expect(source.startsWith(header)).toBe(true);
		// The attribute header is what callers classify from, so it has to be
		// inside the prefix we bother to inflate.
		expect(header).toContain('Attribute VB_Name');

		// Reading source first must not change what sourceHeader reports.
		expect(sourceFirst!.sourceHeader).toBe(source);
	});

	it('keeps sourceHeader in step with an assigned source', () => {
		const file = tempCopy();
		svc.writeModule(file, 'Mod1', 'Public Sub A()\r\nEnd Sub\r\n', 'standard');
		const cfb = Cfb.fromBytes(XlsxWorkbook.fromBuffer(fs.readFileSync(file)).readVbaProject());
		const project = VbaProject.parse(cfb);

		project.setModuleSource('Mod1', 'Attribute VB_Name = "Mod1"\r\nPublic Sub B()\r\nEnd Sub\r\n');
		expect(project.getModule('Mod1')!.sourceHeader).toContain('Public Sub B');

		// Added modules take the same path rather than freezing at birth.
		const added = project.addModule('Mod2', 'Attribute VB_Name = "Mod2"\r\n', 'standard');
		expect(added.sourceHeader).toBe(added.source);
		project.setModuleSource('Mod2', 'Attribute VB_Name = "Mod2"\r\nPublic Sub C()\r\nEnd Sub\r\n');
		expect(project.getModule('Mod2')!.sourceHeader).toContain('Public Sub C');
	});

	it('classifies module types from the header alone', () => {
		// listModules must report the same types whether or not any body was
		// inflated, since classification now reads only the header prefix.
		const file = tempCopy();
		svc.writeModule(file, 'Helper', 'Public Sub H()\r\nEnd Sub\r\n', 'standard');
		svc.writeModule(file, 'Widget', 'Public Sub W()\r\nEnd Sub\r\n', 'class');

		const listed = svc.listModules(file);
		const byName = new Map(listed.map((m) => [m.name, m]));
		expect(byName.get('Helper')?.type).toBe('standard');
		expect(byName.get('Widget')?.type).toBe('class');
		expect(byName.get('ThisWorkbook')?.type).toBe('document');
		expect(byName.get('ThisWorkbook')?.documentType).toBe('workbook');

		// Same answers from a full read, which materializes every source.
		const full = svc.readModules(file, true);
		for (const entry of full) {
			expect(byName.get(entry.name)?.type, entry.name).toBe(entry.type);
			expect(byName.get(entry.name)?.documentType, entry.name).toBe(entry.documentType);
		}
	});

	// The explorer's folder layout, issue #66.
	it('reports the @Folder annotation from the header prefix, at no extra cost', () => {
		const file = tempCopy();
		svc.writeModule(file, 'Ledger', '\'@Folder("Accounts.Ledger")\r\nPublic Sub L()\r\nEnd Sub\r\n', 'standard');
		svc.writeModule(file, 'Loose', 'Public Sub N()\r\nEnd Sub\r\n', 'standard');

		const byName = new Map(svc.listModules(file).map((m) => [m.name, m]));
		expect(byName.get('Ledger')?.folder).toBe('Accounts.Ledger');
		expect(byName.get('Loose')?.folder).toBeUndefined();

		// A full read, which inflates every source, reads the same prefix and
		// must not answer differently: one rule, one answer, whichever call the
		// tree or the analyzer makes.
		for (const entry of svc.readModules(file, true)) {
			expect(byName.get(entry.name)?.folder, entry.name).toBe(entry.folder);
		}
	});

	// Pinned deliberately: reaching past the prefix for this would inflate
	// every module on every listing, which doubled a cold listModules over the
	// test corpus and changed no answer on any real module. See folderOfModule.
	it('does not go looking past the header prefix for a late annotation', () => {
		const file = tempCopy();
		// Well past the 4096 bytes of source the header prefix inflates.
		const longDeclarations = Array.from(
			{ length: 300 },
			(_, i) => `Public Const Constant${i} As Long = ${i}`,
		).join('\r\n');
		svc.writeModule(file, 'Deep', `${longDeclarations}\r\n'@Folder("Accounts.Deep")\r\nPublic Sub D()\r\nEnd Sub\r\n`, 'standard');

		const deep = svc.listModules(file).find((m) => m.name === 'Deep');
		expect(deep?.folder).toBeUndefined();
		// The module lands where an unannotated one lands - the project root -
		// so the fix is visible: move the annotation to the top of the module.
		expect(svc.readModules(file, true).find((m) => m.name === 'Deep')?.folder).toBeUndefined();
	});
});

describe('packaged assets', () => {
	// createProject copies these files out of the installed extension, so each
	// has to be inside the .vsix. The broad `**/*.xls?` rules in .vscodeignore
	// keep test workbooks out of the package and silently took the .xlsm one
	// with it once; the failure only shows up as New Macro-Enabled File
	// throwing ENOENT on a real install, which no unit test would otherwise
	// catch. `.xlsb` and `.xlam` shipped excluded with no re-include at all
	// until the templates existed to need them.
	// Both ignore files exclude workbook extensions broadly. .vscodeignore
	// decides what reaches the .vsix; .gitignore decides what reaches the repo
	// at all, so a missed re-include there means the template is absent from a
	// fresh clone and CI - the same failure one layer earlier.
	it.each([
		['.vscodeignore', '**/*.', 'xlsm', 'blank.xlsm'],
		['.vscodeignore', '**/*.', 'xlsb', 'blank.xlsb'],
		['.vscodeignore', '**/*.', 'xlam', 'blank.xlam'],
		['.gitignore', '*.', 'xlsm', 'blank.xlsm'],
		['.gitignore', '*.', 'xlsb', 'blank.xlsb'],
		['.gitignore', '*.', 'xlam', 'blank.xlam'],
	])('%s re-includes the blank .%s%s template', (ignoreFile, prefix, ext, file) => {
		expect(fs.existsSync(path.join(TEMPLATE_DIR, file))).toBe(true);

		const ignore = fs.readFileSync(
			path.join(__dirname, '..', ignoreFile), 'utf8',
		).split(/\r?\n/).map((line) => line.trim());

		const excluded = ignore.indexOf(`${prefix}${ext}`);
		const reincluded = ignore.indexOf(`!assets/templates/${file}`);
		expect(excluded, `${ignoreFile} no longer excludes ${prefix}${ext}; re-check this guard`).toBeGreaterThan(-1);
		expect(reincluded, `assets/templates/${file} must be re-included after the ${prefix}${ext} exclusion`)
			.toBeGreaterThan(excluded);
	});
});

/** The declared content type of a package's main workbook part. */
function mainContentType(file: string): string {
	const xml = ZipArchive.read(fs.readFileSync(file)).read('[Content_Types].xml').toString('utf8');
	return /ContentType="([^"]*main\+xml)"/.exec(xml)?.[1] ?? '';
}

describe('blank file templates', () => {
	// What makes Excel treat a file as an add-in (ThisWorkbook.IsAddin = True)
	// is the workbook part's content type, so an .xlam cannot be produced by
	// renaming an .xlsm - it has to come from an add-in template. Asserting the
	// marker is the part of the upstream live-Excel check (open the add-in, run
	// a macro, read IsAddin) that can be pinned offline.
	it.each([
		['blank.xlsm', 'application/vnd.ms-excel.sheet.macroEnabled.main+xml'],
		['blank.xlam', 'application/vnd.ms-excel.addin.macroEnabled.main+xml'],
	])('%s declares its format content type', (file, contentType) => {
		expect(mainContentType(path.join(TEMPLATE_DIR, file))).toBe(contentType);
	});

	it.each(['blank.xlsm', 'blank.xlsb', 'blank.xlam'])('%s carries a readable VBA project', (file) => {
		const vba = XlsxWorkbook.fromBuffer(fs.readFileSync(path.join(TEMPLATE_DIR, file))).readVbaProject();
		const project = VbaProject.parse(Cfb.fromBytes(vba));
		const names = project.modules.map((m) => m.name);
		expect(names).toContain('ThisWorkbook');
		expect(names).toContain('Module1');
	});
});

describe('createProject template dispatch', () => {
	// The template is chosen by the target's extension. Before this existed the
	// command's own save dialog offered .xlsb while createProject always copied
	// blank.xlsm, so choosing .xlsb produced an .xlsm-format file under an .xlsb
	// name - a file Excel rejects or repairs.
	it.each([
		['Book.xlsm', 'blank.xlsm'],
		['Book.xlsb', 'blank.xlsb'],
		['AddIn.xlam', 'blank.xlam'],
	])('seeds %s from %s', (targetName, templateFile) => {
		const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'xlide-create-'));
		const target = path.join(dir, targetName);
		svc.createProject(target, path.join(TEMPLATE_DIR, templateFile));

		expect(fs.readFileSync(target).equals(fs.readFileSync(path.join(TEMPLATE_DIR, templateFile)))).toBe(true);
	});

	it('produces an add-in whose VBA project accepts a module and reads it back', () => {
		// The end-to-end shape the issue cares about: create the add-in, write a
		// macro into it, read it back through the same engine an editor uses.
		const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'xlide-create-'));
		const target = path.join(dir, 'SentinelAddIn.xlam');
		svc.createProject(target, path.join(TEMPLATE_DIR, 'blank.xlam'));

		svc.writeModule(target, 'Sentinel', 'Public Sub Ping()\r\n    Debug.Print "ok"\r\nEnd Sub\r\n');

		expect(svc.readModule(target, 'Sentinel', false).source).toContain('Public Sub Ping()');
		// Writing modules must not disturb the add-in marker.
		expect(mainContentType(target)).toBe('application/vnd.ms-excel.addin.macroEnabled.main+xml');
	});
});

describe('workbook parse cache', () => {
	it('serves repeated reads from one parse, with identical results', () => {
		const file = tempCopy();
		svc.resetProjectCacheForTests();

		const first = svc.listModules(file);
		const afterFirst = svc.projectCacheStatsForTests();
		const again = svc.listModules(file);
		const subs = svc.listSubs(file, first[0].name);
		const protection = svc.getProtectionInfo(file);
		const sheets = svc.listSheets(file);
		const after = svc.projectCacheStatsForTests();

		expect(again).toEqual(first);
		expect(subs).toEqual(svc.listSubs(file, first[0].name));
		expect(protection.isSigned).toBe(false);
		expect(sheets.sheets.length).toBeGreaterThan(0);
		// One miss to build the entry; every later call in the burst hits.
		expect(afterFirst.misses).toBe(1);
		expect(after.misses).toBe(1);
		expect(after.hits).toBeGreaterThanOrEqual(4);
	});

	it('a write through the engine invalidates the cached parse', () => {
		const file = tempCopy();
		svc.resetProjectCacheForTests();

		expect(svc.listModules(file).map((m) => m.name)).not.toContain('CacheProbe');
		svc.writeModule(file, 'CacheProbe', 'Public Sub P()\r\nEnd Sub\r\n', 'standard');
		// Served fresh, not from the pre-write parse.
		expect(svc.listModules(file).map((m) => m.name)).toContain('CacheProbe');
		expect(svc.readModule(file, 'CacheProbe', false).source).toContain('Public Sub P()');
	});

	it('an out-of-band rewrite is seen via the mtime/size check', () => {
		const file = tempCopy();
		const other = tempCopy();
		svc.writeModule(other, 'External', 'Public Sub E()\r\nEnd Sub\r\n', 'standard');
		svc.resetProjectCacheForTests();

		expect(svc.listModules(file).map((m) => m.name)).not.toContain('External');
		// Simulate Excel/git/another window replacing the file behind XLIDE's
		// back: no engine write ever touches `file`, so only the stat check can
		// notice. Bump mtime explicitly so filesystems with coarse timestamps
		// cannot make this test flaky.
		fs.copyFileSync(other, file);
		const bumped = new Date(Date.now() + 5000);
		fs.utimesSync(file, bumped, bumped);

		expect(svc.listModules(file).map((m) => m.name)).toContain('External');
	});

	it('caps the cache and evicts the least recently used entry', () => {
		svc.resetProjectCacheForTests();
		const files = Array.from({ length: 6 }, () => tempCopy());
		for (const file of files) { svc.listModules(file); }
		expect(svc.projectCacheStatsForTests().size).toBeLessThanOrEqual(4);
		// The evicted workbook still reads correctly - just via a fresh parse.
		expect(svc.listModules(files[0]).length).toBeGreaterThan(0);
	});
});

// Annotations that stand for hidden attributes, written on the way into the
// container (parity with xlide_vbide). The code pane never shows these and the
// editor gives no way to set them; a comment in the code is where they are
// said, and this is where they are written.
describe('annotations write the hidden attributes on save', () => {
	it('sets a class attribute the header carries', () => {
		const file = tempCopy();
		svc.writeModule(file, 'CThing', "'@PredeclaredId\r\nOption Explicit\r\n", 'class');
		const { source } = svc.readModule(file, 'CThing', true);
		expect(source).toContain('Attribute VB_PredeclaredId = True');
	});

	it('writes a member attribute under its own procedure', () => {
		const file = tempCopy();
		svc.writeModule(file, 'Helpers', [
			'Option Explicit',
			'',
			"'@Description(\"Totals the rows\")",
			'Public Function Total() As Long',
			'End Function',
			'',
		].join('\r\n'));
		const { source } = svc.readModule(file, 'Helpers', true);
		expect(source).toContain('Attribute Total.VB_Description = "Totals the rows"');
	});

	it('reports what it set, so a save can say so', () => {
		const file = tempCopy();
		const result = svc.writeModule(file, 'CThing', "'@Exposed\r\nOption Explicit\r\n", 'class');
		expect(result.attributeChanges).toEqual(['module: VB_Exposed False -> True']);
	});

	it('is inert without an annotation: the module is written byte for byte', () => {
		const file = tempCopy();
		const body = 'Option Explicit\r\n\r\nPublic Sub Go()\r\nEnd Sub\r\n';
		svc.writeModule(file, 'Plain', body);
		const first = svc.readModule(file, 'Plain', true).source;
		svc.writeModule(file, 'Plain', body);
		expect(svc.readModule(file, 'Plain', true).source).toBe(first);
		expect(first).not.toContain('VB_Description');
	});

	it('survives the round trip: the annotation stays in the code it came from', () => {
		const file = tempCopy();
		svc.writeModule(file, 'Helpers', "'@ModuleDescription(\"Helpers\")\r\nOption Explicit\r\n");
		const { source } = svc.readModule(file, 'Helpers', true);
		// Both halves: the comment the developer wrote, and the attribute it set.
		expect(source).toContain("'@ModuleDescription(\"Helpers\")");
		expect(source).toContain('Attribute VB_Description = "Helpers"');
	});
});
