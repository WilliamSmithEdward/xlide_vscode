import { describe, expect, it, vi } from 'vitest';
import * as path from 'path';

vi.mock('vscode', async () => (await import('./helpers/vscodeMock')).vscodeMock());

import { hostTokenForFileName, hostObjectModelForToken, EMPTY_HOST_MODEL } from '../src/analyzer/host/hostRegistry';
import {
	containerAppNameForPath,
	containerContextValue,
	isExcelContainerPath,
	isReadOnlyContainerPath,
	isVb6ProjectPath,
	MACRO_CONTAINER_EXTENSIONS,
} from '../src/macroContainerUi';
import { decodeModuleUri, encodeModuleUri } from '../src/xlideFileSystem';
import { analysisSourceForDocument, moduleDocumentUri, moduleLocationOfDocument } from '../src/vbaDocumentLocation';
import { listModules, readModules, resetProjectCacheForTests } from '../src/vba/projectService';
import { moduleKindFromType } from '../src/vbaProjectAnalysis';

// VB6 support arrived beside the workbook engine, never in front of it. This
// file pins the VBA answers that every VB6 touch point could have moved: the
// host each Office extension maps to, the tree context values, the module
// entries a workbook returns, and what a virtual document's location and
// analysis text are. A VB6 change that alters any of them fails here.

const OFFICE_EXTENSIONS = MACRO_CONTAINER_EXTENSIONS.filter((ext) => ext !== 'vbp');

const HOST_BY_EXTENSION: Record<string, string> = {
	xlsm: 'excel', xlsb: 'excel', xlam: 'excel', xltm: 'excel', xls: 'excel', xlt: 'excel', xla: 'excel',
	docm: 'word', dotm: 'word', doc: 'word', dot: 'word',
	pptm: 'powerpoint', potm: 'powerpoint', ppsm: 'powerpoint', ppam: 'powerpoint', ppt: 'powerpoint', ppa: 'powerpoint',
	accdb: 'access', accda: 'access', mdb: 'access', mda: 'access',
};

describe('the Office containers answer exactly as before VB6', () => {
	it('maps every Office extension to its host and never to vb6', () => {
		for (const ext of OFFICE_EXTENSIONS) {
			const file = `C:\\work\\Book.${ext}`;
			expect(hostTokenForFileName(file), ext).toBe(HOST_BY_EXTENSION[ext]);
			expect(isVb6ProjectPath(file), ext).toBe(false);
			expect(containerContextValue(file), ext).not.toBe('vb6Project');
		}
		expect(Object.keys(HOST_BY_EXTENSION).sort()).toEqual([...OFFICE_EXTENSIONS].sort());
	});

	it('keeps the context values, read-only rule, and app names for Office files', () => {
		expect(containerContextValue('C:\\w\\Book.xlsm')).toBe('xlsm');
		expect(containerContextValue('C:\\w\\Doc.docm')).toBe('macroDocument');
		expect(containerContextValue('C:\\w\\Deck.pptm')).toBe('macroDocument');
		expect(containerContextValue('C:\\w\\Db.accdb')).toBe('macroReadOnly');
		expect(isReadOnlyContainerPath('C:\\w\\Db.accdb')).toBe(true);
		expect(isReadOnlyContainerPath('C:\\w\\Book.xlsm')).toBe(false);
		expect(isExcelContainerPath('C:\\w\\Book.xlsm')).toBe(true);
		expect(isExcelContainerPath('C:\\w\\App.vbp')).toBe(false);
		expect(containerAppNameForPath('C:\\w\\Book.xlsm')).toBe('Excel');
		expect(containerAppNameForPath('C:\\w\\Doc.docm')).toBe('Word');
		// An unrecognized extension still defaults to Excel, as it always did.
		expect(containerAppNameForPath('C:\\w\\thing.unknown')).toBe('Excel');
		expect(hostTokenForFileName('C:\\w\\thing.unknown')).toBeUndefined();
	});

	it('leaves the host models alone: Excel is the default, vb6 answers as itself', () => {
		expect(hostObjectModelForToken(undefined)).toBeUndefined();
		expect(hostObjectModelForToken('excel')).toBeUndefined();
		expect(hostObjectModelForToken('word')).toBeDefined();
		expect(hostObjectModelForToken('word')).not.toBe(EMPTY_HOST_MODEL);
		// The vb6 model exists now (roadmap Slice 3); the collision claim is
		// that it is its own model, never Excel's and never the empty one.
		expect(hostObjectModelForToken('vb6')).not.toBe(EMPTY_HOST_MODEL);
		expect(hostObjectModelForToken('vb6')?.hostName).toBe('VB6');
		expect(hostObjectModelForToken('vb6')).not.toBe(hostObjectModelForToken('word'));
	});

	it('maps the project module kinds as before; only VB6 kinds are new', () => {
		expect(moduleKindFromType('standard')).toBe('standard');
		expect(moduleKindFromType('class')).toBe('class');
		expect(moduleKindFromType('document')).toBe('document');
		expect(moduleKindFromType('userform')).toBe('userform');
		expect(moduleKindFromType(undefined)).toBe('standard');
		expect(moduleKindFromType('anything-else')).toBe('standard');
	});
});

describe('a project module still has no file of its own', () => {
	const FIXTURE = path.join(__dirname, 'fixtures', 'binaries', 'FormFixtureVbide.xlsm');

	it('lists and reads project modules without a filePath', () => {
		resetProjectCacheForTests();
		for (const entry of [...listModules(FIXTURE), ...readModules(FIXTURE)]) {
			expect(entry.filePath, entry.name).toBeUndefined();
		}
		resetProjectCacheForTests();
	});

	it('opens a project module at its virtual document', () => {
		const uri = moduleDocumentUri('C:\\w\\Book.xlsm', { moduleName: 'Module1' });
		expect(uri.toString()).toBe(encodeModuleUri('C:\\w\\Book.xlsm', 'Module1').toString());
	});
});

describe('a virtual document is located and read exactly as before', () => {
	const uri = encodeModuleUri('C:\\w\\Book.xlsm', 'Form1');
	const text = 'VERSION 5.00\r\nBegin VB.Form Form1\r\nEnd\r\nOption Explicit\r\n';
	const document = { uri, languageId: 'xlide-vba', version: 1, getText: () => text } as never;

	it('decodes its URI, the same answer decodeModuleUri gives', () => {
		const decoded = decodeModuleUri(uri);
		expect(moduleLocationOfDocument(document)).toEqual({
			projectPath: decoded.projectPath,
			moduleName: decoded.moduleName,
			native: false,
		});
	});

	it('hands the analyzer its text untouched, even when the text opens with VERSION', () => {
		// The workbook's virtual document IS the code; blanking belongs to
		// files on disk only.
		expect(analysisSourceForDocument(document)).toBe(text);
	});
});
