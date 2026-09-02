import { afterEach, describe, expect, it, vi } from 'vitest';
import * as fs from 'fs';
import * as path from 'path';

vi.mock('vscode', async () => (await import('./helpers/vscodeMock')).vscodeMock());

import { blankDesignerHeader, designerHeaderEnd, listProcedures } from '../src/vba/moduleSource';
import { vbaHeaderBlockEnd } from '../src/vbaSourceScan';
import { analyzeVbaModuleSource } from '../src/vbaModuleAnalysis';
import { ownersFromListings, setVb6ModuleOwnersForTests } from '../src/vb6ProjectLocator';
import { moduleDocumentUri, moduleLocationOfDocument } from '../src/vbaDocumentLocation';
import { workbookIdentityKey } from '../src/workbookIdentity';
import {
	isStandaloneVbaDocument,
	moduleKindFromDocument,
	moduleNameFromDocument,
} from '../src/vbaDocumentIdentity';

// A VB6 module is a file on disk, and the editor shows the whole file - the
// designer block a form opens with included. The analyzer must see VBA only,
// at the file's own offsets, and every surface must know which project the
// file belongs to. These pin both halves.

const FORM = path.join(__dirname, 'fixtures', 'vb6', 'RunAsTrustedInstaller', 'Form1.frm');
const VBP = path.join(__dirname, 'fixtures', 'vb6', 'RunAsTrustedInstaller', 'Project1.vbp');

afterEach(() => {
	setVb6ModuleOwnersForTests(new Map());
});

describe('the designer header, blanked in place', () => {
	const text = fs.readFileSync(FORM, 'latin1');

	it('ends where the form block closes, BeginProperty groups notwithstanding', () => {
		const end = designerHeaderEnd(text);
		expect(end).toBeGreaterThan(0);
		expect(text.slice(0, end)).toMatch(/BeginProperty Font/);
		expect(text.slice(0, end).trimEnd()).toMatch(/\r\nEnd$/);
		expect(text.slice(end)).toMatch(/^Attribute VB_Name = "Form1"/);
	});

	it('keeps every offset: same length, same line breaks, whitespace where the block was', () => {
		const blanked = blankDesignerHeader(text);
		expect(blanked.length).toBe(text.length);
		const end = designerHeaderEnd(text);
		expect(blanked.slice(0, end).replace(/[ \r\n]/g, '')).toBe('');
		expect(blanked.slice(end)).toBe(text.slice(end));
		expect(blanked.split('\r\n').length).toBe(text.split('\r\n').length);
	});

	it('lets the analyzer read a real VB6 form file with no header noise', () => {
		const diagnostics = analyzeVbaModuleSource({
			source: blankDesignerHeader(text),
			moduleName: 'Form1',
			moduleType: 'userform',
			moduleKind: 'userform',
			host: 'vb6',
		} as never).diagnostics;
		expect(diagnostics.filter((d) => d.code === 'statement-outside-procedure')).toEqual([]);
		expect(diagnostics.filter((d) => d.code === 'option-after-declaration')).toEqual([]);
	});

	it('reports procedures at the lines they hold in the file', () => {
		const lines = listProcedures(blankDesignerHeader(text));
		const formLoad = lines.find((p) => p.name === 'Form_Load');
		expect(formLoad?.line).toBe(128);
		expect(text.split('\r\n')[127]).toMatch(/^Private Sub Form_Load\(\)/);
	});

	it('handles a class preamble and leaves headerless text alone', () => {
		const cls = 'VERSION 1.0 CLASS\r\nBEGIN\r\n  MultiUse = -1  \'True\r\nEND\r\nAttribute VB_Name = "C"\r\nOption Explicit\r\n';
		const blanked = blankDesignerHeader(cls);
		expect(blanked.length).toBe(cls.length);
		expect(blanked).toMatch(/^\s+Attribute VB_Name = "C"/);
		expect(blankDesignerHeader('Option Explicit\r\nSub A()\r\nEnd Sub\r\n')).toBe('Option Explicit\r\nSub A()\r\nEnd Sub\r\n');
		// A block that never closes is not a header: nothing is hidden.
		expect(blankDesignerHeader('VERSION 5.00\r\nBegin VB.Form F\r\n   Caption = "x"\r\nSub A()\r\nEnd Sub\r\n'))
			.toMatch(/^VERSION 5\.00/);
	});

	it('is skipped by the sync preview header walker too', () => {
		// The walker that hides headers in sync previews used to stop at the
		// first BeginProperty line and hide nothing on a real VB6 form.
		const end = vbaHeaderBlockEnd(text.split('\r\n'));
		expect(end).toBeGreaterThan(0);
		expect(text.split('\r\n')[end]).toMatch(/^Attribute VB_Name/);
	});
});

describe('which project a file belongs to', () => {
	it('maps every listed file to its manifest, first manifest winning', () => {
		const owners = ownersFromListings([
			{ vbpPath: '/a/A.vbp', modules: [
				{ name: 'Form1', type: 'userform', filePath: '/a/Form1.frm' },
				{ name: 'modShared', type: 'standard', filePath: '/shared/modShared.bas' },
			] },
			{ vbpPath: '/b/B.vbp', modules: [
				{ name: 'modShared', type: 'standard', filePath: '/shared/modShared.bas' },
				{ name: 'NoFile', type: 'standard' },
			] },
		]);
		const key = (p: string) => workbookIdentityKey(p);
		expect(owners.get(key('/a/Form1.frm'))).toEqual({ vbpPath: '/a/A.vbp', moduleName: 'Form1', moduleType: 'userform' });
		expect(owners.get(key('/shared/modShared.bas'))?.vbpPath).toBe('/a/A.vbp');
		expect(owners.size).toBe(2);
	});

	it('folds path case on Windows, where the file system does', () => {
		if (process.platform !== 'win32') {
			return;
		}
		const owners = ownersFromListings([
			{ vbpPath: 'C:\\a\\A.vbp', modules: [{ name: 'modShared', type: 'standard', filePath: 'C:\\Shared\\modShared.bas' }] },
			{ vbpPath: 'C:\\b\\B.vbp', modules: [{ name: 'modShared', type: 'standard', filePath: 'C:\\SHARED\\MODSHARED.BAS' }] },
		]);
		expect(owners.size).toBe(1);
		expect(owners.get(workbookIdentityKey('c:\\shared\\modshared.bas'))?.vbpPath).toBe('C:\\a\\A.vbp');
	});

	it('answers a document\'s location, name and kind from its owner', () => {
		const frm = path.join(path.dirname(VBP), 'Form1.frm');
		setVb6ModuleOwnersForTests(ownersFromListings([
			{ vbpPath: VBP, modules: [{ name: 'Form1', type: 'userform', filePath: frm }] },
		]));
		const owned = fakeDocument('file', frm);
		expect(moduleLocationOfDocument(owned)).toEqual({ xlsmPath: VBP, moduleName: 'Form1', moduleType: 'userform', native: true });
		expect(moduleNameFromDocument(owned)).toBe('Form1');
		expect(moduleKindFromDocument(owned)).toBe('userform');
		expect(isStandaloneVbaDocument(owned)).toBe(false);

		const loose = fakeDocument('file', path.join(path.dirname(VBP), 'Other.bas'));
		expect(moduleLocationOfDocument(loose)).toBeUndefined();
		expect(moduleNameFromDocument(loose)).toBe('Other');
		expect(isStandaloneVbaDocument(loose)).toBe(true);
	});

	it('opens a module at its file when it has one, else at its virtual document', () => {
		const native = moduleDocumentUri(VBP, { moduleName: 'Form1', filePath: 'C:\\p\\Form1.frm' });
		expect(native.scheme).toBe('file');
		const virtual = moduleDocumentUri('C:\\p\\Book.xlsm', { moduleName: 'Module1' });
		expect(virtual.scheme).toBe('xlide-vba');
	});
});

function fakeDocument(scheme: string, fsPath: string) {
	const uriPath = fsPath.replace(/\\/g, '/');
	return {
		uri: { scheme, fsPath, path: scheme === 'file' ? '/' + uriPath : uriPath, toString: () => `${scheme}:///${uriPath}` },
		languageId: 'vba',
		version: 1,
		getText: () => '',
	} as never;
}
