// The Rubberduck `@Folder` annotation, read the way xlide_vbide reads it
// (github.com/WilliamSmithEdward/xlide_vscode/issues/66). Written one way,
// read leniently: the four spellings below all name the same folder.

import { describe, expect, it } from 'vitest';
import { normalizeFolderPath, readFolderAnnotation } from '../src/vba/folderAnnotation';

const folderOf = (source: string): string | undefined =>
	readFolderAnnotation(source).folder;

describe('the spellings that all mean one folder', () => {
	it('reads the bracketed and quoted form', () => {
		expect(folderOf('\'@Folder("Accounts.Ledger")\nOption Explicit\n')).toBe('Accounts.Ledger');
	});

	it('reads it quoted without brackets, bracketed without quotes, and bare', () => {
		expect(folderOf('\'@Folder "Accounts.Ledger"\n')).toBe('Accounts.Ledger');
		expect(folderOf('\'@Folder(Accounts.Ledger)\n')).toBe('Accounts.Ledger');
		expect(folderOf('\'@Folder Accounts.Ledger\n')).toBe('Accounts.Ledger');
	});

	it('does not care about the case of the tag', () => {
		expect(folderOf('\'@folder("accounts.ledger")\n')).toBe('accounts.ledger');
		expect(folderOf('\'@FOLDER("Accounts")\n')).toBe('Accounts');
	});

	it('reads a Rem comment as well as an apostrophe one', () => {
		expect(folderOf('Rem @Folder("Accounts")\n')).toBe('Accounts');
	});

	it('stops the name at the closing bracket, so prose after it is not the folder', () => {
		expect(folderOf('\'@Folder(Accounts) the ledger side\n')).toBe('Accounts');
		expect(folderOf('\'@Folder("Accounts") the ledger side\n')).toBe('Accounts');
	});
});

describe('what is prose rather than an annotation', () => {
	it('requires the name to end at whitespace, a bracket, a quote or the line end', () => {
		expect(folderOf('\'@Folders("Accounts")\n')).toBeUndefined();
		expect(folderOf('\'@Folder-ish("Accounts")\n')).toBeUndefined();
	});

	it('ignores an apostrophe inside a string literal', () => {
		expect(folderOf('Public Const S As String = "it\'s @Folder(Wrong)"\n')).toBeUndefined();
	});

	it('finds nothing in a module that carries none', () => {
		expect(folderOf('Option Explicit\n\nSub T()\nEnd Sub\n')).toBeUndefined();
	});
});

describe('where the read stops', () => {
	it('reads only the declarations section: the first procedure header ends it', () => {
		const source = [
			'Option Explicit',
			'',
			'Sub Recalculate()',
			"    '@Folder(\"Accounts\")",
			'End Sub',
		].join('\n');
		expect(folderOf(source)).toBeUndefined();
	});

	it('treats every procedure header shape as the end', () => {
		for (const header of [
			'Sub T()', 'Function F()', 'Property Get P()', 'Property Let P(v)',
			'Property Set P(v)', 'Public Sub T()', 'Private Static Function F()',
			'Friend Function F()',
		]) {
			expect(folderOf(`${header}\n'@Folder("Late")\n`), header).toBeUndefined();
		}
	});

	it('does not mistake a Declare for a procedure header', () => {
		const source = [
			'Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long',
			"'@Folder(\"Timing\")",
		].join('\n');
		expect(folderOf(source)).toBe('Timing');
	});

	it('keeps the first annotation and ignores later ones', () => {
		expect(folderOf('\'@Folder("First")\n\'@Folder("Second")\n')).toBe('First');
	});
});

describe('normalizing the path', () => {
	it('trims segments, drops empty ones, and rejoins with dots', () => {
		expect(normalizeFolderPath(' Accounts . Ledger ')).toBe('Accounts.Ledger');
		expect(normalizeFolderPath('Accounts..Ledger')).toBe('Accounts.Ledger');
		expect(normalizeFolderPath('.Accounts.')).toBe('Accounts');
	});

	it('reads an annotation naming nothing as no folder at all', () => {
		expect(normalizeFolderPath('  ')).toBe('');
		expect(folderOf('\'@Folder("")\n')).toBeUndefined();
		expect(folderOf('\'@Folder\n')).toBeUndefined();
	});
});

describe('reading a prefix of the module rather than all of it', () => {
	it('settles on the annotation, so the caller need not read further', () => {
		expect(readFolderAnnotation('\'@Folder("A")\nOption Ex', { truncated: true }))
			.toEqual({ folder: 'A', complete: true });
	});

	it('settles on the first procedure header for the same reason', () => {
		expect(readFolderAnnotation('Option Explicit\nSub T()\n    Deb', { truncated: true }))
			.toEqual({ complete: true });
	});

	it('asks for a longer read when the text simply ran out', () => {
		expect(readFolderAnnotation('Option Explicit\nPublic X As Long\n', { truncated: true }))
			.toEqual({ complete: false });
	});

	it('will not read a folder out of a line that was cut mid-word', () => {
		// "Accounts.Led" is not a folder, it is half of one, and answering it
		// would stop the caller from making the read that finds the whole name.
		expect(readFolderAnnotation('Option Explicit\n\'@Folder("Accounts.Led', { truncated: true }))
			.toEqual({ complete: false });
	});
});
