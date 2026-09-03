import { describe, expect, it } from 'vitest';
import {
	executableFromOpenCommand,
	progIdFromDefaultValue,
	progIdFromNamedValue,
	progIdsFromValueNames,
	vb6ProjectAssociation,
} from '../src/vb6ProjectAssociation';

// F5 on a VB6 project hands the .vbp to the shell. On a machine where
// nothing claims the extension the shell fails with a bare "Application not
// found", so F5 reads the association itself and says what happened. These
// pin the reading against `reg.exe` output as Windows prints it.

const DEFAULT_SET = [
	'',
	'HKEY_CLASSES_ROOT\\.xlsm',
	'    (Default)    REG_SZ    Excel.SheetMacroEnabled.12',
	'',
].join('\r\n');

const DEFAULT_UNSET = [
	'',
	'HKEY_CLASSES_ROOT\\.vbp',
	'    (Default)    REG_SZ    (value not set)',
	'',
].join('\r\n');

const OPEN_WITH = [
	'',
	'HKEY_CLASSES_ROOT\\.vbp\\OpenWithProgids',
	'    twinBASIC.ProjectFile    REG_SZ    ',
	'    VisualBasic.Project    REG_SZ    ',
	'',
].join('\r\n');

const USER_CHOICE = [
	'',
	'HKEY_CURRENT_USER\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\FileExts\\.vbp\\UserChoice',
	'    ProgId    REG_SZ    VisualBasic.Project',
	'    Hash    REG_SZ    abc=',
	'',
].join('\r\n');

describe('reading what Windows can do with a .vbp', () => {
	it('takes a ProgId from a default value, and nothing from an unset one', () => {
		expect(progIdFromDefaultValue(DEFAULT_SET)).toBe('Excel.SheetMacroEnabled.12');
		expect(progIdFromDefaultValue(DEFAULT_UNSET)).toBeUndefined();
		expect(progIdFromDefaultValue('')).toBeUndefined();
	});

	it('takes the user s own choice by name', () => {
		expect(progIdFromNamedValue(USER_CHOICE, 'ProgId')).toBe('VisualBasic.Project');
		expect(progIdFromNamedValue(USER_CHOICE, 'Missing')).toBeUndefined();
	});

	it('lists the Open With candidates, which are value names', () => {
		expect(progIdsFromValueNames(OPEN_WITH)).toEqual(['twinBASIC.ProjectFile', 'VisualBasic.Project']);
		expect(progIdsFromValueNames(DEFAULT_UNSET)).toEqual([]);
	});

	it('names the executable a shell command runs, quoted or bare', () => {
		expect(executableFromOpenCommand('"C:\\Program Files\\VB98\\VB6.EXE" "%1"')).toBe('VB6.EXE');
		expect(executableFromOpenCommand('C:\\tools\\thing.exe %1')).toBe('thing.exe');
		expect(executableFromOpenCommand(undefined)).toBeUndefined();
		expect(executableFromOpenCommand('')).toBeUndefined();
	});

	it('says nothing about a platform where file associations do not apply', () => {
		const answer = vb6ProjectAssociation('darwin');
		expect(answer.unknown).toBe(true);
		// Unknown must not read as "cannot open": F5 still tries there.
		expect(answer.opensDirectly).toBe(false);
		expect(answer.candidates).toEqual([]);
	});
});
