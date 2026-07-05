import { describe, expect, it } from 'vitest';
import {
	isMacCltStubMessage,
	isMissingPackageMessage,
	isPythonNotFoundMessage,
	posixPythonCandidatePaths,
} from '../src/pythonEnvironment';

const MAC_CLT_STUB =
	'Python backend exited (code 1).\n'
	+ 'xcrun: error: invalid active developer path (/Library/Developer/CommandLineTools), '
	+ 'missing xcrun at: /Library/Developer/CommandLineTools/usr/bin/xcrun';

const MAC_XCODE_SELECT =
	"Python backend exited (code 1).\nxcode-select: note: no developer tools were found at "
	+ "'/Applications/Xcode.app', requesting install.";

describe('isPythonNotFoundMessage', () => {
	it.each([
		['spawn ENOENT (posix/win)', 'Cannot start Python: spawn python ENOENT'],
		['cmd.exe not recognized', "'python' is not recognized as an internal or external command"],
		['Windows Store alias stub', 'Python was not found; run without arguments to install from the Microsoft Store, or disable this shortcut from Settings.'],
		['macOS CLT stub (xcrun)', MAC_CLT_STUB],
		['macOS CLT stub (xcode-select)', MAC_XCODE_SELECT],
	])('classifies as python-not-found: %s', (_label, msg) => {
		expect(isPythonNotFoundMessage(msg)).toBe(true);
	});

	it.each([
		['missing package', "ModuleNotFoundError: No module named 'pyOpenVBA'"],
		['ordinary crash', 'Python backend exited (code 1).\nTraceback (most recent call last):\n  ...\nValueError: boom'],
	])('does not classify: %s', (_label, msg) => {
		expect(isPythonNotFoundMessage(msg)).toBe(false);
	});
});

describe('isMacCltStubMessage', () => {
	it('recognizes both stub variants and nothing else', () => {
		expect(isMacCltStubMessage(MAC_CLT_STUB)).toBe(true);
		expect(isMacCltStubMessage(MAC_XCODE_SELECT)).toBe(true);
		expect(isMacCltStubMessage('Cannot start Python: spawn python ENOENT')).toBe(false);
		expect(isMacCltStubMessage('ValueError: boom')).toBe(false);
	});
});

describe('isMissingPackageMessage', () => {
	it('matches missing-module errors only', () => {
		expect(isMissingPackageMessage("ModuleNotFoundError: No module named 'openpyxl'")).toBe(true);
		expect(isMissingPackageMessage(MAC_CLT_STUB)).toBe(false);
	});
});

describe('posixPythonCandidatePaths', () => {
	it('offers the well-known macOS locations, Homebrew (Apple Silicon) first', () => {
		expect(posixPythonCandidatePaths('darwin')).toEqual([
			'/opt/homebrew/bin/python3',
			'/usr/local/bin/python3',
			'/Library/Frameworks/Python.framework/Versions/Current/bin/python3',
		]);
	});

	it('is empty on win32 and linux (PATH resolution suffices)', () => {
		expect(posixPythonCandidatePaths('win32')).toEqual([]);
		expect(posixPythonCandidatePaths('linux')).toEqual([]);
	});
});
