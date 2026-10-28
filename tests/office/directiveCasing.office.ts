// The spelling the VBE gives conditional-compilation names (issue #1497),
// read from the applications themselves: XLIDE writes the probe modules, the
// application opens the file, and the check reads each module back from the
// VBE. The VBE keeps one spelling per name for the whole project, so each
// probe is a project of its own.
//
// Excel and Access run in instances the check starts itself. Word and
// PowerPoint only when neither is running, so no instance of theirs is ever
// touched.

import { describe, expect, it } from 'vitest';
import { writeModule } from '../../src/vba/projectService';
import { conditionalConstNames, conditionalNameSpellings } from '../../src/analyzer/format/conditionalNames';
import { formatVbaModule } from '../../src/analyzer/format/formatModule';
import type { OfficeHostApp } from '../../src/officeHostApps';
import { hostInstalled, hostIsFree, psFile, q, scratchCopy, waitForExit } from './officeHarness';

const FIXTURE: Record<OfficeHostApp, string> = {
	excel: 'FormFixture.xlsm',
	word: 'WordFixture.docm',
	powerpoint: 'PowerPointFixture.pptm',
	access: 'AccessFixture.accdb',
};

/** PowerShell that opens `file` as `$proj` (its VBProject), and closes it again. */
function openAndClose(host: OfficeHostApp, file: string): { open: string[]; close: string[] } {
	switch (host) {
		case 'excel':
			return {
				open: [
					'$app = [Activator]::CreateInstance([type]::GetTypeFromProgID("Excel.Application"))',
					'$app.DisplayAlerts = $false',
					`$doc = $app.Workbooks.Open(${q(file)})`,
					'$proj = $doc.VBProject',
				],
				close: ['$doc.Close($false)', '$doc = $null', '$app.Quit()'],
			};
		case 'word':
			return {
				open: [
					'$app = [Activator]::CreateInstance([type]::GetTypeFromProgID("Word.Application"))',
					'$app.DisplayAlerts = 0',
					`$doc = $app.Documents.Open(${q(file)})`,
					'$proj = $doc.VBProject',
				],
				close: ['$doc.Close(0)', '$doc = $null', '$app.Quit(0)'],
			};
		case 'powerpoint':
			return {
				open: [
					'$app = [Activator]::CreateInstance([type]::GetTypeFromProgID("PowerPoint.Application"))',
					// Presentations.Open(FileName, ReadOnly:=msoFalse, Untitled:=msoFalse, WithWindow:=msoFalse)
					`$doc = $app.Presentations.Open(${q(file)}, 0, 0, 0)`,
					'$proj = $doc.VBProject',
				],
				close: ['$doc.Close()', '$doc = $null', '$app.Quit()'],
			};
		case 'access':
			return {
				open: [
					'$app = [Activator]::CreateInstance([type]::GetTypeFromProgID("Access.Application"))',
					`$app.OpenCurrentDatabase(${q(file)})`,
					'$proj = $app.VBE.ActiveVBProject',
				],
				close: ['$app.CloseCurrentDatabase()', '$app.Quit()'],
			};
	}
}

/**
 * Writes `modules` into a copy of the host's fixture, in order, opens it in
 * the host and returns each module as the VBE shows it.
 */
async function shownText(
	host: OfficeHostApp,
	label: string,
	modules: Readonly<Record<string, readonly string[]>>,
	accessArguments?: string,
): Promise<Record<string, string[]>> {
	const file = scratchCopy(FIXTURE[host], `casing-${host}-${label}`);
	if (accessArguments !== undefined) {
		// The project's Conditional Compilation Arguments are an Access
		// option. Access stores them before XLIDE writes the modules, so the
		// open below loads both from the file, as a user's project would.
		await psFile([
			'$ErrorActionPreference = "Continue"',
			'$app = [Activator]::CreateInstance([type]::GetTypeFromProgID("Access.Application"))',
			`$app.OpenCurrentDatabase(${q(file)})`,
			`$app.SetOption("Conditional Compilation Arguments", ${q(accessArguments)})`,
			'$app.CloseCurrentDatabase()',
			'$app.Quit()',
			'$app = $null',
			'[GC]::Collect()',
			'[GC]::WaitForPendingFinalizers()',
		].join('\r\n'), `casing-${host}-${label}-arguments.ps1`);
		await waitForExit('access');
	}
	for (const [name, lines] of Object.entries(modules)) {
		writeModule(file, name, `${lines.join('\r\n')}\r\n`, 'standard');
	}
	const { open, close } = openAndClose(host, file);
	const read = Object.keys(modules).map((name) => [
		`$cm = $proj.VBComponents.Item(${q(name)}).CodeModule`,
		`"${name}=" + [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($cm.Lines(1, $cm.CountOfLines)))`,
	].join('\r\n'));
	const result = await psFile([
		'$ErrorActionPreference = "Continue"',
		...open,
		...read,
		'$cm = $null',
		'$proj = $null',
		...close,
		'$app = $null',
		'[GC]::Collect()',
		'[GC]::WaitForPendingFinalizers()',
	].join('\r\n'), `casing-${host}-${label}.ps1`);
	const shown: Record<string, string[]> = {};
	for (const line of result.out) {
		const at = line.indexOf('=');
		if (at > 0 && line.slice(0, at) in modules) {
			shown[line.slice(0, at)] = Buffer.from(line.slice(at + 1), 'base64').toString('utf8').split(/\r\n/);
		}
	}
	expect(Object.keys(shown), result.err.join('\n')).toEqual(Object.keys(modules));
	return shown;
}

/** VBA's own compiler constants, on directive and code lines. */
const CONSTANTS = {
	XlideA: [
		'#If vba7 Then',
		'#ElseIf vba6 Then',
		'#End If',
		'#If win64 Or win32 Or win16 Or mac Then',
		'#End If',
		'#If twinbasic Then',
		'#End If',
		'Public Sub Probe()',
		'    Dim y As Variant',
		'    y = win64',
		'End Sub',
	],
};
const CONSTANTS_SHOWN = {
	XlideA: [
		'#If VBA7 Then',
		'#ElseIf VBA6 Then',
		'#End If',
		'#If Win64 Or Win32 Or Win16 Or Mac Then',
		'#End If',
		'#If twinbasic Then',
		'#End If',
		'Public Sub Probe()',
		'    Dim y As Variant',
		'    y = Win64',
		'End Sub',
	],
};

/** A declaration of a compiler constant's name takes its spelling over. */
const OVERRIDES = {
	XlideA: ['#Const WIN64 = 0', '#If win64 Then', '#End If', '#If mac Then', '#End If'],
	XlideB: ['Public mac As Long'],
};
const OVERRIDES_SHOWN = {
	XlideA: ['#Const WIN64 = 0', '#If WIN64 Then', '#End If', '#If mac Then', '#End If'],
	XlideB: ['Public mac As Long'],
};

/**
 * A `#Const` name prints as its last declaration in load order: across
 * modules, in an uncompiled branch, inside a procedure, and in the shapes a
 * `#Const` line can take.
 */
const LAST_DECLARATION = {
	XlideA: [
		'#Const Flag = 1',
		'#If FLAG Or cont Or brk Or remv Or colv Or dead Or inproc Then',
		'#End If',
		'Public Sub Q()',
		'    z = flag',
		'End Sub',
	],
	XlideB: [
		'#Const fLaG = 2',
		'# _',
		'Const CONT = 1',
		'#Const [BRK] = 1',
		'#Const REMV = 1 Rem note',
		'#Const COLV = 1:',
		'#If False Then',
		'#Const DEAD = 1',
		'#End If',
		'Public Sub P()',
		'#Const INPROC = 1',
		'End Sub',
	],
};

const LAST_DECLARATION_SHOWN = {
	XlideA: [
		'#Const fLaG = 1',
		'#If fLaG Or CONT Or BRK Or REMV Or COLV Or DEAD Or INPROC Then',
		'#End If',
		'Public Sub Q()',
		'    z = fLaG',
		'End Sub',
	],
	// The VBE drops the brackets of `[BRK]`. The formatter never changes a
	// token, so that line is the one place its output differs.
	XlideB: [...LAST_DECLARATION.XlideB.slice(0, 3), '#Const BRK = 1', ...LAST_DECLARATION.XlideB.slice(4)],
};

/** The project's arguments come first, so a module's `#Const` spelling wins. */
const ARGUMENTS = 'ARGNAME = 1 : ARGTWO = 1';
const ARGUMENT_MODULES = { XlideA: ['#If argname Or argtwo Then', '#End If'], XlideB: ['#Const ArgTwo = 1'] };
const ARGUMENT_MODULES_SHOWN = { XlideA: ['#If ARGNAME Or ArgTwo Then', '#End If'], XlideB: ['#Const ArgTwo = 1'] };

/** Each line without its indent: the formatter indents, and only the casing is compared. */
function unindented(modules: Readonly<Record<string, readonly string[]>>): Record<string, string[]> {
	return Object.fromEntries(Object.entries(modules).map(([name, lines]) => [name, lines.map((line) => line.trimStart())]));
}

/** Format Document's casing for a whole project, unindented, with no ordinary declarations in play. */
function formatted(modules: Readonly<Record<string, readonly string[]>>, projectArguments?: string): Record<string, string[]> {
	const spellings = conditionalNameSpellings(projectArguments, Object.values(modules).map((lines) => conditionalConstNames(lines.join('\r\n'))));
	const formattedModules = Object.fromEntries(Object.entries(modules).map(([name, lines]) => {
		const result = formatVbaModule(lines.join('\r\n'), {
			tabSize: 4,
			insertSpaces: true,
			identifierCase: (identifier) => spellings.get(identifier.toLowerCase()),
		});
		return [name, (result.text ?? '').split('\r\n')];
	}));
	return unindented(formattedModules);
}

const usable = {
	excel: await hostInstalled('excel'),
	word: await hostIsFree('word'),
	powerpoint: await hostIsFree('powerpoint'),
	access: await hostInstalled('access'),
};

describe.each(['excel', 'word', 'powerpoint', 'access'] as const)('conditional-compilation names as %s shows them (issue #1497)', (host) => {
	it.runIf(usable[host])('spells VBA\'s compiler constants its way, on directive and code lines', async () => {
		expect(await shownText(host, 'constants', CONSTANTS)).toEqual(CONSTANTS_SHOWN);
		expect(formatted(CONSTANTS)).toEqual(unindented(CONSTANTS_SHOWN));
	});

	it.runIf(usable[host])('lets a #Const or an ordinary declaration of the name take over', async () => {
		expect(await shownText(host, 'overrides', OVERRIDES)).toEqual(OVERRIDES_SHOWN);
	});

	it.runIf(usable[host])('prints a #Const name as its last declaration in load order', async () => {
		expect(await shownText(host, 'last', LAST_DECLARATION)).toEqual(LAST_DECLARATION_SHOWN);
		expect(formatted(LAST_DECLARATION)).toEqual(unindented({ ...LAST_DECLARATION_SHOWN, XlideB: LAST_DECLARATION.XlideB }));
	});

	// Only Access sets the arguments from automation (Application.SetOption);
	// the other hosts keep them in the same project record.
	it.runIf(usable[host] && host === 'access')('reads the project\'s arguments before the modules', async () => {
		expect(await shownText(host, 'arguments', ARGUMENT_MODULES, ARGUMENTS)).toEqual(ARGUMENT_MODULES_SHOWN);
		expect(formatted(ARGUMENT_MODULES, ARGUMENTS)).toEqual(unindented(ARGUMENT_MODULES_SHOWN));
	});
});
