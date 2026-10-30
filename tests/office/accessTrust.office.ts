// F5 in Access on a database outside a trusted location, through XLIDE's own
// runHostMacro. Access opens such a database with its code disabled behind
// the message bar, and Run then fails with no reason given; XLIDE says what
// to do instead and runs nothing. A database whose code is enabled runs.
//
// Measured on Access 16.0 (build 20527): an Access an automation client
// starts is at AutomationSecurity 2 (by UI), and once it is handed to the
// user, as XLIDE hands every instance, a database it opens is untrusted at
// AutomationSecurity 1 as at 2. A database the check opens itself
// at AutomationSecurity 1 before the hand-over is trusted, and XLIDE's F5,
// attaching to that instance, finds it already open and runs in it.
//
// Runs only when Access is not running, so every instance it touches is one
// the check started, and quits it when done.

import * as fs from 'fs';
import * as path from 'path';
import { afterEach, describe, expect, it } from 'vitest';
import { HostMacroError, runHostMacro } from '../../src/officeHostLauncher';
import { writeModule } from '../../src/vba/projectService';
import { SCRATCH, hostIsFree, ps, q, scratchCopy, waitForExit } from './officeHarness';

/** A scratch database whose procedure Mark writes a file, so a run can be seen. */
function databaseWithMark(label: string): { file: string; marker: string } {
	const file = scratchCopy('AccessFixture.accdb', label);
	const marker = path.join(path.dirname(file), 'ran.txt');
	writeModule(file, 'XlideMark', [
		'Public Sub Mark()',
		`    Open "${marker}" For Output As #1`,
		'    Print #1, "ran"',
		'    Close #1',
		'End Sub',
		'',
	].join('\r\n'), 'standard');
	return { file, marker };
}

/**
 * Quits the Access the check started, and only while the database it holds
 * is the check's own scratch file (or none): anything else is left as it is.
 */
async function quitOwnAccess(): Promise<void> {
	const result = await ps([
		`$root = ${q(SCRATCH)}`,
		'$app = $null',
		'try { $app = [Runtime.InteropServices.Marshal]::GetActiveObject("Access.Application") } catch { }',
		'if ($app) {',
		'  $open = ""',
		'  try { $open = $app.CurrentProject.FullName } catch { }',
		'  if ($open -and -not $open.StartsWith($root, [StringComparison]::OrdinalIgnoreCase)) { "left=" + $open }',
		'  else { try { $app.CloseCurrentDatabase() } catch { }; $app.Quit(2); "quit" }',
		'}',
		'$app = $null',
		'[GC]::Collect()',
	].join('\n'));
	expect(result.out.filter((line) => line.startsWith('left='))).toEqual([]);
	expect(await waitForExit('access')).toBe(true);
}

/** Presses Enable Content in the message bar of the Access the check started. */
const ENABLE_CONTENT = [
	'Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes',
	'$app = [Runtime.InteropServices.Marshal]::GetActiveObject("Access.Application")',
	'$window = [Windows.Automation.AutomationElement]::FromHandle([IntPtr]$app.hWndAccessApp())',
	'$named = New-Object Windows.Automation.PropertyCondition ([Windows.Automation.AutomationElement]::NameProperty), "Enable Content"',
	'$button = $window.FindFirst([Windows.Automation.TreeScope]::Descendants, $named)',
	'if ($button) { $button.GetCurrentPattern([Windows.Automation.InvokePattern]::Pattern).Invoke(); "pressed" }',
	'$until = (Get-Date).AddSeconds(10)',
	'while (-not $app.CurrentProject.IsTrusted -and (Get-Date) -lt $until) { Start-Sleep -Milliseconds 100 }',
	'"trusted=" + $app.CurrentProject.IsTrusted',
].join('\n');

const run = (file: string) => runHostMacro(file, { moduleName: 'XlideMark', procedureName: 'Mark' }, { attachToRunning: true }, () => undefined);

describe.runIf(await hostIsFree('access'))('Access F5 on a database outside a trusted location', () => {
	afterEach(quitOwnAccess);

	it('says how to enable the code and runs nothing, again while it stays disabled, and runs once it is enabled', async () => {
		const { file, marker } = databaseWithMark('trust-disabled');
		const failed = async (): Promise<void> => {
			const failure = await run(file).then(() => undefined, (error: unknown) => error);
			expect(failure).toBeInstanceOf(HostMacroError);
			expect((failure as HostMacroError).code).toBe('RUN_FAILED');
			expect((failure as HostMacroError).message).toMatch(/^Access opened the database with its code disabled, .*Click Enable Content/);
			expect(fs.existsSync(marker)).toBe(false);
		};
		// XLIDE starts Access and opens the database.
		await failed();
		// F5 again attaches to that Access, where the database is still open and disabled.
		await failed();
		// The user enables it, and F5 runs in the same instance.
		const enabled = await ps(ENABLE_CONTENT);
		expect(enabled.out).toEqual(['pressed', 'trusted=True']);
		await run(file);
		expect(fs.existsSync(marker)).toBe(true);
	});

	it('says to unblock a database downloaded from the internet, and then offers Enable Content', async () => {
		const { file, marker } = databaseWithMark('trust-downloaded');
		const marked = await ps([
			`Set-Content -LiteralPath ${q(file)} -Stream Zone.Identifier -Value "[ZoneTransfer]\`r\`nZoneId=3"`,
			`"zone=" + ((Get-Content -LiteralPath ${q(file)} -Stream Zone.Identifier) -join " ")`,
		].join('\n'));
		expect(marked.out).toEqual(['zone=[ZoneTransfer] ZoneId=3']);
		const message = async (): Promise<string> => {
			const failure = await run(file).then(() => undefined, (error: unknown) => error);
			expect(failure).toBeInstanceOf(HostMacroError);
			expect((failure as HostMacroError).code).toBe('RUN_FAILED');
			return (failure as HostMacroError).message;
		};
		expect(await message()).toMatch(/^Access blocked the database's code, because the file was downloaded from the internet\. .*Unblock/);
		// What the message says: close the database, unblock the file, F5 again.
		const unblocked = await ps([
			'$app = [Runtime.InteropServices.Marshal]::GetActiveObject("Access.Application")',
			'$app.CloseCurrentDatabase()',
			`Unblock-File -LiteralPath ${q(file)}`,
			`"zone=" + [bool](Get-Item -LiteralPath ${q(file)} -Stream Zone.Identifier -ErrorAction SilentlyContinue)`,
		].join('\n'));
		expect(unblocked.out).toEqual(['zone=False']);
		expect(await message()).toMatch(/^Access opened the database with its code disabled, .*Click Enable Content/);
		expect(fs.existsSync(marker)).toBe(false);
	});

	it('runs the procedure in a database whose code is enabled', async () => {
		const { file, marker } = databaseWithMark('trust-enabled');
		// Opened by the check at AutomationSecurity 1, then handed over as
		// XLIDE hands its own, so the instance stays when this script ends.
		const opened = await ps([
			'$app = New-Object -ComObject Access.Application',
			'$app.AutomationSecurity = 1',
			`$app.OpenCurrentDatabase(${q(file)})`,
			'"trusted=" + $app.CurrentProject.IsTrusted',
			'try { $app.Visible = $true } catch { }',
			'try { $app.UserControl = $true } catch { }',
		].join('\n'));
		expect(opened.out).toContain('trusted=True');
		await run(file);
		expect(fs.existsSync(marker)).toBe(true);
	});
});
