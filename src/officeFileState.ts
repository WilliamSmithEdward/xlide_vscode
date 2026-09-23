// What XLIDE asks about a file an Office application has open, as PowerShell
// the launcher's and the coordinator's scripts both define: whether anything
// holds the file against writes, and whether a copy holds work that closing
// it would lose. Plain strings, so the browser build carries them harmlessly;
// only the desktop ever runs them.
//
// XLIDE closes a copy of the file in its application in several places: to
// free the lock before a save, to refresh a stale read-only view, before F5
// reopens the file read-only to run a macro, and when "Open Read Only" finds
// the file already open for editing. Every one of those closes WITHOUT saving,
// because the application's copy is the stale one - so every one of them
// first asks this, and leaves a copy with unsaved work alone.
//
// A read-only copy is no exception, which is the case that made this
// necessary. Measured on build 16.0.20326: Excel, Word and PowerPoint all
// accept an edit to a copy opened read-only (it can only be saved under
// another name), and each one's Saved property turns false when it does. XLIDE
// used to close such a copy on the grounds that it "holds nothing to lose",
// and the typing went with it.
//
// Anything this cannot determine counts as unsaved. The cost of being wrong
// that way is a save that asks you to close the file yourself; the cost of
// being wrong the other way is someone's work.

import type { OfficeHostApp } from './officeHostApps';

/**
 * PowerShell function: does anything hold `$targetPath` open against writes?
 * Asks for the file with no sharing at all, which fails while any handle
 * that could block an XLIDE save is open.
 */
export const LOCK_TEST_FUNCTION =
	'function Test-XlideLocked { try { $fs = [System.IO.File]::Open($targetPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None); $fs.Close(); return $false } catch { return $true } }';

/**
 * A PowerShell function, `Test-XlideUnsavedWork $copy`, answering true when
 * the open copy holds unsaved work. Define it once at the top of a script and
 * pass the variable holding the copy.
 *
 * - Excel and Word: `Saved` is a Boolean.
 * - PowerPoint: `Saved` is an MsoTriState, and only msoTrue (-1) means saved.
 * - Access has no document-level flag, so it asks each object the database
 *   has loaded: `SysCmd(acSysCmdGetObjectState, type, name)` carries a dirty
 *   bit for a form, report or other object edited and not yet saved. Measured
 *   on a form in design view: 1 when opened, 3 after a property changed.
 *   The copy argument is ignored there - an instance holds one database, and
 *   the script's `$app` is how its objects are reached.
 */
export function unsavedWorkFunction(host: OfficeHostApp): string {
	return `function Test-XlideUnsavedWork($copy) { try { ${UNSAVED_WORK_TESTS[host]} } catch { return $true } }`;
}

const UNSAVED_WORK_TESTS: Record<OfficeHostApp, string> = {
	excel: 'return (-not [bool]$copy.Saved)',
	word: 'return (-not [bool]$copy.Saved)',
	powerpoint: 'return ($copy.Saved -ne -1)',
	// acSysCmdGetObjectState = 10 and the AcObjectType values (acTable 0,
	// acQuery 1, acForm 2, acReport 3, acMacro 4, acModule 5) are from
	// src/analyzer/host/accessObjectModelData.ts. acObjStateDirty is not in
	// that model; its value, 2, was read from Access's own type library.
	access: [
		'$groups = @{ 0 = $app.CurrentData.AllTables; 1 = $app.CurrentData.AllQueries; '
		+ '2 = $app.CurrentProject.AllForms; 3 = $app.CurrentProject.AllReports; '
		+ '4 = $app.CurrentProject.AllMacros; 5 = $app.CurrentProject.AllModules }',
		'foreach ($type in $groups.Keys) { foreach ($object in $groups[$type]) { '
		+ 'if ($object.IsLoaded -and (($app.SysCmd(10, $type, $object.Name) -band 2) -ne 0)) { return $true } } }',
		'return $false',
	].join('; '),
};
