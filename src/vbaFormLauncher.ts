// The F5 "show this form" launcher: what goes into the workbook, and how to
// tell whether it is already there.
//
// The VBE's F5 shows the form; Excel can only run a MACRO, so XLIDE adds one
// with the user's consent. ONE SUB PER FORM, all in a single module: F5 on a
// second form adds its sub beside the first instead of rewriting it, so every
// form keeps its own entry point and nobody loses a launcher by running a
// different form. A form whose sub already exists needs no consent and no
// write - there is nothing new going into the workbook.
//
// vscode-free on purpose, so the rules are unit-testable without an extension
// host (src/vbaFormPreview.ts owns the command that uses them).

/** The module F5 keeps its per-form launcher subs in. */
export const LAUNCHER_MODULE = 'XlideRun';

/** The launcher sub for one form - one per form, never shared. */
export function launcherSubName(formModule: string): string {
	return `XlideShow_${formModule}`;
}

/** True when `source` already declares that launcher sub. */
export function launcherSubExists(source: string | undefined, subName: string): boolean {
	if (!source) { return false; }
	const escaped = subName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
	return new RegExp(`^\\s*(?:public\\s+|private\\s+)?sub\\s+${escaped}\\s*\\(`, 'im').test(source);
}

/**
 * The launcher module's source with this form's sub ADDED - other forms'
 * launchers and any hand edits are kept, because F5 on a second form must not
 * cost the first one its entry point.
 */
export function composeLauncherSource(existingSource: string | undefined, formModule: string): string {
	const subText = [
		`Sub ${launcherSubName(formModule)}()`,
		`    UserForms.Add("${formModule}").Show`,
		'End Sub',
		'',
	].join('\r\n');
	const existing = (existingSource ?? '').replace(/\s+$/, '');
	return existing.length > 0
		? `${existing}\r\n\r\n${subText}`
		: `' XLIDE Run-Form launchers, injected by F5. Safe to delete.\r\n\r\n${subText}`;
}
