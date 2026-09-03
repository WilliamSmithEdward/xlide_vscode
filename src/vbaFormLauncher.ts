// The F5 "show this form" launcher: what goes into the workbook, and how to
// tell whether it is already there.
//
// The VBE's F5 shows the form; Excel can only run a MACRO, so XLIDE adds one
// with the user's consent. ONE SUB PER FORM, all in a single module: F5 on a
// second form adds its sub beside the first instead of rewriting it, so every
// form keeps its own entry point and nobody loses a launcher by running a
// different form. A form whose sub already exists needs no consent and no
// write - there is nothing new going into the project.
//
// vscode-free on purpose, so the rules are unit-testable without an extension
// host (src/vbaFormPreview.ts owns the command that uses them).

/** The module F5 keeps its per-form launcher subs in. */
export const LAUNCHER_MODULE = 'XlideRun';

const LAUNCHER_HEADER = "' XLIDE Run-Form launchers, injected by F5. Safe to delete.";
const OPTION_EXPLICIT = 'Option Explicit';

/** Whether the module already opens with the declaration XLIDE's own analysis asks for. */
function declaresOptionExplicit(source: string): boolean {
	return /^[ \t]*option[ \t]+explicit\b/im.test(source);
}

/**
 * Whether every line is one XLIDE wrote: the header, the option, and launcher
 * subs. A module anyone has touched fails this, and then the option is left
 * alone - adding it to code that was written without it is how you turn
 * someone's working macro into a compile error.
 */
function isUntouchedLauncherModule(source: string): boolean {
	return source.split(/\r?\n/).every((line) => {
		const text = line.trim();
		return text === ''
			|| text === LAUNCHER_HEADER
			|| text === OPTION_EXPLICIT
			|| /^sub\s+XlideShow_\w+\s*\(\s*\)$/i.test(text)
			|| /^UserForms\.Add\("[^"]*"\)\.Show$/i.test(text)
			|| /^end\s+sub$/i.test(text);
	});
}

/** The module's source with `Option Explicit` in its declarations, below the header comment. */
function withOptionExplicit(source: string): string {
	const lines = source.split(/\r?\n/);
	const header = lines.findIndex((line) => line.trim() === LAUNCHER_HEADER);
	lines.splice(header + 1, 0, OPTION_EXPLICIT);
	return lines.join('\r\n');
}

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
	if (existing.length === 0) {
		return `${LAUNCHER_HEADER}\r\n${OPTION_EXPLICIT}\r\n\r\n${subText}`;
	}
	// A module XLIDE wrote and nobody has edited gains the option it should
	// have had; anything else is left exactly as the user left it.
	const head = declaresOptionExplicit(existing) || !isUntouchedLauncherModule(existing)
		? existing
		: withOptionExplicit(existing);
	return `${head}\r\n\r\n${subText}`;
}
