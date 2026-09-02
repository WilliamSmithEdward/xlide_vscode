// What F5 runs when the focus is on a canvas rather than on a text editor.
//
// Both designers remember the form they are showing here: the MSForms one
// over a project's storage, the VB6 one over a form's own file. F5 from a
// canvas has no active text editor to read, so the launcher asks this
// instead. Kept in its own module so neither designer has to import the
// other.

export interface FormLaunchTarget {
	/** The container: a project path, or a VB6 project's `.vbp`. */
	projectPath: string;
	/** The form the designer is showing. */
	moduleName: string;
}

let lastFocused: FormLaunchTarget | undefined;

/** Records the designer that just took focus. */
export function rememberFormLaunchTarget(target: FormLaunchTarget): void {
	lastFocused = target;
}

/** The last designer to take focus, or undefined when none has. */
export function lastFormLaunchTarget(): FormLaunchTarget | undefined {
	return lastFocused;
}

/** Forgets the target, for a test that must start from nothing. */
export function resetFormLaunchTargetForTests(): void {
	lastFocused = undefined;
}
