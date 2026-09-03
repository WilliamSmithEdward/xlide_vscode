// What F5 runs when the focus is on a canvas rather than on a text editor.
//
// Both designers report the form they are showing here: the MSForms one over
// a project's storage, the VB6 one over a form's own file. A canvas has no
// text editor of its own, and `window.activeTextEditor` does NOT go quiet
// while a webview has focus - it keeps naming the last text editor the user
// touched, which may belong to an entirely different project. So a designer
// that is on screen says so, and the launcher believes it before it believes
// the editor. Kept in its own module so neither designer imports the other.

export interface FormLaunchTarget {
	/** The container: a project path, or a VB6 project's `.vbp`. */
	projectPath: string;
	/** The form the designer is showing. */
	moduleName: string;
}

/**
 * How a designer names the form it shows. A function, not a value, because
 * what a VB6 form belongs to is only known once the project locator has
 * scanned the workspace, which finishes after activation - and a designer
 * restored with the window opens before that. Answering the question when
 * F5 asks it, rather than when the panel opened, removes the race.
 */
export type FormLaunchTargetSource = () => FormLaunchTarget | undefined;

let lastFocused: FormLaunchTarget | undefined;
/** The designer on screen right now, with the panel that claimed it. */
let onScreen: { owner: object; resolve: FormLaunchTargetSource } | undefined;

/**
 * Reports whether this designer holds the screen. `owner` is the panel's own
 * identity, so a panel losing focus clears only its own claim and never the
 * claim of the panel that just took focus from it.
 */
export function setActiveFormDesigner(owner: object, resolve: FormLaunchTargetSource | undefined): void {
	if (resolve) {
		onScreen = { owner, resolve };
		const target = resolve();
		if (target) {
			lastFocused = target;
		}
		return;
	}
	if (onScreen?.owner === owner) {
		onScreen = undefined;
	}
}

/** The designer on screen, which F5 belongs to whatever the text editors say. */
export function activeFormLaunchTarget(): FormLaunchTarget | undefined {
	const target = onScreen?.resolve();
	if (target) {
		lastFocused = target;
	}
	return target;
}

/** The last designer to have been on screen, for an F5 from somewhere else. */
export function lastFormLaunchTarget(): FormLaunchTarget | undefined {
	return lastFocused;
}

/** Forgets both, for a test that must start from nothing. */
export function resetFormLaunchTargetForTests(): void {
	lastFocused = undefined;
	onScreen = undefined;
}
