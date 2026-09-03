import { beforeEach, describe, expect, it } from 'vitest';
import {
	activeFormLaunchTarget,
	lastFormLaunchTarget,
	resetFormLaunchTargetForTests,
	setActiveFormDesigner,
} from '../src/vbaFormLaunchTarget';

// F5 belongs to the designer on screen. VS Code keeps reporting the last text
// editor the user touched while a webview has focus, so the canvas has to say
// so itself, and two open designers must not clear each other's claim.
describe('which designer F5 belongs to', () => {
	const excel = { projectPath: 'C:\\work\\Book.xlsm', moduleName: 'EntryForm' };
	const vb6 = { projectPath: 'C:\\work\\Gallery.vbp', moduleName: 'frmGallery' };
	const one = {};
	const two = {};

	beforeEach(() => resetFormLaunchTargetForTests());

	it('has nothing on screen until a designer says so', () => {
		expect(activeFormLaunchTarget()).toBeUndefined();
		expect(lastFormLaunchTarget()).toBeUndefined();
	});

	it('reports the designer that holds the screen, and remembers it after it goes', () => {
		setActiveFormDesigner(one, () => vb6);
		expect(activeFormLaunchTarget()).toEqual(vb6);
		setActiveFormDesigner(one, undefined);
		expect(activeFormLaunchTarget()).toBeUndefined();
		expect(lastFormLaunchTarget()).toEqual(vb6);
	});

	it('lets a panel taking focus keep it when the one it replaced reports later', () => {
		setActiveFormDesigner(one, () => excel);
		setActiveFormDesigner(two, () => vb6);
		// The Excel designer now reports that it lost focus; the VB6 one keeps it.
		setActiveFormDesigner(one, undefined);
		expect(activeFormLaunchTarget()).toEqual(vb6);
	});

	it('clears only its own claim on disposal', () => {
		setActiveFormDesigner(one, () => vb6);
		setActiveFormDesigner(two, undefined);
		expect(activeFormLaunchTarget()).toEqual(vb6);
		setActiveFormDesigner(one, undefined);
		expect(activeFormLaunchTarget()).toBeUndefined();
	});
	it('answers when F5 asks, not when the panel opened', () => {
		// A VB6 designer restored with the window opens before the project
		// locator has scanned, so its form belongs to nothing yet. Asking
		// again later is what makes F5 work on that designer at all.
		let known: typeof vb6 | undefined;
		setActiveFormDesigner(one, () => known);
		expect(activeFormLaunchTarget()).toBeUndefined();
		known = vb6;
		expect(activeFormLaunchTarget()).toEqual(vb6);
		// And once it has answered, it is what a later F5 elsewhere falls back to.
		setActiveFormDesigner(one, undefined);
		expect(lastFormLaunchTarget()).toEqual(vb6);
	});
});
