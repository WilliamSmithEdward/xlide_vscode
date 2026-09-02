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
		setActiveFormDesigner(one, vb6);
		expect(activeFormLaunchTarget()).toEqual(vb6);
		setActiveFormDesigner(one, undefined);
		expect(activeFormLaunchTarget()).toBeUndefined();
		expect(lastFormLaunchTarget()).toEqual(vb6);
	});

	it('lets a panel taking focus keep it when the one it replaced reports later', () => {
		setActiveFormDesigner(one, excel);
		setActiveFormDesigner(two, vb6);
		// The Excel designer now reports that it lost focus; the VB6 one keeps it.
		setActiveFormDesigner(one, undefined);
		expect(activeFormLaunchTarget()).toEqual(vb6);
	});

	it('clears only its own claim on disposal', () => {
		setActiveFormDesigner(one, vb6);
		setActiveFormDesigner(two, undefined);
		expect(activeFormLaunchTarget()).toEqual(vb6);
		setActiveFormDesigner(one, undefined);
		expect(activeFormLaunchTarget()).toBeUndefined();
	});
});
