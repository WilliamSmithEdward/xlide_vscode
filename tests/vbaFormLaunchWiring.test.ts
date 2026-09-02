import { readFileSync } from 'fs';
import * as path from 'path';
import { describe, expect, it } from 'vitest';
import {
	composeLauncherSource,
	launcherSubExists,
	launcherSubName,
	LAUNCHER_MODULE,
} from '../src/vbaFormLauncher';

// F5 shows a form by running a macro XLIDE adds with consent. One sub per
// form, all in one module: running a second form must not cost the first its
// launcher, and a form whose sub is already there needs no consent and no
// write at all.
describe('the F5 launcher module', () => {
	it('names a sub per form, so two forms never share one entry point', () => {
		expect(LAUNCHER_MODULE).toBe('XlideRun');
		expect(launcherSubName('EntryForm')).toBe('XlideShow_EntryForm');
		expect(launcherSubName('OrderForm')).not.toBe(launcherSubName('EntryForm'));
	});

	it('creates the module with a header and this form-s sub', () => {
		const source = composeLauncherSource(undefined, 'EntryForm');
		expect(source).toContain('Safe to delete');
		expect(source).toContain('Sub XlideShow_EntryForm()');
		expect(source).toContain('UserForms.Add("EntryForm").Show');
		expect(source.trimEnd().endsWith('End Sub')).toBe(true);
	});

	it('ADDS the second form-s sub, keeping the first and any hand edits', () => {
		const first = composeLauncherSource(undefined, 'EntryForm');
		const both = composeLauncherSource(`${first}\r\n' a hand written note\r\n`, 'OrderForm');
		expect(both).toContain('Sub XlideShow_EntryForm()');
		expect(both).toContain('Sub XlideShow_OrderForm()');
		expect(both).toContain('a hand written note');
		// One header only, however many forms accumulate.
		expect(both.split('Safe to delete').length - 1).toBe(1);
	});

	it('recognises an installed sub - that is what suppresses the prompt', () => {
		const source = composeLauncherSource(undefined, 'EntryForm');
		expect(launcherSubExists(source, 'XlideShow_EntryForm')).toBe(true);
		expect(launcherSubExists(source, 'XlideShow_OrderForm')).toBe(false);
		expect(launcherSubExists(undefined, 'XlideShow_EntryForm')).toBe(false);
		// Whatever spelling the user left behind still counts as installed.
		expect(launcherSubExists('Public Sub XlideShow_EntryForm()\r\nEnd Sub\r\n', 'XlideShow_EntryForm')).toBe(true);
		expect(launcherSubExists('private sub xlideshow_entryform ()\r\nend sub\r\n', 'XlideShow_EntryForm')).toBe(true);
		// A near-miss name is NOT this form-s launcher.
		expect(launcherSubExists('Sub XlideShow_EntryForm2()\r\nEnd Sub\r\n', 'XlideShow_EntryForm')).toBe(false);
		// A mention in a comment or a call is not a declaration.
		expect(launcherSubExists("' XlideShow_EntryForm is gone\r\n", 'XlideShow_EntryForm')).toBe(false);
	});
});

// The command itself cannot run outside a real extension host, so its wiring
// is pinned at the source: both halves of the launch must sit inside the
// reopen suppression, the write must be skipped when the sub is already
// there, and the prompt must be skipped for an installed launcher.
const LAUNCH_SOURCE = readFileSync(path.join(__dirname, '..', 'src', 'vbaFormPreview.ts'), 'utf8');

describe('the F5 form launch', () => {
	it('suppresses the reopen across BOTH the write and the macro run', () => {
		const suppress = LAUNCH_SOURCE.indexOf('withWorkbookReopenSuppressed(');
		const write = LAUNCH_SOURCE.indexOf('runWriteWithExcelCoordination(');
		const run = LAUNCH_SOURCE.indexOf('runWorkbookMacroReadOnly(');
		expect(suppress).toBeGreaterThan(-1);
		expect(write).toBeGreaterThan(suppress);
		expect(run).toBeGreaterThan(suppress);
	});

	it('saves the pending designer edits BEFORE running, and stops if that fails', () => {
		// F5 runs what you see: without this the macro shows the last saved
		// form. The save sits inside the reopen suppression, or its own
		// post-save reopen races the macro host.
		const suppress = LAUNCH_SOURCE.indexOf('withWorkbookReopenSuppressed(wbPath, () => savePendingLaunchEdits(');
		expect(suppress).toBeGreaterThan(-1);
		expect(LAUNCH_SOURCE).toContain('if (!saved) {');
		expect(LAUNCH_SOURCE.indexOf('const saved = excel'))
			.toBeLessThan(LAUNCH_SOURCE.indexOf('runWorkbookMacroReadOnly('));
		// The designer's own document is saved even with no active editor,
		// which is exactly the F5-from-the-canvas case.
		expect(LAUNCH_SOURCE).toContain('encodeFormMarkupUri(filePath, formModule).toString()');
	});

	it('does not prompt, and does not write, when the sub is already installed', () => {
		expect(LAUNCH_SOURCE).toContain("if (mode === 'ask' && subExists) { mode = 'once'; }");
		expect(LAUNCH_SOURCE).toContain('if (!subExists) {');
	});

	it('runs one launch at a time, so a second F5 cannot stack another', () => {
		expect(LAUNCH_SOURCE).toContain('if (launchInFlight) { return; }');
		expect(LAUNCH_SOURCE).toContain('launchInFlight = false;');
	});

	it('tracks the workbook the macro host reopened, so a later save can free it', () => {
		expect(LAUNCH_SOURCE).toContain('markWorkbookOpenedByXlide(wbPath)');
	});
});

// A VB6 project's host application is Visual Basic itself. F5 saves the
// files and hands the `.vbp` to whatever the shell has registered for it;
// XLIDE neither builds nor runs the project (Slice 6 of the VB6 roadmap),
// and nothing here knows about twinBASIC.
describe('F5 on a VB6 form', () => {
	const PACKAGE = JSON.parse(readFileSync(path.join(__dirname, '..', 'package.json'), 'utf8')) as {
		contributes: { keybindings: { command: string; key: string; when: string }[] };
	};
	const DESIGNER_SOURCE = readFileSync(path.join(__dirname, '..', 'src', 'vb6FormDesigner.ts'), 'utf8');

	it('is bound in the VB6 designer and on a VB6 module document', () => {
		const f5 = PACKAGE.contributes.keybindings.filter((k) => k.command === 'xlide.launchFormHost' && k.key === 'F5');
		expect(f5.some((k) => k.when.includes("activeCustomEditorId == 'xlideVb6FormDesigner'"))).toBe(true);
		expect(f5.some((k) => /resourceExtname =~ .*frm\|ctl\|pag/.test(k.when))).toBe(true);
		// The MSForms bindings stay as they were.
		expect(f5.some((k) => k.when.includes("activeCustomEditorId == 'xlideFormDesigner'"))).toBe(true);
		expect(f5.some((k) => k.when.includes("resourceExtname == '.form'"))).toBe(true);
	});

	it('knows which project a focused VB6 form belongs to', () => {
		// The canvas has no text editor, so it says when it is on screen; a
		// focused module file names its project through the locator.
		expect(DESIGNER_SOURCE).toContain('setActiveFormDesigner(');
		expect(DESIGNER_SOURCE).toContain('onDidChangeViewState');
		expect(LAUNCH_SOURCE).toContain('lastFormLaunchTarget()');
		expect(LAUNCH_SOURCE).toContain('isVb6ProjectPath(location.projectPath)');
	});

	it('gives the designer on screen the launch, whatever the text editors say', () => {
		// `activeTextEditor` keeps naming the last text editor even while a
		// canvas has focus, so an Excel module open in another tab would
		// otherwise take F5 away from the VB6 form being looked at.
		expect(LAUNCH_SOURCE).toContain('const onScreen = activeFormLaunchTarget();');
		expect(LAUNCH_SOURCE.indexOf('const onScreen = activeFormLaunchTarget();'))
			.toBeLessThan(LAUNCH_SOURCE.indexOf("active.document.uri.scheme === XLIDE_SCHEME"));
		// Both designers report their own focus, and release it when they go.
		for (const source of [LAUNCH_SOURCE, DESIGNER_SOURCE]) {
			expect(source).toContain('setActiveFormDesigner(panelOwner, undefined)');
			expect(source).toContain('e.webviewPanel.active ?');
		}
	});

	it('saves every dirty file of the project before opening it', () => {
		// A VB6 module IS its file, so there is no markup document to save;
		// the form's own save is also what writes its pending .frx records.
		expect(LAUNCH_SOURCE).toContain('if (isVb6ProjectPath(filePath)) {');
		expect(LAUNCH_SOURCE).toContain("moduleLocationOfDocument(doc)?.projectPath.toLowerCase() === wanted");
		expect(LAUNCH_SOURCE.indexOf('savePendingLaunchEdits(wbPath, formModule)'))
			.toBeLessThan(LAUNCH_SOURCE.indexOf("'xlide.openInOfficeApp'"));
	});

	it('opens the project rather than building it, and names no build tool', () => {
		expect(LAUNCH_SOURCE).toContain('in Visual Basic...');
		// The launcher-macro path is Excel's alone: a .vbp never reaches it.
		expect(LAUNCH_SOURCE).toContain("const excel = /\\.(xlsm|xlsb|xlam|xls)$/i.test(wbPath)");
		expect(LAUNCH_SOURCE).toContain('excel && formModule');
		expect(LAUNCH_SOURCE.toLowerCase()).not.toContain('twinbasic');
		expect(DESIGNER_SOURCE.toLowerCase()).not.toContain('twinbasic');
	});
});
