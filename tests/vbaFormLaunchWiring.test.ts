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
