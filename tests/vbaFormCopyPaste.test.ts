import { afterEach, describe, expect, it } from 'vitest';
import * as fs from 'fs';
import * as path from 'path';
import * as os from 'os';
import {
	duplicateFormControls,
	readFormMarkup,
	removeFormControls,
	resetProjectCacheForTests,
} from '../src/vba/projectService';

// Copy, paste, and deleting a multi-selection. All three are MARKUP
// transforms: a clone travels the same authoring path a hand-typed control
// takes, so a container's children come along and are renamed with it, and a
// multi-delete is one write and so one undo.

const FIXTURE = path.join('tests', 'fixtures', 'binaries', 'FormFixtureVbide.xlsm');

const tempDirs: string[] = [];
afterEach(() => {
	resetProjectCacheForTests();
	for (const dir of tempDirs.splice(0)) {
		fs.rmSync(dir, { recursive: true, force: true });
	}
});

function project(): string {
	const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'xlide-copy-'));
	tempDirs.push(dir);
	const wb = path.join(dir, 'Forms.xlsm');
	fs.copyFileSync(FIXTURE, wb);
	return wb;
}

/** Markup as the engine prints it back, read from the saved bytes. */
function markupOf(wb: string): string {
	const markup = readFormMarkup(wb, 'EntryForm').markup;
	resetProjectCacheForTests();
	return markup;
}

function lineFor(wb: string, name: string): string {
	const line = markupOf(wb).split('\r\n').find((l) => l.includes(`Name="${name}"`));
	if (line === undefined) { throw new Error(`no line for ${name}`); }
	return line;
}

function attr(line: string, key: string): string | undefined {
	return new RegExp(`\\s${key}="([^"]*)"`).exec(line)?.[1];
}

function names(wb: string): string[] {
	return [...markupOf(wb).matchAll(/<\w+ Name="([^"]+)"/g)].map((m) => m[1]);
}

describe('duplicateFormControls', () => {
	it('clones a control under a fresh name, nudged clear of the original', () => {
		const wb = project();
		const before = lineFor(wb, 'NameLabel');

		const result = duplicateFormControls(wb, 'EntryForm', ['NameLabel']);

		expect(result.newNames).toEqual(['Label1']);
		const copy = lineFor(wb, 'Label1');
		expect(Number(attr(copy, 'Left'))).toBe(Number(attr(before, 'Left')) + 6);
		expect(Number(attr(copy, 'Top'))).toBe(Number(attr(before, 'Top')) + 6);
		// Everything that is not identity or position came across.
		expect(attr(copy, 'Caption')).toBe(attr(before, 'Caption'));
		expect(attr(copy, 'Width')).toBe(attr(before, 'Width'));
		expect(attr(copy, 'Font.Name')).toBe(attr(before, 'Font.Name'));
		// And the original is still there, where it was.
		expect(lineFor(wb, 'NameLabel')).toBe(before);
	});

	it('takes a container\'s children with it, renamed, at their old offsets', () => {
		const wb = project();

		const result = duplicateFormControls(wb, 'EntryForm', ['Options']);

		expect(result.newNames).toEqual(['Frame1']);
		const after = markupOf(wb);
		const copy = after.split('\r\n')
			.slice(after.split('\r\n').findIndex((l) => l.includes('Name="Frame1"')));
		const children = copy.slice(1, copy.findIndex((l) => l.trim() === '</Frame>'));
		expect(children.map((l) => /Name="([^"]+)"/.exec(l)?.[1]))
			.toEqual(['OptionButton1', 'OptionButton2']);
		// A child's Left/Top is relative to its frame, so it must NOT move.
		expect(attr(children[0], 'Left')).toBe(attr(lineFor(wb, 'PickGround'), 'Left'));
		expect(attr(children[0], 'Top')).toBe(attr(lineFor(wb, 'PickGround'), 'Top'));
		expect(attr(children[0], 'Caption')).toBe('Ground');
	});

	it('duplicates a whole selection in one write, answering in the order asked', () => {
		const wb = project();

		const result = duplicateFormControls(wb, 'EntryForm', ['NameLabel', 'NameBox', 'OkButton']);

		expect(result.newNames).toEqual(['Label1', 'TextBox1', 'CommandButton1']);
		for (const name of result.newNames) {
			expect(names(wb)).toContain(name);
		}
	});

	it('refuses an unknown name and leaves the form alone', () => {
		const wb = project();
		const before = markupOf(wb);

		expect(() => duplicateFormControls(wb, 'EntryForm', ['NameLabel', 'Ghost']))
			.toThrow(/no control named Ghost/);

		expect(markupOf(wb)).toBe(before);
	});

	it('refuses to clone a Page: a page is added through the markup', () => {
		const wb = project();
		const before = markupOf(wb);

		expect(() => duplicateFormControls(wb, 'EntryForm', ['Page1'])).toThrow(/is a Page/);

		expect(markupOf(wb)).toBe(before);
	});

	it('does nothing at all when nothing was copied', () => {
		const wb = project();
		const before = markupOf(wb);

		expect(duplicateFormControls(wb, 'EntryForm', []).newNames).toEqual([]);

		expect(markupOf(wb)).toBe(before);
	});
});

describe('removeFormControls', () => {
	it('deletes a whole selection in one write', () => {
		const wb = project();

		const result = removeFormControls(wb, 'EntryForm', ['NameLabel', 'Taxable', 'HoldToggle']);

		expect(result.removed).toEqual(['NameLabel', 'Taxable', 'HoldToggle']);
		const left = names(wb);
		expect(left).not.toContain('NameLabel');
		expect(left).not.toContain('Taxable');
		expect(left).not.toContain('HoldToggle');
		expect(left).toContain('NameBox');
	});

	it('takes a container\'s children out with it', () => {
		const wb = project();

		removeFormControls(wb, 'EntryForm', ['Options']);

		const left = names(wb);
		expect(left).not.toContain('Options');
		expect(left).not.toContain('PickGround');
		expect(left).not.toContain('PickAir');
	});

	it('accepts a child named alongside its container: the parent already takes it', () => {
		const wb = project();

		const result = removeFormControls(wb, 'EntryForm', ['PickGround', 'Options']);

		expect(result.removed).toEqual(['Options']);
		expect(names(wb)).not.toContain('PickGround');
	});

	it('refuses an unknown name rather than half-deleting a stale selection', () => {
		const wb = project();
		const before = markupOf(wb);

		expect(() => removeFormControls(wb, 'EntryForm', ['NameLabel', 'Ghost']))
			.toThrow(/no control named Ghost/);

		expect(markupOf(wb)).toBe(before);
	});

	it('refuses to delete a Page', () => {
		const wb = project();
		const before = markupOf(wb);

		expect(() => removeFormControls(wb, 'EntryForm', ['Page1'])).toThrow(/is a Page/);

		expect(markupOf(wb)).toBe(before);
	});

	it('pastes and then deletes the copies, leaving the form as it started', () => {
		const wb = project();
		const before = markupOf(wb);

		const pasted = duplicateFormControls(wb, 'EntryForm', ['NameLabel', 'Options']);
		expect(names(wb)).toContain('Frame1');
		removeFormControls(wb, 'EntryForm', pasted.newNames);

		expect(markupOf(wb)).toBe(before);
	});
});
