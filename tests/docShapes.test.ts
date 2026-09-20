import { describe, expect, it } from 'vitest';
import * as fs from 'fs';
import * as path from 'path';
import { ZipArchive } from '../src/vba/zip';
import {
	defaultStory,
	documentStories,
	editStoryShape,
	listStoryShapes,
	requireStory,
} from '../src/vba/docShapes';
import type { ShapeInfo } from '../src/vba/shapes';

// Word 16 saved the fixture. The body has InlineOval (inline with the text),
// AnchoredBox (floating, with text and alt text) and Board, a drawing canvas
// holding the group Pair over GroupedRect and GroupedOval. The primary
// header holds HeaderBadge - and Word put it in word/header2.xml, since
// header1.xml is the EVEN-page header. Every edit here was also checked by
// having Word open the file the engine wrote and report its shapes.
const FIXTURES = path.join(__dirname, 'fixtures', 'binaries');
const FIXTURE = 'WordShapesFixture.docm';

const open = (): ZipArchive => ZipArchive.read(fs.readFileSync(path.join(FIXTURES, FIXTURE)));
const reopen = (zip: ZipArchive): ZipArchive => ZipArchive.read(zip.toBytes());
const body = (zip: ZipArchive) => defaultStory(documentStories(zip));
const story = (zip: ZipArchive, surface: string) => requireStory(documentStories(zip), surface);
const shapesOn = (zip: ZipArchive, surface?: string): ShapeInfo[] =>
	listStoryShapes(zip, surface === undefined ? body(zip) : story(zip, surface));
const named = (zip: ZipArchive, name: string, surface?: string): ShapeInfo | undefined =>
	shapesOn(zip, surface).find((s) => s.name === name);
const part = (zip: ZipArchive, name: string): string => zip.read(name).toString('utf8');

describe('finding surfaces', () => {
	it('lists the body and each header and footer, named as Word names them', () => {
		expect(documentStories(open()).map((s) => `${s.name} -> ${s.path}`)).toEqual([
			'Document -> word/document.xml',
			'Even Page Header -> word/header1.xml',
			'Header -> word/header2.xml',
			'Even Page Footer -> word/footer1.xml',
			'Footer -> word/footer2.xml',
			'First Page Header -> word/header3.xml',
			'First Page Footer -> word/footer3.xml',
			'Footnotes -> word/footnotes.xml',
			'Endnotes -> word/endnotes.xml',
		]);
	});

	it('takes the part file name too, since the numbering is not the order', () => {
		expect(story(open(), 'header2').name).toBe('Header');
		expect(story(open(), 'Header').path).toBe('word/header2.xml');
	});

	it('says what there is when the surface is not there', () => {
		expect(() => story(open(), 'Slide 1')).toThrow(/No surface named 'Slide 1'.*'Document'/s);
	});
});

describe('listing document shapes', () => {
	it('reports inline and floating shapes with their size and placement', () => {
		const shapes = shapesOn(open());

		expect(shapes.slice(0, 2)).toEqual([
			{ name: 'InlineOval', kind: 'shape', geometry: 'ellipse', placement: 'inline', width: 120, height: 120 },
			{
				name: 'AnchoredBox', kind: 'shape', geometry: 'rect', placement: 'anchored',
				left: 8, top: 8, width: 180, height: 70, text: 'Anchored', altText: 'a described box',
			},
		]);
	});

	it('reports a canvas as a container, with the group and its members inside', () => {
		const board = named(open(), 'Board');

		expect(board?.kind).toBe('canvas');
		expect(board?.shapes?.map((s) => s.name)).toEqual(['Pair']);
		expect(board?.shapes?.[0].shapes?.map((s) => `${s.name}:${s.geometry}`))
			.toEqual(['GroupedRect:rect', 'GroupedOval:ellipse']);
	});

	it('finds the shape in the header, which is not word/header1.xml', () => {
		expect(shapesOn(open(), 'Header').map((s) => s.name)).toEqual(['HeaderBadge']);
		expect(shapesOn(open(), 'Even Page Header')).toEqual([]);
	});
});

describe('what Word cannot do', () => {
	it('refuses a macro, and says where Word does keep one', () => {
		const zip = open();
		expect(() => editStoryShape(zip, body(zip), { action: 'update', name: 'AnchoredBox', macro: 'SayHello' }))
			.toThrow(/Word cannot run a macro from a shape.*EntryMacro/s);
	});

	it('refuses the Excel-only properties rather than ignoring them', () => {
		const zip = open();
		expect(() => editStoryShape(zip, body(zip), { action: 'update', name: 'AnchoredBox', range: 'B2:D4' }))
			.toThrow(/A document has no cells/);
	});

	it('refuses to move an inline shape, which Word positions by the text', () => {
		const zip = open();
		expect(() => editStoryShape(zip, body(zip), { action: 'update', name: 'InlineOval', left: 40 }))
			.toThrow(/is inline with the text/);
	});

	it('resizes an inline shape, which Word does allow', () => {
		const zip = open();
		editStoryShape(zip, body(zip), { action: 'update', name: 'InlineOval', width: 60, height: 40 });

		expect(named(reopen(zip), 'InlineOval')).toMatchObject({ width: 60, height: 40, placement: 'inline' });
	});
});

describe('changing a document shape', () => {
	it('moves and resizes a floating shape, in points', () => {
		const zip = open();
		editStoryShape(zip, body(zip), { action: 'update', name: 'AnchoredBox', left: 40, top: 50, width: 200, height: 90 });

		expect(named(reopen(zip), 'AnchoredBox')).toMatchObject({ left: 40, top: 50, width: 200, height: 90 });
	});

	it('keeps the VML twin in step, so both copies say the same thing', () => {
		const zip = open();
		editStoryShape(zip, body(zip), { action: 'update', name: 'AnchoredBox', left: 40, top: 50, width: 200, height: 90 });

		const vml = /<v:rect\b[^>]*\bstyle="([^"]*)"/.exec(part(zip, 'word/document.xml'))![1];
		expect(vml).toContain('margin-left:40pt');
		expect(vml).toContain('margin-top:50pt');
		expect(vml).toContain('width:200pt');
		expect(vml).toContain('height:90pt');
	});

	it('sets text in both copies at once', () => {
		const zip = open();
		editStoryShape(zip, body(zip), { action: 'update', name: 'AnchoredBox', text: 'new text' });

		const xml = part(zip, 'word/document.xml');
		expect(named(reopen(zip), 'AnchoredBox')?.text).toBe('new text');
		expect([...xml.matchAll(/<w:t[^>]*>new text<\/w:t>/g)]).toHaveLength(2);
		expect(xml).not.toContain('>Anchored<');
	});

	it('renames in both copies, since the VML twin names it with its id', () => {
		const zip = open();
		editStoryShape(zip, body(zip), { action: 'update', name: 'AnchoredBox', newName: 'Callout' });

		const xml = part(zip, 'word/document.xml');
		expect(named(reopen(zip), 'Callout')).toBeDefined();
		expect(xml).toMatch(/<v:rect\b[^>]*\bid="Callout"/);
		expect(xml).not.toContain('AnchoredBox');
	});

	it('changes a shape inside the canvas without disturbing its sibling', () => {
		const zip = open();
		editStoryShape(zip, body(zip), { action: 'update', name: 'GroupedOval', width: 30, height: 30, newName: 'Dot' });

		// A canvas member's a:off is inside the canvas, not on the page.
		const pair = named(reopen(zip), 'Board')?.shapes?.[0];
		expect(pair?.shapes).toEqual([
			{ name: 'GroupedRect', kind: 'shape', geometry: 'rect', left: 10, top: 10, width: 100, height: 50 },
			{ name: 'Dot', kind: 'shape', geometry: 'ellipse', left: 150, top: 30, width: 30, height: 30 },
		]);
	});

	it('removes alt text on an empty string', () => {
		const zip = open();
		editStoryShape(zip, body(zip), { action: 'update', name: 'AnchoredBox', altText: '' });

		expect(named(reopen(zip), 'AnchoredBox')?.altText).toBeUndefined();
	});
});

describe('adding a document shape', () => {
	it('adds a floating shape when it is given a position', () => {
		const zip = open();
		const name = editStoryShape(zip, body(zip), {
			action: 'add', type: 'rectangle', name: 'Stamp', left: 100, top: 200, width: 150, height: 60, text: 'Draft',
		});

		expect(named(reopen(zip), name)).toEqual({
			name: 'Stamp', kind: 'shape', geometry: 'rect', placement: 'anchored',
			left: 100, top: 200, width: 150, height: 60, text: 'Draft',
		});
	});

	it('adds an inline shape when it is given none', () => {
		const zip = open();
		const name = editStoryShape(zip, body(zip), { action: 'add', type: 'oval', name: 'Bullet', width: 20, height: 20 });

		expect(named(reopen(zip), name)).toEqual({
			name: 'Bullet', kind: 'shape', geometry: 'ellipse', placement: 'inline', width: 20, height: 20,
		});
	});

	it('adds to a header as readily as to the body', () => {
		const zip = open();
		editStoryShape(zip, story(zip, 'Header'), { action: 'add', type: 'textBox', name: 'Tagline', width: 200, height: 20, text: 'Confidential' });

		expect(shapesOn(reopen(zip), 'Header').map((s) => s.name)).toEqual(['HeaderBadge', 'Tagline']);
		expect(shapesOn(reopen(zip)).some((s) => s.name === 'Tagline'), 'leaked into the body').toBe(false);
	});

	it('keeps a floating shape out of the text, and gives an inline one its own paragraph', () => {
		const zip = open();
		const before = (part(zip, 'word/document.xml').match(/<w:p[ >]/g) ?? []).length;
		editStoryShape(zip, body(zip), { action: 'add', type: 'oval', name: 'Floater', left: 10, top: 10 });
		const afterFloating = (part(zip, 'word/document.xml').match(/<w:p[ >]/g) ?? []).length;
		editStoryShape(zip, body(zip), { action: 'add', type: 'oval', name: 'Inliner' });
		const afterInline = (part(zip, 'word/document.xml').match(/<w:p[ >]/g) ?? []).length;

		expect(afterFloating, 'a floating shape should not add a paragraph').toBe(before);
		expect(afterInline).toBe(before + 1);
	});

	it('refuses what Word cannot hold, and says what to do instead', () => {
		const zip = open();
		const on = body(zip);
		expect(() => editStoryShape(zip, on, { action: 'add', type: 'button' }))
			.toThrow(/button is an Excel form control/);
		expect(() => editStoryShape(zip, on, { action: 'add', type: 'oval', width: 0 }))
			.toThrow(/width and a height above zero/);
		expect(() => editStoryShape(zip, on, { action: 'add', type: 'oval', name: 'AnchoredBox' }))
			.toThrow(/already has a shape named 'AnchoredBox'/);
	});

	it('keeps text with markup characters as typed', () => {
		const zip = open();
		const name = editStoryShape(zip, body(zip), {
			action: 'add', type: 'textBox', name: 'Raw', text: 'a < b & "c" > d',
		});

		expect(named(reopen(zip), name)?.text).toBe('a < b & "c" > d');
	});
});

describe('deleting a document shape', () => {
	it('removes the shape and the VML twin with it', () => {
		const zip = open();
		editStoryShape(zip, body(zip), { action: 'delete', name: 'AnchoredBox' });

		const xml = part(zip, 'word/document.xml');
		expect(shapesOn(reopen(zip)).map((s) => s.name)).toEqual(['InlineOval', 'Board']);
		expect(xml).not.toContain('AnchoredBox');
		expect(xml).not.toMatch(/<v:rect\b[^>]*\bid="AnchoredBox"/);
	});

	it('keeps the text of a run that held more than the shape', () => {
		const zip = open();
		editStoryShape(zip, body(zip), { action: 'delete', name: 'InlineOval' });

		expect(part(zip, 'word/document.xml')).toContain('A document with shapes.');
	});

	it('removes a canvas with everything in it', () => {
		const zip = open();
		editStoryShape(zip, body(zip), { action: 'delete', name: 'Board' });

		expect(shapesOn(reopen(zip)).map((s) => s.name)).toEqual(['InlineOval', 'AnchoredBox']);
		expect(part(zip, 'word/document.xml')).not.toContain('GroupedOval');
	});

	it('refuses to delete one shape out of a canvas, and names the container', () => {
		const zip = open();
		expect(() => editStoryShape(zip, body(zip), { action: 'delete', name: 'GroupedRect' }))
			.toThrow(/'GroupedRect' is inside 'Board'/);
	});

	it('leaves every other part of the package byte for byte as it was', () => {
		const before = open();
		const after = open();
		editStoryShape(after, body(after), { action: 'delete', name: 'AnchoredBox' });

		for (const name of before.names()) {
			if (name === 'word/document.xml') { continue; }
			expect(after.read(name).equals(before.read(name)), name).toBe(true);
		}
	});
});
