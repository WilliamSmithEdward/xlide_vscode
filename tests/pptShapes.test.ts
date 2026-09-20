import { describe, expect, it } from 'vitest';
import * as fs from 'fs';
import * as path from 'path';
import { ZipArchive } from '../src/vba/zip';
import { editSlideShape, listSlideShapes, presentationSlides, requireSlide } from '../src/vba/pptShapes';
import type { ShapeInfo } from '../src/vba/shapes';

// PowerPoint 16 saved the fixture. Slide 1 has ClickMe, a rectangle running
// Macros.SayHello; Caption, a text box; and Badge, an oval. Slide 2 has one
// rectangle, so per-slide addressing is exercised. Every edit here was also
// checked by having PowerPoint open the file the engine wrote and report its
// shapes (scripts in the session's verification run).
const FIXTURES = path.join(__dirname, 'fixtures', 'binaries');
const FIXTURE = 'PowerPointShapesFixture.pptm';

const open = (name = FIXTURE): ZipArchive => ZipArchive.read(fs.readFileSync(path.join(FIXTURES, name)));
const reopen = (zip: ZipArchive): ZipArchive => ZipArchive.read(zip.toBytes());
const slide = (zip: ZipArchive, surface: string) => requireSlide(presentationSlides(zip), surface);
const shapesOn = (zip: ZipArchive, surface: string): ShapeInfo[] => listSlideShapes(zip, slide(zip, surface));
const named = (zip: ZipArchive, surface: string, name: string): ShapeInfo | undefined =>
	shapesOn(zip, surface).find((s) => s.name === name);
const part = (zip: ZipArchive, name: string): string => zip.read(name).toString('utf8');

describe('finding slides', () => {
	it('orders them by the presentation, not by the slideN file numbering', () => {
		expect(presentationSlides(open())).toEqual([
			{ index: 1, name: 'Slide 1', path: 'ppt/slides/slide1.xml' },
			{ index: 2, name: 'Slide 2', path: 'ppt/slides/slide2.xml' },
		]);
	});

	it('takes a position spelled any of the ways a caller might spell it', () => {
		const zip = open();
		for (const surface of ['Slide 2', 'slide 2', 'Slide2', '2']) {
			expect(slide(zip, surface).path, surface).toBe('ppt/slides/slide2.xml');
		}
	});

	it('says what there is when the slide is not there', () => {
		expect(() => slide(open(), 'Slide 9')).toThrow(/has 2 slides; there is no slide 9/);
		expect(() => slide(open(), 'Summary')).toThrow(/No slide named 'Summary'/);
	});
});

describe('listing slide shapes', () => {
	it('reports each shape with its kind, macro, text and position in points', () => {
		expect(shapesOn(open(), 'Slide 1')).toEqual([
			{
				name: 'ClickMe', kind: 'shape', geometry: 'rect',
				left: 60, top: 60, width: 220, height: 90,
				macro: 'Macros.SayHello', text: 'Run the macro',
			},
			{
				name: 'Caption', kind: 'textBox', left: 60, top: 190, width: 260, height: 29.08,
				text: 'A caption with no macro',
			},
			{ name: 'Badge', kind: 'shape', geometry: 'ellipse', left: 330, top: 60, width: 120, height: 120 },
		]);
	});

	it('keeps the slides apart', () => {
		expect(shapesOn(open(), 'Slide 2').map((s) => s.name)).toEqual(['SecondSlideShape', 'Pair']);
	});

	it('reports a placeholder at the box it inherits from its layout', () => {
		// PowerPointPlaceholderFixture has a title slide and a text slide, and
		// neither placeholder carries an a:xfrm: PowerPoint reports where the
		// LAYOUT puts it. Every number below is what PowerPoint answered for
		// this file, so a reader that stopped at the slide part reports no
		// position at all for the title of any deck that is not blank-layout.
		const zip = open('PowerPointPlaceholderFixture.pptm');
		const say = (surface: string) => shapesOn(zip, surface)
			.map((s) => `${s.name}: ${s.kind} ${s.width}x${s.height} at ${s.left},${s.top}`);

		expect(say('Slide 1')).toEqual([
			'Title 1: placeholder 720x188 at 120,88.38',
			'Subtitle 2: placeholder 720x130.37 at 120,283.63',
			'Plain: shape 120x60 at 40,400',
		]);
		expect(say('Slide 2')).toEqual([
			'Title 1: placeholder 828x104.38 at 66,28.75',
			'Text Placeholder 2: placeholder 828x342.63 at 66,143.75',
		]);
	});

	it("reports a group's own box, not the box of the shape inside it", () => {
		// PowerPoint gave Pair the same a:off as its first member and a wider
		// a:ext, so a reader that searched the subtree for a p:spPr would get
		// the position right and the size wrong.
		expect(shapesOn(open(), 'Slide 2')[1]).toEqual({
			name: 'Pair', kind: 'group', left: 300, top: 200, width: 230, height: 140,
			shapes: [
				{ name: 'GroupedRect', kind: 'shape', geometry: 'rect', left: 300, top: 200, width: 100, height: 50 },
				{ name: 'GroupedOval', kind: 'shape', geometry: 'ellipse', left: 450, top: 260, width: 80, height: 80 },
			],
		});
	});
});

describe('linking a slide shape to a macro', () => {
	it('writes the action URI PowerPoint reads, with an empty relationship id', () => {
		const zip = open();
		editSlideShape(zip, slide(zip, 'Slide 1'), { action: 'update', name: 'Badge', macro: 'Unlinked' });

		expect(part(zip, 'ppt/slides/slide1.xml')).toContain(
			'<a:hlinkClick r:id="" action="ppaction://macro?name=Unlinked"/>');
		expect(named(reopen(zip), 'Slide 1', 'Badge')?.macro).toBe('Unlinked');
	});

	it('puts the link first in cNvPr, before the extLst already there', () => {
		const zip = open();
		editSlideShape(zip, slide(zip, 'Slide 1'), { action: 'update', name: 'Caption', macro: 'Macros.SayHello' });

		const cNvPr = /<p:cNvPr id="3"[\s\S]*?<\/p:cNvPr>/.exec(part(zip, 'ppt/slides/slide1.xml'))![0];
		expect(cNvPr.indexOf('<a:hlinkClick')).toBeLessThan(cNvPr.indexOf('<a:extLst>'));
	});

	it('delinks on an empty macro, and leaves the shape otherwise as it was', () => {
		const zip = open();
		editSlideShape(zip, slide(zip, 'Slide 1'), { action: 'update', name: 'ClickMe', macro: '' });

		const shape = named(reopen(zip), 'Slide 1', 'ClickMe');
		expect(shape?.macro).toBeUndefined();
		expect(shape?.text).toBe('Run the macro');
		expect(part(zip, 'ppt/slides/slide1.xml')).not.toContain('hlinkClick');
	});

	it('refuses a group, which PowerPoint runs no macro for', () => {
		const zip = open();
		expect(() => editSlideShape(zip, slide(zip, 'Slide 2'), { action: 'update', name: 'Pair', macro: 'Macros.SayHello' }))
			.toThrow(/'Pair' is a group, which PowerPoint runs no macro for/);
	});

	it('names the shape that is missing, rather than failing vaguely', () => {
		const zip = open();
		expect(() => editSlideShape(zip, slide(zip, 'Slide 1'), { action: 'update', name: 'Missing', macro: 'X' }))
			.toThrow(/No shape named 'Missing' on Slide 1/);
	});

	it('links a shape inside a group, which PowerPoint does run', () => {
		const zip = open();
		editSlideShape(zip, slide(zip, 'Slide 2'), { action: 'update', name: 'GroupedOval', macro: 'Unlinked' });

		const pair = shapesOn(reopen(zip), 'Slide 2').find((s) => s.name === 'Pair');
		expect(pair?.shapes?.map((s) => s.macro)).toEqual([undefined, 'Unlinked']);
	});
});

describe('changing a slide shape', () => {
	it('moves and resizes it, in points', () => {
		const zip = open();
		editSlideShape(zip, slide(zip, 'Slide 1'), { action: 'update', name: 'Badge', left: 10, top: 20, width: 30, height: 40 });

		expect(named(reopen(zip), 'Slide 1', 'Badge')).toMatchObject({ left: 10, top: 20, width: 30, height: 40 });
	});

	it('moves without resizing when only left and top are given', () => {
		const zip = open();
		editSlideShape(zip, slide(zip, 'Slide 1'), { action: 'update', name: 'Badge', left: 5, top: 6 });

		expect(named(reopen(zip), 'Slide 1', 'Badge')).toMatchObject({ left: 5, top: 6, width: 120, height: 120 });
	});

	it('sets text, alt text and a new name in one edit', () => {
		const zip = open();
		editSlideShape(zip, slide(zip, 'Slide 1'), {
			action: 'update', name: 'Caption', text: 'line one\nline two', altText: 'a caption', newName: 'Note',
		});

		expect(named(reopen(zip), 'Slide 1', 'Note')).toMatchObject({
			text: 'line one\nline two', altText: 'a caption',
		});
	});

	it('refuses a rename onto a name the slide already uses', () => {
		const zip = open();
		expect(() => editSlideShape(zip, slide(zip, 'Slide 1'), { action: 'update', name: 'Badge', newName: 'ClickMe' }))
			.toThrow(/already has a shape named 'ClickMe'/);
	});

	it('refuses the Excel-only properties rather than ignoring them', () => {
		const zip = open();
		const on = slide(zip, 'Slide 1');
		expect(() => editSlideShape(zip, on, { action: 'update', name: 'Badge', range: 'B2:D4' }))
			.toThrow(/A slide has no cells/);
		expect(() => editSlideShape(zip, on, { action: 'update', name: 'Badge', linkedCell: '$A$1' }))
			.toThrow(/Excel form-control properties/);
	});
});

describe('adding a slide shape', () => {
	it('names it as PowerPoint would and places it where it was asked', () => {
		const zip = open();
		const name = editSlideShape(zip, slide(zip, 'Slide 1'), {
			action: 'add', type: 'roundedRectangle', left: 100, top: 200, width: 80, height: 40,
		});

		expect(name).toBe('Rectangle: Rounded Corners 4');
		expect(named(reopen(zip), 'Slide 1', name)).toEqual({
			name, kind: 'shape', geometry: 'roundRect', left: 100, top: 200, width: 80, height: 40,
		});
	});

	it('adds a text box with a macro and text in one call', () => {
		const zip = open();
		const name = editSlideShape(zip, slide(zip, 'Slide 2'), {
			action: 'add', type: 'textBox', name: 'Footer', left: 20, top: 400, text: 'Click me', macro: 'Macros.SayHello',
		});
		expect(name).toBe('Footer');

		expect(named(reopen(zip), 'Slide 2', name)).toEqual({
			name: 'Footer', kind: 'textBox', left: 20, top: 400, width: 120, height: 60,
			text: 'Click me', macro: 'Macros.SayHello',
		});
	});

	it('adds it last, so it lands on top as a new shape does', () => {
		const zip = open();
		editSlideShape(zip, slide(zip, 'Slide 1'), { action: 'add', type: 'oval', name: 'Top', left: 1, top: 1 });

		expect(shapesOn(reopen(zip), 'Slide 1').map((s) => s.name)).toEqual(['ClickMe', 'Caption', 'Badge', 'Top']);
	});

	it('refuses what a slide cannot hold, and says what to do instead', () => {
		const zip = open();
		const on = slide(zip, 'Slide 1');
		expect(() => editSlideShape(zip, on, { action: 'add', type: 'button', left: 0, top: 0 }))
			.toThrow(/button is an Excel form control/);
		expect(() => editSlideShape(zip, on, { action: 'add', type: 'oval' }))
			.toThrow(/needs left and top/);
		expect(() => editSlideShape(zip, on, { action: 'add', type: 'oval', left: 0, top: 0, width: 0 }))
			.toThrow(/width and a height above zero/);
	});

	it('keeps text with markup characters as typed', () => {
		const zip = open();
		const name = editSlideShape(zip, slide(zip, 'Slide 1'), {
			action: 'add', type: 'rectangle', name: 'Raw', left: 0, top: 0, text: 'a < b & "c" > d',
		});

		expect(named(reopen(zip), 'Slide 1', name)?.text).toBe('a < b & "c" > d');
	});
});

describe('deleting a slide shape', () => {
	it('removes it and leaves the others alone', () => {
		const zip = open();
		editSlideShape(zip, slide(zip, 'Slide 1'), { action: 'delete', name: 'Caption' });

		expect(shapesOn(reopen(zip), 'Slide 1').map((s) => s.name)).toEqual(['ClickMe', 'Badge']);
	});

	it('removes a group with everything in it', () => {
		const zip = open();
		editSlideShape(zip, slide(zip, 'Slide 2'), { action: 'delete', name: 'Pair' });

		const after = reopen(zip);
		expect(shapesOn(after, 'Slide 2').map((s) => s.name)).toEqual(['SecondSlideShape']);
		expect(part(after, 'ppt/slides/slide2.xml')).not.toContain('GroupedOval');
	});

	it('refuses to delete one shape out of a group', () => {
		const zip = open();
		expect(() => editSlideShape(zip, slide(zip, 'Slide 2'), { action: 'delete', name: 'GroupedRect' }))
			.toThrow(/is inside a group; ungroup it in PowerPoint, or delete the group/);
	});

	it('leaves every other part of the package byte for byte as it was', () => {
		const before = open();
		const after = open();
		editSlideShape(after, slide(after, 'Slide 1'), { action: 'delete', name: 'Badge' });

		for (const name of before.names()) {
			if (name === 'ppt/slides/slide1.xml') { continue; }
			expect(after.read(name).equals(before.read(name)), name).toBe(true);
		}
	});
});
