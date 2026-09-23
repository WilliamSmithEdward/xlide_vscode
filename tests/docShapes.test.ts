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
import type { ShapeEdit, ShapeInfo } from '../src/vba/shapes';

// Word 16 saved the fixture. The body has InlineOval (inline with the text),
// AnchoredBox (floating, with text and alt text) and Board, a drawing canvas
// holding the group Pair over GroupedRect and GroupedOval. The primary
// header holds HeaderBadge - and Word put it in word/header2.xml, since
// header1.xml is the EVEN-page header. Every edit here was also checked by
// having Word open the file the engine wrote and report its shapes.
const FIXTURES = path.join(__dirname, 'fixtures', 'binaries');
const FIXTURE = 'WordShapesFixture.docm';

const open = (name = FIXTURE): ZipArchive => ZipArchive.read(fs.readFileSync(path.join(FIXTURES, name)));
const reopen = (zip: ZipArchive): ZipArchive => ZipArchive.read(zip.toBytes());
const body = (zip: ZipArchive) => defaultStory(documentStories(zip));
const story = (zip: ZipArchive, surface: string) => requireStory(documentStories(zip), surface);
/**
 * A shape as most of these tests describe it: what it is, where, and what
 * it says. Its look and stacking have tests of their own below.
 */
function essentials(shape: ShapeInfo): ShapeInfo {
	const { fill: _fill, line: _line, font: _font, rotation: _rotation, zOrder: _zOrder, ...rest } = shape;
	return rest.shapes ? { ...rest, shapes: rest.shapes.map(essentials) } : rest;
}
/** Every shape on a surface, with its look, for the formatting tests. */
const looksOn = (zip: ZipArchive, surface?: string): ShapeInfo[] =>
	listStoryShapes(zip, surface === undefined ? body(zip) : story(zip, surface));
const shapesOn = (zip: ZipArchive, surface?: string): ShapeInfo[] => looksOn(zip, surface).map(essentials);
const named = (zip: ZipArchive, name: string, surface?: string): ShapeInfo | undefined =>
	shapesOn(zip, surface).find((s) => s.name === name);
/** A shape with its look, members of a canvas included. */
const looked = (zip: ZipArchive, name: string): ShapeInfo | undefined => {
	const all = (shapes: ShapeInfo[]): ShapeInfo[] => shapes.flatMap((s) => [s, ...all(s.shapes ?? [])]);
	return all(looksOn(zip)).find((s) => s.name === name);
};
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

describe('formatting a document shape', () => {
	// ShapesFormattedFixture is WordShapesFixture after Word, driven through
	// its object model, made exactly these changes and saved. AnchoredBox was
	// already the back of the floating shapes, so sending it back moved
	// nothing.
	const WORDS_EDITS: ShapeEdit[] = [
		{
			action: 'update', name: 'AnchoredBox', rotation: 30, zOrder: 'back',
			fill: { type: 'solid', color: '#FF0000', transparency: 25 },
			line: { type: 'solid', color: '#008000', weight: 2.5, dash: 'dash' },
			font: { name: 'Arial', size: 14, bold: true, italic: true, underline: true, color: '#0000FF' },
		},
		{ action: 'update', name: 'Board', hidden: true },
	];
	// ShapesArrangedFixture is WordShapesFixture after Word made these.
	const WORDS_ARRANGEMENT: ShapeEdit[] = [
		{ action: 'update', name: 'AnchoredBox', rotation: 90, zOrder: 'front' },
		{ action: 'update', name: 'GroupedRect', hidden: true },
		{ action: 'update', name: 'GroupedOval', fill: { type: 'solid', color: '#FFFF00' }, line: { type: 'none' }, rotation: 20 },
	];

	const documentOf = (zip: ZipArchive): string => part(zip, 'word/document.xml');
	/** The mc:AlternateContent holding a top-level shape. */
	const entryOf = (zip: ZipArchive, name: string): string => {
		const xml = documentOf(zip);
		const at = xml.indexOf(`name="${name}"`);
		const start = xml.lastIndexOf('<mc:AlternateContent>', at);
		return xml.slice(start, xml.indexOf('</mc:AlternateContent>', at) + '</mc:AlternateContent>'.length);
	};
	/** The first element of `name` in a string, whole. */
	const elementIn = (xml: string, name: string): string =>
		xml.slice(xml.indexOf(`<${name}`), xml.indexOf(`</${name}>`) + `</${name}>`.length);
	/** A VML element's start tag, found by its id, without the preview image Word caches in it. */
	const vmlTag = (zip: ZipArchive, id: string): string => {
		const xml = documentOf(zip);
		const at = xml.indexOf(`id="${id}"`);
		return xml.slice(xml.lastIndexOf('<v:', at), xml.indexOf('>', at) + 1).replace(/ o:gfxdata="[^"]*"/, '');
	};
	/** A floating shape's wp:anchor relativeHeight. */
	const heightOf = (zip: ZipArchive, name: string): number =>
		Number(/relativeHeight="(\d+)"/.exec(entryOf(zip, name))![1]);
	/** A top-level shape's wp:effectExtent, l t r b in EMU. */
	const effectOf = (zip: ZipArchive, name: string): number[] => {
		const tag = /<wp:effectExtent\b[^>]*>/.exec(entryOf(zip, name))![0];
		return ['l', 't', 'r', 'b'].map((side) => Number(new RegExp(`\\b${side}="(\\d+)"`).exec(tag)![1]));
	};
	/** Word's revision marks and paragraph ids, which say who typed the text and not how it looks. */
	const withoutRevisions = (xml: string): string => xml.replace(/ w(?:14)?:(?:rsid\w*|paraId|textId)="[^"]*"/g, '');
	/** Each shape's box, look and stacking, canvas members included, for comparing two files. */
	const looks = (zip: ZipArchive): unknown[] => {
		const all = (shapes: ShapeInfo[]): ShapeInfo[] => shapes.flatMap((s) => [s, ...all(s.shapes ?? [])]);
		return all(looksOn(zip)).map(({ name, zOrder, hidden, rotation, fill, line, font, left, top, width, height }) =>
			({ name, zOrder, hidden, rotation, fill, line, font, left, top, width, height }));
	};

	it('reads what Word wrote for each property', () => {
		const zip = open('ShapesFormattedFixture.docm');
		expect(looked(zip, 'AnchoredBox')).toMatchObject({
			rotation: 30, zOrder: 1,
			fill: { type: 'solid', color: '#FF0000', transparency: 25 },
			line: { type: 'solid', color: '#008000', weight: 2.5, dash: 'dash' },
			font: { name: 'Arial', size: 14, bold: true, italic: true, underline: true, color: '#0000FF' },
		});
		expect(looked(zip, 'Board')).toMatchObject({ hidden: true, zOrder: 2 });
		// An inline shape is part of the text and stacks with nothing.
		expect(looked(zip, 'InlineOval')?.zOrder).toBeUndefined();
	});

	it('reports the font Word shows where the text sets none itself', () => {
		// Word reported Aptos at 12 points from the document defaults, and
		// drew the text white: its color is automatic, which in a shape is
		// the style's lt1. A Normal style that sets a color would win.
		expect(looked(open(), 'AnchoredBox')).toMatchObject({
			fill: { type: 'solid', color: '#156082', themeColor: 'accent1', automatic: true },
			line: { type: 'solid', themeColor: 'accent1', weight: 1.5, automatic: true },
			font: { name: 'Aptos', size: 12, color: '#FFFFFF' },
		});
	});

	it('writes each property the way Word does, in both copies of the shape', () => {
		const zip = open();
		for (const edit of WORDS_EDITS) {
			editStoryShape(zip, body(zip), edit);
		}
		const words = open('ShapesFormattedFixture.docm');
		expect(looks(reopen(zip))).toEqual(looks(words));

		const ours = entryOf(zip, 'AnchoredBox');
		const theirs = entryOf(words, 'AnchoredBox');
		expect(elementIn(ours, 'wps:spPr')).toBe(elementIn(theirs, 'wps:spPr'));
		const boxes = (entry: string): string[] => [...entry.matchAll(/<w:txbxContent>[\s\S]*?<\/w:txbxContent>/g)]
			.map((m) => withoutRevisions(m[0]));
		expect(boxes(ours)).toEqual(boxes(theirs));
		expect(boxes(ours)).toHaveLength(2);

		// The VML twin carries the same look. Word names a pure color where
		// VML has a name for it ("red"); the hex is the same color.
		const twin = vmlTag(zip, 'AnchoredBox');
		expect(twin).toContain('fillcolor="#FF0000"');
		expect(twin).toContain('strokecolor="#008000"');
		expect(twin).toContain('strokeweight="2.5pt"');
		expect(twin).toMatch(/style="[^"]*rotation:30[;"]/);
		expect(ours).toContain('<v:fill opacity="49087f"/><v:stroke dashstyle="dash"/><v:textbox>');
		// A hidden canvas hides its background shape with it.
		const style = (tag: string): string => /style="([^"]*)"/.exec(tag)![1];
		expect(style(vmlTag(zip, 'Board'))).toBe(style(vmlTag(words, 'Board')));
		expect(style(vmlTag(zip, '_x0000_s1027'))).toBe(style(vmlTag(words, '_x0000_s1027')));
	});

	it('turns, restacks and restyles shapes in a canvas as Word does', () => {
		const zip = open();
		for (const edit of WORDS_ARRANGEMENT) {
			editStoryShape(zip, body(zip), edit);
		}
		const words = open('ShapesArrangedFixture.docm');
		expect(looks(reopen(zip))).toEqual(looks(words));
		// To the front is one step of 1024 above the highest; nothing else moves.
		expect(heightOf(zip, 'AnchoredBox')).toBe(heightOf(words, 'AnchoredBox'));
		expect(heightOf(zip, 'Board')).toBe(heightOf(words, 'Board'));
		expect(vmlTag(zip, 'AnchoredBox')).toContain(`z-index:${heightOf(words, 'AnchoredBox')};`);
		// Turned a quarter, the shape's ink is measured from the turned box:
		// only the outline shows past it, as Word wrote within a point.
		const effect = effectOf(zip, 'AnchoredBox');
		effectOf(words, 'AnchoredBox').forEach((value, side) => expect(Math.abs(effect[side] - value), `side ${side}`).toBeLessThanOrEqual(12700));
		// Word left the VML of the shapes in the canvas as it was.
		for (const id of ['GroupedRect', 'GroupedOval']) {
			expect(vmlTag(zip, id), id).toBe(vmlTag(open(), id));
		}
	});

	it('makes room for the corners of a shape turned less than a quarter', () => {
		const zip = open();
		editStoryShape(zip, body(zip), WORDS_EDITS[0]);
		// Word's top and bottom: the turned corners and half the outline,
		// within a point and a half. Its left and right come from its
		// renderer and are not symmetric; these are the geometry.
		const [left, top, right, bottom] = effectOf(zip, 'AnchoredBox');
		const [, wordsTop, , wordsBottom] = effectOf(open('ShapesFormattedFixture.docm'), 'AnchoredBox');
		expect(Math.abs(top - wordsTop)).toBeLessThanOrEqual(19050);
		expect(Math.abs(bottom - wordsBottom)).toBeLessThanOrEqual(19050);
		expect(left).toBe(right);
		expect(left).toBeGreaterThan(0);
	});

	it('trades places with the shape above or below it', () => {
		const zip = open();
		editStoryShape(zip, body(zip), { action: 'update', name: 'Board', zOrder: 'backward' });
		const again = reopen(zip);
		expect(looked(again, 'Board')?.zOrder).toBe(1);
		expect(looked(again, 'AnchoredBox')?.zOrder).toBe(2);
		expect([heightOf(again, 'Board'), heightOf(again, 'AnchoredBox')]).toEqual([251659264, 251660288]);
	});

	it('shows a hidden shape again, and a font color goes back to the style\'s', () => {
		const zip = open('ShapesFormattedFixture.docm');
		editStoryShape(zip, body(zip), { action: 'update', name: 'Board', hidden: false });
		editStoryShape(zip, body(zip), { action: 'update', name: 'AnchoredBox', font: { color: '' }, rotation: 0 });
		const again = reopen(zip);
		expect(looked(again, 'Board')?.hidden).toBeUndefined();
		expect(vmlTag(again, 'Board')).toContain('visibility:visible');
		expect(vmlTag(again, '_x0000_s1027')).toContain('visibility:visible');
		expect(looked(again, 'AnchoredBox')?.rotation).toBeUndefined();
		expect(looked(again, 'AnchoredBox')?.font?.color).toBe('#FFFFFF');
		expect(vmlTag(again, 'AnchoredBox')).not.toContain('rotation:');
	});

	it('gives text to a shape that has none, in both copies', () => {
		const zip = open();
		editStoryShape(zip, body(zip), { action: 'update', name: 'InlineOval', text: 'Hi' });
		const entry = entryOf(zip, 'InlineOval');
		expect(named(reopen(zip), 'InlineOval')?.text).toBe('Hi');
		expect(looked(reopen(zip), 'InlineOval')?.font).toEqual({ name: 'Aptos', size: 12, color: '#FFFFFF' });
		expect(entry.indexOf('<wps:txbx>')).toBeLessThan(entry.indexOf('<wps:bodyPr'));
		expect(entry).toMatch(/<v:oval\b[^>]*>[\s\S]*<v:textbox><w:txbxContent><w:p><w:pPr><w:jc w:val="center"\/><\/w:pPr><w:r><w:t xml:space="preserve">Hi<\/w:t>/);
		expect([...entry.matchAll(/>Hi</g)]).toHaveLength(2);
	});

	it('adds a shape with its look in one edit, on top of the others', () => {
		const zip = open();
		const name = editStoryShape(zip, body(zip), {
			action: 'add', type: 'oval', left: 10, top: 10, width: 50, height: 40, text: 'Hi', rotation: 45,
			fill: { type: 'solid', color: '#00FF00' }, line: { type: 'none' }, font: { bold: true, color: '#000000' },
		});
		expect(looked(reopen(zip), name)).toMatchObject({
			left: 10, top: 10, width: 50, height: 40, rotation: 45, text: 'Hi', zOrder: 3,
			fill: { type: 'solid', color: '#00FF00' }, line: { type: 'none' },
			font: { bold: true, color: '#000000', size: 12, name: 'Aptos' },
		});
	});

	it('refuses what a kind of shape does not have, and a value that is not one', () => {
		const zip = open();
		const refusals: Array<[ShapeEdit, RegExp]> = [
			[{ action: 'update', name: 'Board', fill: { type: 'none' } }, /'Board' is a canvas, which has no fill/],
			[{ action: 'update', name: 'Board', rotation: 10 }, /'Board' is a canvas, which has no rotation/],
			[{ action: 'update', name: 'InlineOval', zOrder: 'front' }, /inline with the text, which does not stack/],
			[{ action: 'update', name: 'GroupedRect', zOrder: 'front' }, /'GroupedRect' is inside 'Board' and stacks with it/],
			[{ action: 'update', name: 'InlineOval', font: { bold: true } }, /has no text yet; give it text, and the font with it/],
			[{ action: 'update', name: 'AnchoredBox', fill: { type: 'solid', color: 'red' } }, /'red' is not a fill color/],
			[{ action: 'update', name: 'AnchoredBox', font: { size: 500 } }, /A font size is in points, from 1 to 409/],
		];
		for (const [edit, message] of refusals) {
			expect(() => editStoryShape(zip, body(zip), edit), JSON.stringify(edit)).toThrow(message);
		}
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
