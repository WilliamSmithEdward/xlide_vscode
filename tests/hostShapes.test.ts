import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { editShape, listShapes, shapeMacros } from '../src/vba/projectService';
import { PRESET_GEOMETRY, PRESET_LABELS } from '../src/vba/shapes';

/** The shape types the agent tool's schema offers. */
const toolTypes = (): string[] => {
	const manifest = JSON.parse(fs.readFileSync(path.join(__dirname, '..', 'package.json'), 'utf8'));
	const tool = manifest.contributes.languageModelTools
		.find((one: { name: string }) => one.name === 'xlide_editShape');
	return tool.inputSchema.properties.type.enum;
};

// The routing seam: one pair of entry points for three hosts. The per-host
// markup is covered by xlsxShapes, pptShapes and docShapes; what is checked
// here is that a file reaches the right one, that a macro is validated
// against the project before it is written, and that each host refuses what
// it cannot do in its own words.
const FIXTURES = path.join(__dirname, 'fixtures', 'binaries');

let dir: string;
const copy = (fixture: string, as = fixture): string => {
	const target = path.join(dir, as);
	fs.copyFileSync(path.join(FIXTURES, fixture), target);
	return target;
};

beforeEach(() => {
	dir = fs.mkdtempSync(path.join(os.tmpdir(), 'xlide-host-shapes-'));
});

afterEach(() => {
	fs.rmSync(dir, { recursive: true, force: true });
});

describe('listing shapes, whichever host owns the file', () => {
	it('lists a workbook by worksheet', () => {
		const { surfaces } = listShapes(copy('ShapesFixture.xlsm'));

		expect(surfaces.map((s) => s.surface)).toEqual(['Sheet1', 'Sheet2']);
		expect(surfaces[0].shapes[0]).toMatchObject({ name: 'RunButton', range: 'B3:D5', macro: 'DoIt' });
	});

	it('lists a presentation by slide', () => {
		const { surfaces } = listShapes(copy('PowerPointShapesFixture.pptm'));

		expect(surfaces.map((s) => s.surface)).toEqual(['Slide 1', 'Slide 2']);
		expect(surfaces[0].shapes[0]).toMatchObject({ name: 'ClickMe', left: 60, macro: 'Macros.SayHello' });
	});

	it('lists a document by story, the body first', () => {
		const { surfaces } = listShapes(copy('WordShapesFixture.docm'));

		expect(surfaces[0].surface).toBe('Document');
		expect(surfaces.map((s) => s.surface)).toContain('Header');
		expect(surfaces[0].shapes.map((s) => s.name)).toEqual(['InlineOval', 'AnchoredBox', 'Board']);
	});

	it('takes one surface by name, in each host', () => {
		expect(listShapes(copy('PowerPointShapesFixture.pptm'), 'Slide 2').surfaces.map((s) => s.surface))
			.toEqual(['Slide 2']);
		expect(listShapes(copy('WordShapesFixture.docm'), 'Header').surfaces[0].shapes.map((s) => s.name))
			.toEqual(['HeaderBadge']);
	});

	it('refuses a container with no drawing surface, and says what it is', () => {
		expect(() => listShapes(copy('AccessFixture.accdb')))
			.toThrow(/is an Access database; listing shapes need an OOXML/);
	});
});

describe('the Subs a shape can run, whichever host owns the file', () => {
	it('offers a slide shape the Subs in standard modules', () => {
		expect(shapeMacros(copy('PowerPointShapesFixture.pptm')).macros).toEqual([
			{ macro: 'SayHello', module: 'Macros', proc: 'SayHello' },
			{ macro: 'Unlinked', module: 'Macros', proc: 'Unlinked' },
		]);
	});

	it('offers a Word shape none, since Word runs no macro from one', () => {
		expect(shapeMacros(copy('WordShapesFixture.docm')).macros).toEqual([]);
	});
});

describe('the AutoShapes that can be added', () => {
	it('offers the same set through the type union, both tables and the tool', () => {
		// Three tables and a JSON schema have to agree, or a type the tool
		// accepts reaches a writer with no geometry to write.
		const declared = Object.keys(PRESET_GEOMETRY).sort();

		expect(Object.keys(PRESET_LABELS).sort()).toEqual(declared);
		expect(toolTypes().filter((t) => t !== 'textBox' && t !== 'button').sort()).toEqual(declared);
	});

	it('writes the geometry Office writes, including the four that are not named after their constant', () => {
		// Measured from a file holding one of each, saved by Excel and by
		// PowerPoint: msoShapeCross is `plus`, not `cross`, and the other
		// three are likewise named for a shape they are not.
		expect(PRESET_GEOMETRY.cross).toBe('plus');
		expect(PRESET_GEOMETRY.pentagon).toBe('pentagon');
		expect(PRESET_GEOMETRY.foldedCorner).toBe('foldedCorner');
		expect(PRESET_GEOMETRY.smileyFace).toBe('smileyFace');
		expect(PRESET_LABELS.triangle).toBe('Isosceles Triangle');
	});

	it('adds every one of them to a slide, named as the host would name it', () => {
		for (const type of Object.keys(PRESET_GEOMETRY) as Array<keyof typeof PRESET_GEOMETRY>) {
			const deck = copy('PowerPointShapesFixture.pptm', `${type}.pptm`);
			const { name } = editShape(deck, 'Slide 1', { action: 'add', type, left: 10, top: 10 });
			const added = listShapes(deck, 'Slide 1').surfaces[0].shapes.find((s) => s.name === name);

			expect(name, type).toBe(`${PRESET_LABELS[type]} 4`);
			expect(added?.geometry, type).toBe(PRESET_GEOMETRY[type]);
		}
	});
});

describe('editing shapes, whichever host owns the file', () => {
	it('links a slide shape to a Sub the presentation really has', () => {
		const deck = copy('PowerPointShapesFixture.pptm');

		// Spelled as the project spells it, not as the caller typed it.
		const { name } = editShape(deck, 'Slide 1', { action: 'update', name: 'Badge', macro: 'macros.sayhello' });

		expect(name).toBe('Badge');
		expect(listShapes(deck, 'Slide 1').surfaces[0].shapes.find((s) => s.name === 'Badge')?.macro)
			.toBe('Macros.SayHello');
	});

	it('refuses a slide macro the presentation does not have, and leaves the file alone', () => {
		const deck = copy('PowerPointShapesFixture.pptm');
		const before = fs.readFileSync(deck);

		expect(() => editShape(deck, 'Slide 1', { action: 'update', name: 'Badge', macro: 'NoSuchSub' }))
			.toThrow(/no Public Sub named NoSuchSub/);
		expect(fs.readFileSync(deck).equals(before)).toBe(true);
	});

	it("says PowerPoint's name, not Excel's, when it turns a macro down", () => {
		const deck = copy('PowerPointShapesFixture.pptm');

		expect(() => editShape(deck, 'Slide 1', { action: 'update', name: 'Badge', macro: 'Other!Run' }))
			.toThrow(/links shapes only to Subs in this presentation/);
	});

	it('adds a shape to a slide and saves it', () => {
		const deck = copy('PowerPointShapesFixture.pptm');

		const { name } = editShape(deck, 'Slide 2', { action: 'add', type: 'oval', left: 10, top: 20 });

		expect(listShapes(deck, 'Slide 2').surfaces[0].shapes.map((s) => s.name)).toContain(name);
	});

	it("adds a shape to a document's body when no surface is named", () => {
		const doc = copy('WordShapesFixture.docm');

		const { name } = editShape(doc, '', { action: 'add', type: 'rectangle', name: 'Stamp', left: 10, top: 20 });

		expect(name).toBe('Stamp');
		expect(listShapes(doc, 'Document').surfaces[0].shapes.map((s) => s.name)).toContain('Stamp');
	});

	it('refuses a Word macro with the reason, and leaves the file alone', () => {
		const doc = copy('WordShapesFixture.docm');
		const before = fs.readFileSync(doc);

		expect(() => editShape(doc, '', { action: 'update', name: 'AnchoredBox', macro: 'SayHello' }))
			.toThrow(/Word cannot run a macro from a shape/);
		expect(fs.readFileSync(doc).equals(before)).toBe(true);
	});

	it('asks which surface when the host has more than one obvious one', () => {
		expect(() => editShape(copy('PowerPointShapesFixture.pptm'), '', { action: 'add', type: 'oval', left: 0, top: 0 }))
			.toThrow(/needs the slide it is on/);
		expect(() => editShape(copy('ShapesFixture.xlsm'), '', { action: 'add', type: 'oval', range: 'B2:C3' }))
			.toThrow(/needs the worksheet it is on/);
	});
});
