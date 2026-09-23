import { describe, expect, it } from 'vitest';
import * as fs from 'fs';
import * as path from 'path';
import { shapeEditFromForm, shapeEditorModel, shapeFormFields, type ShapeFormValues } from '../src/shapeEditorModel';
import { listShapes } from '../src/vba/projectService';
import type { ShapeInfo } from '../src/vba/shapes';

// The editor's form, held to shapes as the engine lists them from files the
// applications saved: which fields each shape gets, and that a Save sends
// what changed and nothing else.
const FIXTURES = path.join(__dirname, 'fixtures', 'binaries');
const shapeIn = (fixture: string, surface: string, name: string): ShapeInfo => {
	const file = path.join(FIXTURES, fixture);
	expect(fs.existsSync(file), fixture).toBe(true);
	const all = (shapes: ShapeInfo[]): ShapeInfo[] => shapes.flatMap((s) => [s, ...all(s.shapes ?? [])]);
	return all(listShapes(file, surface).surfaces[0].shapes).find((s) => s.name === name)!;
};

const MACROS = [
	{ macro: 'DoIt', module: 'Macros', proc: 'DoIt' },
	{ macro: 'Sheet1.OnSheet', module: 'Sheet1', proc: 'OnSheet' },
];

/** The form as it opens, with `changes` typed in. */
const submitted = (values: ShapeFormValues, changes: Partial<ShapeFormValues> = {}): ShapeFormValues => ({ ...values, ...changes });

describe('the fields a shape gets', () => {
	it('gives an Excel AutoShape its cells, text, macro and whole look', () => {
		const fields = shapeFormFields('excel', 'shape', { inGroup: false });
		expect(fields).toMatchObject({
			position: 'range', text: true, macro: true, rotation: true, zOrder: true,
			fill: true, line: true, font: true, delete: true, linkedCell: false, inputRange: false,
		});
	});

	it('gives an Excel form control what its Format Control dialog has, and no rotation', () => {
		expect(shapeFormFields('excel', 'button', { inGroup: false })).toMatchObject({
			text: true, font: true, fill: false, line: false, rotation: false, linkedCell: false,
		});
		expect(shapeFormFields('excel', 'checkBox', { inGroup: false })).toMatchObject({
			text: true, font: false, fill: true, line: true, linkedCell: true, inputRange: false,
		});
		expect(shapeFormFields('excel', 'dropDown', { inGroup: false })).toMatchObject({
			text: false, linkedCell: true, inputRange: true,
		});
	});

	it('moves a shape in an Excel group with the group, and stacks it with the group', () => {
		expect(shapeFormFields('excel', 'shape', { inGroup: true })).toMatchObject({
			position: 'none', positionNote: 'A shape in a group moves with the group.',
			zOrder: false, zOrderNote: 'A shape in a group stacks with the group.', delete: false,
		});
	});

	it('explains that Word runs no macro from a shape, and sizes an inline one only', () => {
		const fields = shapeFormFields('word', 'shape', { inGroup: false, inline: true });
		expect(fields).toMatchObject({ position: 'size', macro: false, zOrder: false });
		expect(fields.macroNote).toMatch(/Word cannot run a macro from a shape/);
		expect(fields.zOrderNote).toBe('A shape inline with the text does not stack.');
	});

	it('runs no macro from a group, and says where the macro goes instead', () => {
		const fields = shapeFormFields('powerpoint', 'group', { inGroup: false });
		expect(fields).toMatchObject({ macro: false, fill: false, rotation: true, position: 'box' });
		expect(fields.macroNote).toBe('A group runs no macro; give the macro to a shape in it.');
	});
});

describe('what a Save sends', () => {
	const runButton = shapeIn('ShapesFixture.xlsm', 'Sheet1', 'RunButton');
	const model = shapeEditorModel({ host: 'excel', fileName: 'Shapes.xlsm', surface: 'Sheet1', shape: runButton }, MACROS);

	it('opens on the shape as the file has it, its theme look marked automatic', () => {
		expect(model.values).toMatchObject({
			name: 'RunButton', range: 'B3:D5', macro: 'DoIt', rotation: '0',
			fillType: 'auto', fillColor: '#156082', lineType: 'auto', lineWeight: '1.5',
			fontName: 'Aptos Narrow', fontSize: '11', fontColor: '#FFFFFF',
		});
	});

	it('sends nothing when nothing changed', () => {
		expect(shapeEditFromForm(model, submitted(model.values))).toEqual({ surface: 'Sheet1', errors: {} });
	});

	it('sends only what changed, a color field touched making the fill solid', () => {
		const result = shapeEditFromForm(model, submitted(model.values, {
			name: 'Go', macro: 'Sheet1.OnSheet', fillType: 'solid', fillColor: '#ff0000', fontBold: true, zOrder: 'front',
		}));
		expect(result.errors).toEqual({});
		expect(result.edit).toEqual({
			action: 'update', name: 'RunButton', newName: 'Go', macro: 'Sheet1.OnSheet', zOrder: 'front',
			fill: { type: 'solid', color: '#FF0000' }, font: { bold: true },
		});
	});

	it('unlinks the macro on (none), and takes the text back to its style\'s color', () => {
		const result = shapeEditFromForm(model, submitted(model.values, { macro: '', fontColorFromStyle: true }));
		expect(result.edit).toEqual({ action: 'update', name: 'RunButton', macro: '', font: { color: '' } });
	});

	it('turns an outline on with the color, weight and dash the form shows', () => {
		const oval = shapeIn('ShapesFormattedFixture.xlsm', 'Sheet1', 'Oval 3');
		const noLine = shapeEditorModel({ host: 'excel', fileName: 'Shapes.xlsm', surface: 'Sheet1', shape: oval }, MACROS);
		expect(noLine.values.lineType).toBe('none');
		const result = shapeEditFromForm(noLine, submitted(noLine.values, { lineType: 'solid', lineColor: '#00ff00', lineWeight: '3', lineDash: 'dash' }));
		expect(result.edit).toEqual({ action: 'update', name: 'Oval 3', line: { type: 'solid', color: '#00FF00', weight: 3, dash: 'dash' } });
	});

	it('says what is wrong with each field, and sends nothing', () => {
		const result = shapeEditFromForm(model, submitted(model.values, {
			name: ' ', range: 'nowhere', rotation: 'ten', fillType: 'solid', fillColor: 'red', fontSize: '500',
		}));
		expect(result.edit).toBeUndefined();
		expect(result.errors).toEqual({
			name: 'A shape needs a name.',
			range: "'nowhere' is not a cell range; give one such as B2:D4.",
			rotation: "The rotation is a number; 'ten' is not one.",
			fillColor: 'The fill color is a color as #RRGGBB, such as #FF0000.',
			fontSize: 'A font size is from 1 to 409 points.',
		});
	});

	it('does not rewrite a position shown rounded when it was left alone', () => {
		const caption = shapeIn('PowerPointShapesFixture.pptm', 'Slide 1', 'Caption');
		const slide = shapeEditorModel({ host: 'powerpoint', fileName: 'Deck.pptm', surface: 'Slide 1', shape: caption }, []);
		expect(slide.values.height).toBe('29.08');
		expect(shapeEditFromForm(slide, submitted(slide.values)).edit).toBeUndefined();
		expect(shapeEditFromForm(slide, submitted(slide.values, { height: '40' })).edit).toEqual({ action: 'update', name: 'Caption', height: 40 });
	});
});

describe('adding with the form', () => {
	it('adds with the cells and the look given, leaving the rest to the application', () => {
		const model = shapeEditorModel({ host: 'excel', fileName: 'Shapes.xlsm', surface: 'Sheet1' }, MACROS);
		expect(model.types?.map((t) => t.value)).toContain('button');
		const result = shapeEditFromForm(model, submitted(model.values, {
			type: 'oval', range: 'k2:l4', text: 'Hi', macro: 'DoIt', fillType: 'solid', fillColor: '#00FF00',
		}));
		expect(result.edit).toEqual({
			action: 'add', type: 'oval', range: 'K2:L4', text: 'Hi', macro: 'DoIt', fill: { type: 'solid', color: '#00FF00' },
		});
	});

	it('needs the cells in Excel, and left and top on a slide', () => {
		const excel = shapeEditorModel({ host: 'excel', fileName: 'Shapes.xlsm', surface: 'Sheet1' }, []);
		expect(shapeEditFromForm(excel, excel.values).errors.range).toBe('Give the cells the shape covers, such as B2:D4.');
		const slide = shapeEditorModel({ host: 'powerpoint', fileName: 'Deck.pptm', surface: 'Slide 1' }, []);
		expect(slide.types?.map((t) => t.value)).not.toContain('button');
		expect(shapeEditFromForm(slide, slide.values).errors.left).toMatch(/Give left and top/);
	});

	it('floats a Word shape given left and top, puts it inline given neither, and refuses one of the two', () => {
		const doc = shapeEditorModel({ host: 'word', fileName: 'Report.docm', surface: 'Document', surfaces: ['Document', 'Header'] }, []);
		expect(doc.fields.macro).toBe(false);
		expect(shapeEditFromForm(doc, submitted(doc.values, { surface: 'Header' }))).toEqual({
			edit: { action: 'add', type: 'rectangle' }, surface: 'Header', errors: {},
		});
		expect(shapeEditFromForm(doc, submitted(doc.values, { left: '10', top: '20' })).edit)
			.toEqual({ action: 'add', type: 'rectangle', left: 10, top: 20 });
		expect(shapeEditFromForm(doc, submitted(doc.values, { left: '10' })).errors.top)
			.toBe('Give both left and top to make the shape float, or neither.');
	});
});
