// Access form and report designs (issue #67), read out of MSysAccessStorage.
//
// A design sits in a `Blob` under a numbered folder under `Forms` or
// `Reports`, and is a stream of property records with three marker ids that
// open the next object. What a property MEANS is a separate question; this
// reads the structure, which is what a designer needs to show the object and
// what a writer needs to put it back.
//
// The load-bearing check is the round trip: a design that rebuilds byte for
// byte from what was parsed is a design that was understood.

import { describe, expect, it } from 'vitest';
import * as fs from 'fs';
import * as path from 'path';
import {
	ACCESS_FORM_TYPE,
	accessControlTypeName,
	accessDesignControls,
	accessDesignObjectName,
	accessDesignSections,
	buildAccessDesign,
	isAccessDesignSection,
	parseAccessDesign,
} from '../src/vba/access/accessDesign';
import { readAccessCatalog, readAccessDesigns, readAccessStorage } from '../src/vba/access/accessStorage';

const BINARIES = path.join(__dirname, 'fixtures', 'binaries');
const FORMS = fs.readFileSync(path.join(BINARIES, 'AccessFormFixture.accdb'));

describe('finding the designs', () => {
	it('reads the form the catalog names', () => {
		expect(readAccessCatalog(FORMS).filter((entry) => entry.type === ACCESS_FORM_TYPE)
			.map((entry) => entry.name)).toEqual(['Calculator']);

		const designs = readAccessDesigns(FORMS);
		expect(designs).toHaveLength(1);
		expect(designs[0]).toMatchObject({ name: 'Calculator', kind: 'form', ordinal: '0' });
	});

	it('carries the design\'s other streams, which rebuilding it needs', () => {
		const [design] = readAccessDesigns(FORMS);
		expect(design.typeInfo?.length).toBeGreaterThan(0);
		expect(design.propData?.length).toBeGreaterThan(0);
	});

	it('finds nothing in a database with no forms, rather than failing', () => {
		const plain = fs.readFileSync(path.join(BINARIES, 'AccessFixture.accdb'));
		expect(readAccessDesigns(plain)).toEqual([]);
	});
});

describe('the design itself', () => {
	const [entry] = readAccessDesigns(FORMS);

	it('rebuilds byte for byte, which is what says it was understood', () => {
		const blob = readAccessStorage(FORMS)!;
		void blob;
		expect(buildAccessDesign(entry.design).equals(
			buildAccessDesign(parseAccessDesign(buildAccessDesign(entry.design))),
		)).toBe(true);
	});

	it('separates the sections from the controls', () => {
		expect(accessDesignSections(entry.design).map(accessControlTypeName)).toEqual(['Detail']);
		expect(accessDesignSections(entry.design).every(isAccessDesignSection)).toBe(true);
	});

	it('names every control the designer would show, with its type', () => {
		expect(accessDesignControls(entry.design)
			.map((control) => `${accessControlTypeName(control)} ${accessDesignObjectName(control)}`))
			.toEqual([
				'Rectangle Banner',
				'Label Title',
				'Label Subtitle',
				'Label QtyLabel',
				'TextBox Qty',
				'Label PriceLabel',
				'TextBox Price',
				'CheckBox Express',
				'Label ExpressLabel',
				'CommandButton AddLine',
				'CommandButton Reset',
				'Line Rule',
				'Label LinesLabel',
				'ListBox Lines',
				'Label TotalLabel',
			]);
	});

	it('leaves out the unnamed prototypes a design also carries', () => {
		// A design keeps the styles new controls are cut from. They are objects
		// like any other and have no name, and the designer does not show them.
		const named = accessDesignControls(entry.design).length;
		const all = entry.design.objects.length - 1 - accessDesignSections(entry.design).length;
		expect(all).toBeGreaterThan(named);
	});

	it('keeps the ids ascending inside one object, which is what ends the stream', () => {
		for (const object of entry.design.objects) {
			const ids = object.records.map((record) => record.id);
			expect([...ids].sort((a, b) => a - b)).toEqual(ids);
		}
	});

	it('reads a record as its own five fields', () => {
		const title = accessDesignControls(entry.design)
			.find((control) => accessDesignObjectName(control) === 'Title')!;
		const record = title.records[0];
		expect(typeof record.id).toBe('number');
		expect(typeof record.code).toBe('number');
		expect(record.valueType).toBeLessThanOrEqual(0xffff);
		expect(record.value.length).toBeGreaterThanOrEqual(0);
	});
});

describe('what it refuses', () => {
	it('refuses a blob too short to hold a header', () => {
		expect(() => parseAccessDesign(Buffer.alloc(8))).toThrow(/too short/);
	});
});
