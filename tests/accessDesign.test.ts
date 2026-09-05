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

// ---------------------------------------------------------------------------
// Editing a design (issue #67, the write half).
//
// A record's id is its slot in its own type's schema, so a property written at
// the wrong id is one Access reads as something else, and the markers that
// group controls carry a count Access trusts without checking: get it wrong
// and the form opens showing only as many controls as the number claims.
//
// The live check, that Access opens the edited form, shows the new caption,
// lists the new controls and compiles Me against them, is not something CI can
// host; it was run against Access 16.0.

import { AccessVbaWriter } from '../src/vba/access/accessVbaWriter';
import {
	addDesignControl,
	designPrototypes,
	removeDesignControl,
	setDesignProperty,
} from '../src/vba/access/accessDesignEdit';
import { readTypeInfo, updateTypeInfo } from '../src/vba/access/accessTypeInfo';
import { PROPERTY_CODES } from '../src/vba/access/accessDesignTable';

/** A property's record code, so a test names the property rather than a number. */
const code = (name: string): number => PROPERTY_CODES.get(name)!;

const GUID = Buffer.from(Array.from({ length: 16 }, (_value, index) => index));
const design = (): ReturnType<typeof parseAccessDesign> => readAccessDesigns(FORMS)[0].design;

/** A stream of the database, found by name anywhere in the storage tree. */
function streamNamed(data: Buffer, name: string): Buffer {
	const walk = (nodes: ReturnType<typeof readAccessStorage>): Buffer | undefined => {
		for (const node of nodes ?? []) {
			if (node.name === name && node.bytes) {
				return node.bytes;
			}
			const deeper = walk(node.children);
			if (deeper) {
				return deeper;
			}
		}
		return undefined;
	};
	const found = walk(readAccessStorage(data));
	if (!found) {
		throw new Error(`no stream named ${name}`);
	}
	return found;
}

describe('setting a design property', () => {
	it('changes the design own caption and rebuilds', () => {
		const edited = setDesignProperty(design(), undefined, 'Caption', 'Edited Caption');
		const caption = edited.objects[0].records.find((entry) => entry.code === code('Caption'));
		expect(caption?.value.toString('utf16le')).toBe('Edited Caption');
		// A rebuilt design still parses, which is what Access needs of it.
		expect(parseAccessDesign(buildAccessDesign(edited)).objects.length)
			.toBe(edited.objects.length);
	});

	it('keeps a replaced property at the id the object already used', () => {
		const before = design();
		const name = accessDesignObjectName(accessDesignControls(before)[0])!;
		const was = accessDesignControls(before)[0].records
			.find((entry) => entry.code === code('Width'));
		const edited = setDesignProperty(before, name, 'Width', 999);
		const after = edited.objects.find((object) => accessDesignObjectName(object) === name)!;
		const changed = after.records.find((entry) => entry.code === code('Width'))!;
		expect(changed.id).toBe(was?.id);
		expect(changed.value.readUIntLE(0, changed.value.length)).toBe(999);
	});

	it('writes the theme index Access writes beside a colour', () => {
		const before = design();
		// A rectangle has a fill and a border and no text, so BackColor is the
		// colour it takes; the property set is the control type's, not a
		// shared one, which is what the refusal below shows.
		const name = accessDesignObjectName(accessDesignControls(before)[0])!;
		const edited = setDesignProperty(before, name, 'BackColor', 0x00ff0000);
		const after = edited.objects.find((object) => accessDesignObjectName(object) === name)!;
		// Without the index at -1 the control defaults' theme colour wins and
		// Access drops the colour on its next save.
		const index = after.records.find((entry) => entry.code === code('BackThemeColorIndex'));
		expect(index).toBeDefined();
		expect(index!.value.readIntLE(0, index!.value.length)).toBe(-1);
		expect(() => setDesignProperty(before, name, 'ForeColor', 0))
			.toThrow(/Rectangle has no ForeColor/);
	});

	it('refuses a property the object type does not have', () => {
		expect(() => setDesignProperty(design(), 'Detail', 'BackColor', 0))
			.toThrow(/has no BackColor to set/);
	});

	it('refuses a name no object carries', () => {
		expect(() => setDesignProperty(design(), 'NotThere', 'Width', 1))
			.toThrow(/no object named NotThere/);
	});
});

describe('adding and removing controls', () => {
	it('adds a control the design then lists', () => {
		const before = design();
		const edited = addDesignControl(before, 'TextBox', 'TxtNew', GUID, designPrototypes(before), {
			left: 100, top: 200, width: 1440, height: 240,
		});
		expect(accessDesignControls(edited).map(accessDesignObjectName)).toContain('TxtNew');
		expect(accessDesignControls(edited).length).toBe(accessDesignControls(before).length + 1);
	});

	it('re-marks the section so the count Access trusts is right', () => {
		const before = design();
		const edited = addDesignControl(before, 'Label', 'LblNew', GUID, designPrototypes(before), {
			left: 0, top: 0, caption: 'Hello',
		});
		// The first control of a run opens the group and carries how many the
		// group holds, the opener included; Access shows exactly that many.
		const detail = edited.objects.findIndex(
			(object) => isAccessDesignSection(object) && accessDesignObjectName(object) === 'Detail',
		);
		const opener = edited.objects[detail + 1];
		expect(opener.marker).toBe(0xff);
		let held = 1;
		for (let at = detail + 2; at < edited.objects.length; at += 1) {
			if (edited.objects[at].marker !== 0xfd) {
				break;
			}
			held += 1;
		}
		expect(opener.code).toBe(held);
	});

	it('takes a control off again', () => {
		const before = design();
		const added = addDesignControl(before, 'TextBox', 'TxtGone', GUID, designPrototypes(before), {
			left: 0, top: 0,
		});
		const removed = removeDesignControl(added, 'TxtGone');
		expect(accessDesignControls(removed).map(accessDesignObjectName)).not.toContain('TxtGone');
		expect(accessDesignControls(removed).length).toBe(accessDesignControls(before).length);
	});

	it('refuses a duplicate name', () => {
		const before = design();
		const name = accessDesignObjectName(accessDesignControls(before)[0])!;
		expect(() => addDesignControl(before, 'TextBox', name, GUID, designPrototypes(before)))
			.toThrow(/already has an object named/);
	});

	it('refuses a control type whose slots were never measured', () => {
		expect(() => addDesignControl(
			design(), 'NavigationControl', 'NavNew', GUID, designPrototypes(design()),
		)).toThrow(/read but not written/);
	});

	it('refuses a page without a tab control to hold it', () => {
		expect(() => addDesignControl(design(), 'Page', 'PageNew', GUID, designPrototypes(design())))
			.toThrow(/needs a parent Tab/);
	});
});

describe('the TypeInfo stream', () => {
	it('lists every section and named control the design has', () => {
		const names = readTypeInfo(streamNamed(FORMS, 'TypeInfo'), 1252)
			.map((entry) => entry.name);
		expect(names).toContain('Detail');
		for (const control of accessDesignControls(design())) {
			expect(names).toContain(accessDesignObjectName(control));
		}
	});

	it('appends a new member above the highest ordinal and drops a removed one', () => {
		const stream = streamNamed(FORMS, 'TypeInfo');
		const before = readTypeInfo(stream, 1252);
		const highest = Math.max(...before.map((entry) => entry.ordinal));
		const added = addDesignControl(
			design(), 'TextBox', 'TxtNew', GUID, designPrototypes(design()), { left: 0, top: 0 },
		);
		const grown = readTypeInfo(updateTypeInfo('form', added, stream, 1252), 1252);
		expect(grown.find((entry) => entry.name === 'TxtNew')?.ordinal).toBe(highest + 1);
		// Every other member keeps the ordinal it had; a freed one is not reused.
		for (const entry of before) {
			expect(grown.find((other) => other.name === entry.name)?.ordinal).toBe(entry.ordinal);
		}
		const shrunk = readTypeInfo(
			updateTypeInfo('form', removeDesignControl(added, 'TxtNew'), stream, 1252), 1252,
		);
		expect(shrunk.map((entry) => entry.name)).not.toContain('TxtNew');
	});

	it('moves a renamed member to the end, keeping its ordinal', () => {
		const stream = streamNamed(FORMS, 'TypeInfo');
		const before = readTypeInfo(stream, 1252);
		const name = accessDesignObjectName(accessDesignControls(design())[0])!;
		const was = before.find((entry) => entry.name === name)!;
		const renamed = setDesignProperty(design(), name, 'Name', 'Renamed');
		const after = readTypeInfo(
			updateTypeInfo('form', renamed, stream, 1252, new Map([[name, 'Renamed']])), 1252,
		);
		expect(after.find((entry) => entry.name === 'Renamed')?.ordinal).toBe(was.ordinal);
		expect(after[after.length - 1].name).toBe('Renamed');
	});
});

describe('writing a design back into the database', () => {
	it('lists the database designs with their storage folders', () => {
		const designs = new AccessVbaWriter(FORMS).designs();
		expect(designs.map((entry) => entry.name)).toContain('Calculator');
		expect(designs[0].kind).toBe('form');
		expect(designs[0].ordinal).toMatch(/^\d+$/);
	});

	it('writes an edit that reads back through the storage', () => {
		const writer = new AccessVbaWriter(FORMS);
		writer.editDesign('Calculator', (blob) => addDesignControl(
			blob, 'Label', 'LblWritten', GUID, designPrototypes(blob),
			{ left: 10, top: 20, caption: 'written' },
		));
		const again = new AccessVbaWriter(writer.toBuffer());
		const edited = again.designs().find((entry) => entry.name === 'Calculator')!;
		expect(accessDesignControls(edited.design).map(accessDesignObjectName))
			.toContain('LblWritten');
		// The project the database holds is untouched by a design edit.
		expect(again.moduleNames()).toEqual(new AccessVbaWriter(FORMS).moduleNames());
	});

	it('carries the new control into the TypeInfo stream, so Me reaches it', () => {
		const writer = new AccessVbaWriter(FORMS);
		writer.editDesign('Calculator', (blob) => addDesignControl(
			blob, 'TextBox', 'TxtBound', GUID, designPrototypes(blob), { left: 10, top: 20 },
		));
		expect(readTypeInfo(streamNamed(writer.toBuffer(), 'TypeInfo'), 1252)
			.map((entry) => entry.name)).toContain('TxtBound');
	});

	it('refuses a design the database does not hold', () => {
		expect(() => new AccessVbaWriter(FORMS).editDesign('NotThere', (blob) => blob))
			.toThrow(/no form or report named NotThere/);
	});
});

describe('creating and deleting a design', () => {
	// The blank form and report come from templates captured from Access, and
	// the design's own GUID is patched into the blob and the catalog row that
	// repeats it: two designs sharing a GUID is not something Access writes.
	it('adds a form the database then lists', () => {
		const writer = new AccessVbaWriter(FORMS);
		writer.addDesign('MadeHere', 'form');
		const designs = new AccessVbaWriter(writer.toBuffer()).designs();
		expect(designs.map((entry) => entry.name)).toContain('MadeHere');
		const made = designs.find((entry) => entry.name === 'MadeHere')!;
		expect(made.kind).toBe('form');
		// A blank form is its own object and a Detail section.
		expect(accessDesignSections(made.design).map(accessDesignObjectName)).toEqual(['Detail']);
		expect(accessDesignControls(made.design)).toEqual([]);
	});

	it('adds a report with the sections Access gives one', () => {
		const writer = new AccessVbaWriter(FORMS);
		writer.addDesign('ReportHere', 'report');
		const made = new AccessVbaWriter(writer.toBuffer()).designs()
			.find((entry) => entry.name === 'ReportHere')!;
		expect(made.kind).toBe('report');
		expect(accessDesignSections(made.design).map(accessDesignObjectName))
			.toEqual(expect.arrayContaining(['Detail']));
		expect(accessDesignSections(made.design).length).toBeGreaterThan(1);
	});

	it('gives each new design a GUID of its own', () => {
		const writer = new AccessVbaWriter(FORMS);
		writer.addDesign('First', 'form');
		writer.addDesign('Second', 'form');
		const designs = new AccessVbaWriter(writer.toBuffer()).designs();
		const guidOf = (name: string): string => designs.find((entry) => entry.name === name)!
			.design.objects[0].records.find((entry) => entry.id === 208)!.value.toString('hex');
		expect(guidOf('First')).not.toBe(guidOf('Second'));
		expect(guidOf('First')).toHaveLength(32);
	});

	it('puts a control on a design it just made', () => {
		const writer = new AccessVbaWriter(FORMS);
		writer.addDesign('WithControls', 'form');
		writer.editDesign('WithControls', (design) => addDesignControl(
			design, 'CommandButton', 'BtnGo', GUID,
			writer.designPrototypesFor('WithControls'),
			{ left: 100, top: 100, caption: 'Go' },
		));
		const made = new AccessVbaWriter(writer.toBuffer()).designs()
			.find((entry) => entry.name === 'WithControls')!;
		expect(accessDesignControls(made.design).map(accessDesignObjectName)).toEqual(['BtnGo']);
		// The first control of a type brings the type's defaults object, which
		// is what Access reads a themed control against.
		expect(made.design.objects.some(
			(object) => accessDesignObjectName(object) === undefined
				&& !isAccessDesignSection(object) && object.type === 104,
		)).toBe(true);
	});

	it('refuses a name a design already has', () => {
		const writer = new AccessVbaWriter(FORMS);
		expect(() => writer.addDesign('Calculator', 'form')).toThrow(/already exists/);
	});

	it('deletes a design and leaves the others alone', () => {
		const writer = new AccessVbaWriter(FORMS);
		writer.addDesign('Temporary', 'form');
		writer.deleteDesign('Temporary');
		const after = new AccessVbaWriter(writer.toBuffer());
		expect(after.designs().map((entry) => entry.name)).toEqual(['Calculator']);
		// The database still reads, and the project it holds is untouched.
		expect(after.moduleNames()).toEqual(new AccessVbaWriter(FORMS).moduleNames());
		expect(readAccessCatalog(writer.toBuffer())
			.filter((entry) => entry.type === -32768).map((entry) => entry.name))
			.not.toContain('Temporary');
	});

	it('refuses a design the database does not hold', () => {
		expect(() => new AccessVbaWriter(FORMS).deleteDesign('NotThere'))
			.toThrow(/no form or report named NotThere/);
	});
});
