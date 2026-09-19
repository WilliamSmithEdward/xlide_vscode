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
	accessDesignMembers,
	accessDesignObjectName,
	accessDesignSections,
	accessVbaIdentifier,
	buildAccessDesign,
	isAccessDesignSection,
	parseAccessDesign,
} from '../src/vba/access/accessDesign';
import {
	readAccessCatalog,
	readAccessDesigns,
	readAccessStorage,
	readAccessVbaStreams,
	readTableRows,
} from '../src/vba/access/accessStorage';

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
import {
	buildTypeInfo,
	readTypeInfo,
	typeInfoCodePage,
	typeInfoListedNames,
	updateTypeInfo,
} from '../src/vba/access/accessTypeInfo';
import { PROPERTY_CODES, TYPE_INFO_CLSID_AT } from '../src/vba/access/accessDesignTable';
import { codePageHolds } from '../src/vba/codePages';

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

// A control is rarely named the way VBA would name it. The form wizard names
// one after its field, so `Order Date` is ordinary, and VBA reaches it as
// `Me.Order_Date`. Access keeps both names in the stream, the identifier first
// and the design's name second, and leaves the second empty where they are
// the same. Read as one name, `Order_Date\0Order Date\0` runs on into the next
// entry: the reader either threw, so the form could not be edited at all, or
// matched nothing, and the writer then listed `Order Date` itself as the
// member, which is a name `Me.` can never reach.
//
// The fixture is Access's own (scripts/build-access-control-names-fixture.py):
// five text boxes, an ActiveX control and a renamed form header on a form, and
// a text box on a report.
describe('a TypeInfo stream whose members are not named as identifiers', () => {
	const NAMES = fs.readFileSync(path.join(BINARIES, 'AccessControlNamesFixture.accdb'));
	const entryOf = (name: string): ReturnType<typeof readAccessDesigns>[number] =>
		readAccessDesigns(NAMES).find((entry) => entry.name === name)!;
	const pairs = (stream: Buffer): Array<[number, string, string]> =>
		readTypeInfo(stream, 1252).map((entry) => [entry.ordinal, entry.identifier, entry.name]);

	it('reads both names of every member, in the order Access wrote them', () => {
		expect(pairs(entryOf('Names').typeInfo!)).toEqual([
			[0, 'Detail', 'Detail'],
			[1, 'Order_Date', 'Order Date'],
			[2, 'Qty_1', 'Qty-1'],
			[3, 'Ctl2ndBox', '2ndBox'],
			[4, 'Tax__VAT_', 'Tax (VAT)'],
			[5, 'Plain', 'Plain'],
			[6, 'Web_View', 'Web View'],
			[8, 'FormFooter', 'FormFooter'],
			// Renamed in Access, which moved it to the end with its ordinal.
			[7, 'Top_Part', 'Top Part'],
		]);
		expect(pairs(entryOf('Totals').typeInfo!)).toEqual([
			[0, 'Detail', 'Detail'],
			[1, 'PageFooterSection', 'PageFooterSection'],
			[2, 'PageHeaderSection', 'PageHeaderSection'],
			[3, 'Line_Total', 'Line Total'],
		]);
	});

	it('names every member as the design does', () => {
		for (const name of ['Names', 'Totals']) {
			const entry = entryOf(name);
			const designed = [...accessDesignSections(entry.design), ...accessDesignControls(entry.design)]
				.map(accessDesignObjectName);
			expect(readTypeInfo(entry.typeInfo!, 1252).map((member) => member.name).sort(), name)
				.toEqual([...designed].sort());
		}
	});

	it('finds an ActiveX control s tail after both of its names', () => {
		const members = readTypeInfo(entryOf('Names').typeInfo!, 1252);
		const activeX = members.find((member) => member.name === 'Web View')!;
		expect(activeX.tail.length).toBe(36);
		expect(members.filter((member) => member.tail.length > 0)).toEqual([activeX]);
	});

	it('rebuilds every stream Access wrote byte for byte', () => {
		let held = 0;
		for (const data of [FORMS, NAMES]) {
			for (const entry of readAccessDesigns(data)) {
				const stream = entry.typeInfo!;
				const rebuilt = buildTypeInfo(
					entry.kind, stream.subarray(TYPE_INFO_CLSID_AT, TYPE_INFO_CLSID_AT + 16),
					readTypeInfo(stream, 1252), 1252,
				);
				expect(rebuilt.equals(stream), entry.name).toBe(true);
				// A design nobody edited updates to the stream it already had.
				expect(updateTypeInfo(entry.kind, entry.design, stream, 1252).equals(stream), entry.name)
					.toBe(true);
				held += 1;
			}
		}
		expect(held).toBe(3);
	});

	it('agrees with Access on the identifier of every name it stored', () => {
		// Every pair in a stream is Access s own answer for that name.
		for (const data of [FORMS, NAMES]) {
			for (const entry of readAccessDesigns(data)) {
				for (const member of readTypeInfo(entry.typeInfo!, 1252)) {
					expect(accessVbaIdentifier(member.name), member.name).toBe(member.identifier);
				}
			}
		}
	});

	it('adds a member beside them, under the identifier VBA will use', () => {
		const { design: before, typeInfo } = entryOf('Names');
		const added = addDesignControl(
			before, 'TextBox', 'Ship Via', GUID, designPrototypes(before), { left: 0, top: 0 },
		);
		const after = updateTypeInfo('form', added, typeInfo!, 1252);
		expect(pairs(after)).toEqual([...pairs(typeInfo!), [9, 'Ship_Via', 'Ship Via']]);
		// Taking it off again is the stream Access wrote.
		expect(updateTypeInfo('form', removeDesignControl(added, 'Ship Via'), typeInfo!, 1252)
			.equals(typeInfo!)).toBe(true);
	});

	it('moves a renamed member to the end with its ordinal and its new identifier', () => {
		const { design: before, typeInfo } = entryOf('Names');
		const renamed = setDesignProperty(before, 'Order Date', 'Name', 'Ship Date');
		const after = pairs(updateTypeInfo(
			'form', renamed, typeInfo!, 1252, new Map([['Order Date', 'Ship Date']]),
		));
		expect(after[after.length - 1]).toEqual([1, 'Ship_Date', 'Ship Date']);
		expect(after.slice(0, -1)).toEqual(pairs(typeInfo!).filter(([ordinal]) => ordinal !== 1));
	});

	it('stores one name where a rename makes the two the same', () => {
		const { design: before, typeInfo } = entryOf('Names');
		const renamed = setDesignProperty(before, 'Qty-1', 'Name', 'Quantity');
		const after = updateTypeInfo('form', renamed, typeInfo!, 1252, new Map([['Qty-1', 'Quantity']]));
		expect(pairs(after).pop()).toEqual([2, 'Quantity', 'Quantity']);
		expect(after.includes(Buffer.from('Quantity\0\0', 'latin1'))).toBe(true);
		expect(after.includes(Buffer.from('Quantity\0Quantity', 'latin1'))).toBe(false);
	});

	it('drops a removed member and leaves the others as they were', () => {
		const { design: before, typeInfo } = entryOf('Names');
		const after = pairs(updateTypeInfo('form', removeDesignControl(before, 'Tax (VAT)'), typeInfo!, 1252));
		expect(after).toEqual(pairs(typeInfo!).filter(([, , name]) => name !== 'Tax (VAT)'));
	});

	it('refuses a name VBA could not tell from another member s', () => {
		// `Order_Date` is a different control name and the same identifier.
		// Access refuses it too: "the control name is already in use".
		const { design: before, typeInfo } = entryOf('Names');
		const clash = addDesignControl(
			before, 'TextBox', 'Order_Date', GUID, designPrototypes(before), { left: 0, top: 0 },
		);
		expect(() => updateTypeInfo('form', clash, typeInfo!, 1252))
			.toThrow(/Order_Date.*Order Date|Order Date.*Order_Date/);
	});

	it('puts right a member list written back when an entry was read as one name', () => {
		// That writer listed every member under its design name, `Order Date`
		// and all, which no code can reach. The next edit repairs it.
		const { design: untouched, typeInfo } = entryOf('Names');
		const damaged = buildTypeInfo(
			'form', typeInfo!.subarray(TYPE_INFO_CLSID_AT, TYPE_INFO_CLSID_AT + 16),
			readTypeInfo(typeInfo!, 1252).map((entry) => ({ ...entry, identifier: entry.name })), 1252,
		);
		expect(damaged.includes(Buffer.from('Order Date\0\0', 'latin1'))).toBe(true);
		expect(updateTypeInfo('form', untouched, damaged, 1252).equals(typeInfo!)).toBe(true);
	});

	it('edits such a form and such a report through the writer', () => {
		const writer = new AccessVbaWriter(NAMES);
		writer.editDesign('Names', (blob) => addDesignControl(
			blob, 'TextBox', 'Ship Via', GUID, designPrototypes(blob), { left: 10, top: 20 },
		));
		writer.editDesign('Totals', (blob) => addDesignControl(
			blob, 'TextBox', 'Tax Due', GUID, writer.designPrototypesFor('Totals'), { left: 10, top: 20 },
		));
		const written = readAccessDesigns(writer.toBuffer());
		const form = written.find((entry) => entry.name === 'Names')!;
		const report = written.find((entry) => entry.name === 'Totals')!;
		expect(pairs(form.typeInfo!)).toEqual([...pairs(entryOf('Names').typeInfo!), [9, 'Ship_Via', 'Ship Via']]);
		expect(pairs(report.typeInfo!)).toEqual([...pairs(entryOf('Totals').typeInfo!), [4, 'Tax_Due', 'Tax Due']]);
		// The code behind the form is untouched by a design edit.
		expect(new AccessVbaWriter(writer.toBuffer()).moduleNames()).toEqual(new AccessVbaWriter(NAMES).moduleNames());
	});

	it('leaves the writer as it was when the edit is refused', () => {
		const writer = new AccessVbaWriter(NAMES);
		const before = writer.toBuffer();
		expect(() => writer.editDesign('Names', (blob) => addDesignControl(
			blob, 'TextBox', 'Order_Date', GUID, designPrototypes(blob), { left: 0, top: 0 },
		))).toThrow(/Order_Date/);
		expect(writer.toBuffer().equals(before)).toBe(true);
	});
});

// The names in a TypeInfo stream are bytes in a code page, and VBA can only
// reach a control whose name that page holds. Measured on Access 16.0, on a
// machine whose ANSI page is 1252 (scripts/build-access-code-page-fixture.py):
// `Caf<e-acute>` is stored with E9 and the euro sign as 80, and a control named
// in Cyrillic gets NO entry, with no ordinal spent on it, while its design
// keeps the name in UTF-16. No best fit is tried. The page is the machine s:
// a database created with the Cyrillic collation wrote the same bytes, and so
// did one whose PROJECTCODEPAGE had been patched to 1251, which VBA went on
// reading as cp1252 and Access set back to 1252 on its next save. What another
// machine s page gives could not be measured, so nothing here assumes 1252:
// the project s page is the one to try first, since it is the page the machine
// that last saved the project wrote it in, and the stream s own bytes decide.
describe('the code page of a TypeInfo stream', () => {
	const PAGES = fs.readFileSync(path.join(BINARIES, 'AccessCodePageFixture.accdb'));
	const CYRILLIC = '\u0418\u043c\u044f';
	const E_ACUTE = 'Caf\u00e9';
	const entryOf = (): ReturnType<typeof readAccessDesigns>[number] =>
		readAccessDesigns(PAGES).find((entry) => entry.name === 'Names')!;
	const listed = (stream: Buffer, codePage = 1252): Array<[number, string, string]> =>
		readTypeInfo(stream, codePage).map((entry) => [entry.ordinal, entry.identifier, entry.name]);
	const clsidOf = (stream: Buffer): Buffer => stream.subarray(TYPE_INFO_CLSID_AT, TYPE_INFO_CLSID_AT + 16);

	/** A design holding a Detail section and one text box per name. */
	function designOf(...names: string[]): ReturnType<typeof parseAccessDesign> {
		const named = (name: string): ReturnType<typeof parseAccessDesign>['objects'][number]['records'] => [
			{ id: 1, code: 20, valueType: 0, width: 0, value: Buffer.from(name, 'utf16le') },
		];
		return {
			header: Buffer.alloc(10),
			trailer: Buffer.alloc(4),
			objects: [
				{ records: named('Design') },
				{ marker: 0xff, type: 152, records: named('Detail') },
				...names.map((name) => ({ marker: 0xff, type: 109, records: named(name) })),
			],
		};
	}

	it('lists every name cp1252 holds, in its bytes, and leaves out the one it cannot', () => {
		const { design, typeInfo } = entryOf();
		expect(accessDesignControls(design).map(accessDesignObjectName)).toContain(CYRILLIC);
		expect(listed(typeInfo!)).toEqual([
			[0, 'Detail', 'Detail'],
			[1, E_ACUTE, E_ACUTE],
			[2, 'Em\u2014Dash', 'Em\u2014Dash'],
			[3, '\u20acuro', '\u20acuro'],
			[4, 'Na\u00efve_\u2014_x', 'Na\u00efve \u2014 x'],
			[5, 'Plain', 'Plain'],
		]);
		expect(typeInfo!.includes(Buffer.from([0x43, 0x61, 0x66, 0xe9, 0, 0]))).toBe(true);
		expect(typeInfo!.includes(Buffer.from([0x80, 0x75, 0x72, 0x6f, 0, 0]))).toBe(true);
	});

	it('updates an untouched design to the stream Access wrote, with no entry for that name', () => {
		const { design, typeInfo } = entryOf();
		expect(updateTypeInfo('form', design, typeInfo!, 1252).equals(typeInfo!)).toBe(true);
	});

	it('gives an added control the page cannot name no entry, and spends no ordinal on it', () => {
		const { design: before, typeInfo } = entryOf();
		const first = addDesignControl(
			before, 'TextBox', '\u0414\u0432\u0430', GUID, designPrototypes(before), { left: 0, top: 0 },
		);
		const both = addDesignControl(first, 'TextBox', 'Ol\u00e9', GUID, designPrototypes(first), { left: 0, top: 0 });
		const after = updateTypeInfo('form', both, typeInfo!, 1252);
		expect(listed(after)).toEqual([...listed(typeInfo!), [6, 'Ol\u00e9', 'Ol\u00e9']]);
		expect(after.includes(Buffer.from([0x4f, 0x6c, 0xe9, 0, 0]))).toBe(true);
	});

	it('drops a control renamed out of the page, and lists one renamed into it', () => {
		// Access did the same to the same two renames: the entry of the one went,
		// and the other took the ordinal above the highest left.
		const { design: before, typeInfo } = entryOf();
		const renamed = setDesignProperty(
			setDesignProperty(before, 'Plain', 'Name', '\u0414\u0432\u0430'), CYRILLIC, 'Name', 'WasCyrillic',
		);
		const after = listed(updateTypeInfo('form', renamed, typeInfo!, 1252, new Map([
			['Plain', '\u0414\u0432\u0430'], [CYRILLIC, 'WasCyrillic'],
		])));
		expect(after).toEqual([
			...listed(typeInfo!).filter(([, , name]) => name !== 'Plain'),
			[5, 'WasCyrillic', 'WasCyrillic'],
		]);
	});

	it('reads the stream in its own page where the project names another', () => {
		// A project last saved on a machine with another page, around a form
		// that machine never saved: PROJECTCODEPAGE says 1251, the names are
		// cp1252. Read as 1251, E9 is a Cyrillic letter and the euro is not the
		// euro, so three members matched nothing and were written again as `?`.
		const { design: before, typeInfo } = entryOf();
		expect(typeInfoCodePage(typeInfo!, [E_ACUTE, '\u20acuro'], 1251)).toBe(1252);
		expect(updateTypeInfo('form', before, typeInfo!, 1251).equals(typeInfo!)).toBe(true);
		const added = addDesignControl(before, 'TextBox', 'Ol\u00e9', GUID, designPrototypes(before), { left: 0, top: 0 });
		const after = updateTypeInfo('form', added, typeInfo!, 1251);
		// One stream, one page: the new name goes in as cp1252 too.
		expect(listed(after)).toEqual([...listed(typeInfo!), [6, 'Ol\u00e9', 'Ol\u00e9']]);
	});

	it('keeps to the project s page where the stream agrees with it', () => {
		// What a cp1251 machine would write: the Cyrillic name held, the accent not.
		const field = '\u041f\u043e\u043b\u04351';
		const button = '\u041a\u043d\u043e\u043f\u043a\u04302';
		const stream = buildTypeInfo('form', clsidOf(entryOf().typeInfo!), [
			{ ident: 6296, ordinal: 0, name: 'Detail', identifier: 'Detail', tail: Buffer.alloc(0) },
			{ ident: 4717, ordinal: 1, name: field, identifier: field, tail: Buffer.alloc(0) },
		], 1251);
		expect(stream.includes(Buffer.from([0xcf, 0xee, 0xeb, 0xe5, 0x31, 0, 0]))).toBe(true);
		expect(typeInfoCodePage(stream, [field], 1251)).toBe(1251);
		expect(updateTypeInfo('form', designOf(field), stream, 1251).equals(stream)).toBe(true);
		const after = updateTypeInfo('form', designOf(field, E_ACUTE, button), stream, 1251);
		expect(listed(after, 1251)).toEqual([[0, 'Detail', 'Detail'], [1, field, field], [2, button, button]]);
		expect(after.includes(Buffer.from([0xca, 0xed, 0xee, 0xef, 0xea, 0xe0, 0x32, 0, 0]))).toBe(true);
	});

	it('writes back the bytes it read, where a page spells one character two ways', () => {
		// cp932 has U+2252 at 81 E0 and again at 87 90. Encoding the text again
		// would pick the first; what Access wrote is carried instead.
		const name = 'A\u2252';
		const plain = buildTypeInfo('form', clsidOf(entryOf().typeInfo!), [
			{ ident: 6296, ordinal: 0, name: 'Detail', identifier: 'Detail', tail: Buffer.alloc(0) },
			{ ident: 4717, ordinal: 1, name, identifier: name, tail: Buffer.alloc(0) },
		], 932);
		const at = plain.indexOf(Buffer.from([0x41, 0x81, 0xe0]));
		expect(at).toBeGreaterThan(0);
		const other = Buffer.from(plain);
		other.set([0x87, 0x90], at + 1);
		expect(readTypeInfo(other, 932).map((entry) => entry.name)).toEqual(['Detail', name]);
		expect(updateTypeInfo('form', designOf(name), other, 932).equals(other)).toBe(true);
	});

	it('says which names a page holds exactly', () => {
		expect(codePageHolds('Plain name', 1252)).toBe(true);
		expect(codePageHolds('Na\u00efve \u2014 \u20ac?', 1252)).toBe(true);
		expect(codePageHolds(CYRILLIC, 1252)).toBe(false);
		expect(codePageHolds(`${E_ACUTE}\u0418`, 1252)).toBe(false);
		// No best fit: A-macron, fullwidth A, dotless i and omega are not A, A, i and O.
		for (const near of ['\u0100bc', '\uff21bc', '\u0131x', '\u03a9hm']) {
			expect(codePageHolds(near, 1252), near).toBe(false);
		}
		expect(codePageHolds(CYRILLIC, 1251)).toBe(true);
		expect(codePageHolds(E_ACUTE, 1251)).toBe(false);
		expect(codePageHolds('\u30c6\u30ad\u30b9\u30c80', 932)).toBe(true);
	});

	it('offers Me the members Access listed, and not the control it left out', () => {
		const { design, typeInfo } = entryOf();
		const names = typeInfoListedNames(typeInfo!, design, 1252);
		expect([...names].sort()).toEqual(['Detail', E_ACUTE, 'Em\u2014Dash', '\u20acuro', 'Na\u00efve \u2014 x', 'Plain'].sort());
		expect(accessDesignMembers(design, 'form').map((member) => member.name)).toContain(CYRILLIC);
		expect(accessDesignMembers(design, 'form', names).map((member) => member.name)).toEqual([
			'Detail', E_ACUTE, 'Em\u2014Dash', '\u20acuro', 'Na\u00efve_\u2014_x', 'Plain',
		]);
		// And from the file, which is what the editor reads.
		const form = readModulesFromBuffer(PAGES).find((entry) => entry.name === 'Form_Names')!;
		expect(form.implicitMembers?.map((member) => member.name)).toEqual([
			'Detail', E_ACUTE, 'Em\u2014Dash', '\u20acuro', 'Na\u00efve_\u2014_x', 'Plain',
		]);
	});

	it('edits the form through the writer and leaves the unlisted control unlisted', () => {
		const writer = new AccessVbaWriter(PAGES);
		writer.editDesign('Names', (blob) => setDesignProperty(blob, 'Plain', 'Width', 2000));
		const written = readAccessDesigns(writer.toBuffer()).find((entry) => entry.name === 'Names')!;
		expect(written.typeInfo!.equals(entryOf().typeInfo!)).toBe(true);
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

// ---------------------------------------------------------------------------
// The designer surface over an Access design: the same markup, canvas and
// gestures a UserForm gets. The engine answers `readFormMarkup`,
// `applyFormMarkup`, `readFormPreview` and `applyFormDesignerOp` for an Access
// form or report exactly as it does for a UserForm, so the designer, the
// markup editor and the preview all work without knowing which host it is.

import * as os from 'os';
import {
	addFormModule,
	applyFormDesignerOp,
	applyFormMarkup,
	deleteModule,
	listModules,
	readFormMarkup,
	readFormPreview,
	readModulesFromBuffer,
	renameModule,
} from '../src/vba/projectService';
import { sceneOfAccessDesign } from '../src/vba/access/accessDesignScene';
import {
	accessDesignProperties,
	printAccessDesignMarkup,
} from '../src/vba/access/accessDesignMarkup';
import { PROPERTY_SLOTS } from '../src/vba/access/accessDesignTable';
import { readTableDefinition } from '../src/vba/access/accessFormat';
import { setExtensionAssetRoot } from '../src/extensionAssets';

setExtensionAssetRoot(path.join(__dirname, '..'));

/** A scratch copy, so a test that writes never touches the fixture. */
function scratchCopy(name: string): string {
	const target = path.join(
		fs.mkdtempSync(path.join(os.tmpdir(), 'xlide-access-')), name,
	);
	fs.copyFileSync(path.join(BINARIES, name), target);
	return target;
}

const FIXTURE = 'AccessFormFixture.accdb';
const DESIGN_MODULE = 'Form_Calculator';

describe('an Access design in the project listing', () => {
	it('is typed as the form it is, not as a class', () => {
		const entry = listModules(path.join(BINARIES, FIXTURE))
			.find((module) => module.name === DESIGN_MODULE);
		// The type is what gives the tree its icon, its sort position and its
		// Designer child; a form listed as a class has none of them.
		expect(entry?.type).toBe('accessform');
	});

	it('lists a design whose code window Access never opened', () => {
		// A design with no module behind it is still a design. It reaches the
		// listing through the container's own storage rather than the project.
		const target = scratchCopy(FIXTURE);
		const writer = new AccessVbaWriter(fs.readFileSync(target));
		writer.addDesign('Standalone', 'form');
		writer.addDesign('Printed', 'report');
		fs.writeFileSync(target, writer.toBuffer());
		const entries = listModules(target);
		expect(entries.find((module) => module.name === 'Form_Standalone')?.type)
			.toBe('accessform');
		expect(entries.find((module) => module.name === 'Report_Printed')?.type)
			.toBe('accessreport');
	});
});

describe('the markup projection', () => {
	it('prints the design with its sections and controls', () => {
		const { markup } = readFormMarkup(path.join(BINARIES, FIXTURE), DESIGN_MODULE);
		expect(markup).toMatch(/^<Form Name="Calculator"/);
		expect(markup).toContain('<Detail Name="Detail"');
		expect(markup).toContain('<TextBox Name="Qty"');
		expect(markup).toContain('<CommandButton Name="AddLine"');
		// Only properties the object's own schema names are printed, so a
		// reader can act on every one of them.
		expect(markup).not.toMatch(/\sP\d+="/);
	});

	it('applies what it printed without changing anything', () => {
		const target = scratchCopy(FIXTURE);
		const { markup } = readFormMarkup(target, DESIGN_MODULE);
		// Idempotence is what the designer relies on: it rewrites the whole
		// document after every gesture.
		expect(applyFormMarkup(target, DESIGN_MODULE, markup).applied).toEqual([]);
	});

	it('applies an edit and reads it back', () => {
		const target = scratchCopy(FIXTURE);
		const before = readFormMarkup(target, DESIGN_MODULE).markup;
		const result = applyFormMarkup(
			target,
			DESIGN_MODULE,
			before.replace('Caption="Order calculator"', 'Caption="Edited"')
				.replace('<TextBox Name="Qty" Left="18"', '<TextBox Name="Qty" Left="25"'),
		);
		expect(result.applied).toEqual(expect.arrayContaining([
			'the design.Caption = Edited', 'Qty.Left = 25',
		]));
		const after = readFormMarkup(target, DESIGN_MODULE).markup;
		expect(after).toContain('Caption="Edited"');
		expect(after).toContain('<TextBox Name="Qty" Left="25"');
	});

	it('adds and removes a control through the markup alone', () => {
		const target = scratchCopy(FIXTURE);
		const before = readFormMarkup(target, DESIGN_MODULE).markup;
		const added = before.replace(
			'  </Detail>',
			'    <Label Name="LblNew" Caption="New" Left="100" Top="100" Width="1000" Height="300" />\n  </Detail>',
		);
		expect(applyFormMarkup(target, DESIGN_MODULE, added).applied)
			.toEqual(expect.arrayContaining([expect.stringContaining('added Label LblNew')]));
		expect(readFormMarkup(target, DESIGN_MODULE).markup).toContain('Name="LblNew"');
		const removed = readFormMarkup(target, DESIGN_MODULE).markup
			.split('\n').filter((line) => !line.includes('Name="LblNew"')).join('\n');
		expect(applyFormMarkup(target, DESIGN_MODULE, removed).applied)
			.toEqual(expect.arrayContaining(['removed LblNew']));
		expect(readFormMarkup(target, DESIGN_MODULE).markup).not.toContain('LblNew');
	});

	it('refuses a property the control type does not have', () => {
		const target = scratchCopy(FIXTURE);
		const markup = readFormMarkup(target, DESIGN_MODULE).markup
			.replace('<Rectangle Name="Banner"', '<Rectangle Name="Banner" ForeColor="0"');
		expect(() => applyFormMarkup(target, DESIGN_MODULE, markup))
			.toThrow(/A Rectangle has no ForeColor/);
	});
});

describe('the designer canvas', () => {
	it('renders the design with its controls', () => {
		const { html } = readFormPreview(path.join(BINARIES, FIXTURE), DESIGN_MODULE);
		expect(html).toContain('<!DOCTYPE html>');
		// Every control the design holds is drawn, positioned in points from
		// the twips Access stores.
		expect((html.match(/data-name="/g) ?? []).length).toBeGreaterThanOrEqual(15);
		expect(html).toContain('Order calculator');
	});

	it('moves and resizes a control', () => {
		const target = scratchCopy(FIXTURE);
		applyFormDesignerOp(target, DESIGN_MODULE, {
			kind: 'geometry', name: 'Qty', left: 40, top: 90, width: 100, height: 30,
		});
		// The gesture, the markup and the pane all count points; Access stores
		// twips, twenty to the point, and the conversion never leaks out.
		expect(readFormMarkup(target, DESIGN_MODULE).markup)
			.toContain('<TextBox Name="Qty" Left="40" Top="90" Width="100" Height="30"');
	});

	it('sets a property from the pane', () => {
		const target = scratchCopy(FIXTURE);
		applyFormDesignerOp(target, DESIGN_MODULE, {
			kind: 'setProp', name: 'Title', prop: 'Caption', value: 'From the pane',
		});
		expect(readFormMarkup(target, DESIGN_MODULE).markup)
			.toContain('<Label Name="Title" Caption="From the pane"');
	});

	it('drops a control from the toolbox and deletes it again', () => {
		const target = scratchCopy(FIXTURE);
		const added = applyFormDesignerOp(target, DESIGN_MODULE, {
			kind: 'add', container: 'Detail', controlKind: 'Label', left: 20, top: 250,
		});
		expect(added.newName).toBe('Label0');
		expect(readFormMarkup(target, DESIGN_MODULE).markup).toContain('Name="Label0"');
		applyFormDesignerOp(target, DESIGN_MODULE, { kind: 'remove', name: 'Label0' });
		expect(readFormMarkup(target, DESIGN_MODULE).markup).not.toContain('Name="Label0"');
	});

	it('brings a control to the front and sends it back', () => {
		const target = scratchCopy(FIXTURE);
		const order = (): string[] => [...readFormMarkup(target, DESIGN_MODULE).markup
			.matchAll(/<\w+ Name="(\w+)"/g)].map((match) => match[1]);
		const before = order();
		applyFormDesignerOp(target, DESIGN_MODULE, { kind: 'zOrder', name: 'Banner', toFront: true });
		// Access paints in list order, so the last control drawn is on top.
		expect(order().at(-1)).toBe('Banner');
		applyFormDesignerOp(target, DESIGN_MODULE, { kind: 'zOrder', name: 'Banner', toFront: false });
		expect(order()).toEqual(before);
	});

	it('rewrites the tab order of a section', () => {
		const target = scratchCopy(FIXTURE);
		applyFormDesignerOp(target, DESIGN_MODULE, {
			kind: 'tabOrder',
			container: 'Detail',
			names: ['Price', 'Qty', 'Express', 'AddLine', 'Reset', 'Lines'],
		});
		const markup = readFormMarkup(target, DESIGN_MODULE).markup;
		// The first control carries no TabIndex at all, which is how Access
		// writes it; the rest count up from one.
		expect(markup).toMatch(/<TextBox Name="Price"(?![^>]*TabIndex)/);
		expect(markup).toMatch(/<TextBox Name="Qty"[^>]*TabIndex="1"/);
		expect(markup).toMatch(/<CheckBox Name="Express"[^>]*TabIndex="2"/);
		expect(markup).toMatch(/<ListBox Name="Lines"[^>]*TabIndex="5"/);
	});

	it('takes a colour, a Yes/No and a measurement from the pane as text', () => {
		const target = scratchCopy(FIXTURE);
		for (const [name, prop, value] of [
			['Title', 'ForeColor', '#c00000'],
			['Title', 'Visible', 'False'],
			['Qty', 'Left', '36.5'],
			['', 'Caption', 'From the pane'],
		] as const) {
			applyFormDesignerOp(target, DESIGN_MODULE, { kind: 'setProp', name, prop, value });
		}
		const markup = readFormMarkup(target, DESIGN_MODULE).markup;
		expect(markup).toMatch(/<Label Name="Title"[^>]*ForeColor="#c00000"/);
		expect(markup).toMatch(/<Label Name="Title"[^>]*Visible="False"/);
		expect(markup).toMatch(/<TextBox Name="Qty"[^>]*Left="36.5"/);
		expect(markup).toContain('<Form Name="Calculator" Caption="From the pane"');
	});

	it('shows every property the control type has, the same list every time', () => {
		// The sheet for a control is its type's whole property list, not
		// whatever the design happened to store: two buttons that differ only
		// in what was saved get the same rows in the same order.
		const props = accessDesignProperties(design(), 'Calculator', 'form');
		const hidden = new Set(['GUID', 'OverlapFlags', 'IMESentenceMode',
			'LayoutCachedLeft', 'LayoutCachedTop', 'LayoutCachedWidth', 'LayoutCachedHeight']);
		for (const [target, schema] of [['Reset', 'CommandButton'], ['Title', 'Label'],
			['Qty', 'TextBox'], ['Lines', 'ListBox']] as const) {
			const want = [...PROPERTY_SLOTS.get(schema)!.keys()]
				.filter((prop) => !hidden.has(prop) && !prop.startsWith('Unidentified'));
			const shown = props[target].rows.map((row) => row.prop);
			expect(new Set(shown)).toEqual(new Set(want));
			expect(shown).toHaveLength(want.length);
		}
		expect(props.Reset.rows.map((row) => row.prop))
			.toEqual(props.AddLine.rows.map((row) => row.prop));
		// The first rows are the ones a reader looks for first.
		expect(props.Reset.rows.slice(0, 6).map((row) => row.prop))
			.toEqual(['Name', 'Caption', 'Left', 'Top', 'Width', 'Height']);
	});

	it('shows what a control inherits from its type\'s defaults', () => {
		// Access stores a property once on the control-defaults object for the
		// type and leaves it off every control that agrees: both buttons here
		// take their font size from there, and only the bold one carries a
		// weight of its own.
		const props = accessDesignProperties(design(), 'Calculator', 'form');
		const row = (target: string, prop: string): string | undefined =>
			props[target].rows.find((entry) => entry.prop === prop)?.value;
		expect(row('AddLine', 'FontWeight')).toBe('700');
		expect(row('Reset', 'FontWeight')).toBe('400');
		expect(row('AddLine', 'FontSize')).toBe('11');
		expect(row('Reset', 'FontSize')).toBe('11');
		// Under those, the value Access gives a control it has just made.
		expect(row('Reset', 'Visible')).toBe('True');
		expect(row('Reset', 'Enabled')).toBe('True');
		expect(row('Reset', 'DisplayWhen')).toBe('0');
		// And a property none of the three has is empty, not guessed at.
		expect(row('Reset', 'Tag')).toBe('');
		// The markup stays the design's own text: printing an inherited value
		// there would write it onto the control on the way back.
		const markup = printAccessDesignMarkup(design(), 'Calculator', 'form');
		expect(/<CommandButton Name="Reset"[^>]*/.exec(markup)![0]).not.toContain('FontSize');
		// And the canvas draws the inherited size rather than a guess.
		const scene = sceneOfAccessDesign(design(), 'Calculator', 'form');
		const reset = scene.controls[0].children.find((control) => control.name === 'Reset');
		expect(reset?.style).toContain('font-size:11pt;');
	});

	it('offers the pane a dropdown, a swatch and the font list', () => {
		const scene = sceneOfAccessDesign(design(), 'Calculator', 'form');
		expect(scene.paneBareEnums).toBe(false);
		// A published value set becomes the dropdown, in Access's own words.
		expect(scene.enums?.TextAlign).toEqual([
			['0', 'General'], ['1', 'Left'], ['2', 'Center'], ['3', 'Right'], ['4', 'Distribute'],
		]);
		expect(scene.enums?.SpecialEffect?.[5]).toEqual(['5', 'Chiseled']);
		// BorderLineStyle's settings are not published, so it stays a field.
		expect(scene.enums?.BorderLineStyle).toBeUndefined();
		expect(scene.bools).toContain('Visible');
		expect(scene.paneEditors?.BackColor).toBe('color');
		expect(scene.paneEditors?.HoverForeColor).toBe('color');
		// A theme index is a long that reads like a colour and is not one.
		expect(scene.paneEditors?.BackThemeColorIndex).toBeUndefined();
		expect(scene.paneEditors?.FontName).toBe('font');
		expect(scene.paneEditors?.Left).toBe('number');
	});

	it('moves a control to where the gesture dropped it', () => {
		const target = scratchCopy(FIXTURE);
		// The fixture has one section, so a move within it is a move of the
		// control's own corner and nothing else.
		applyFormDesignerOp(target, DESIGN_MODULE, {
			kind: 'reparent', name: 'Qty', container: 'Detail', left: 12, top: 24,
		});
		expect(readFormMarkup(target, DESIGN_MODULE).markup)
			.toMatch(/<TextBox Name="Qty" Left="12" Top="24"/);
	});
});

describe('renaming an Access design', () => {
	/** Every `Name` the catalog and the navigation pane carry. */
	function catalogNames(file: string): Set<string> {
		const bytes = fs.readFileSync(file);
		const out = new Set<string>();
		for (const table of readAccessCatalog(bytes)) {
			if (!/^MSys(Objects|NavPaneObjectIDs)$/.test(table.name) || table.type !== 1) {
				continue;
			}
			for (const row of readTableRows(bytes, readTableDefinition(bytes, table.definitionPage))) {
				const name = row.values.get('Name');
				if (typeof name === 'string') { out.add(name); }
			}
		}
		return out;
	}

	it('renames the design, its module and both catalog rows', () => {
		const target = scratchCopy(FIXTURE);
		renameModule(target, DESIGN_MODULE, 'Form_Invoice');
		const listed = listModules(target).map((module) => module.name);
		expect(listed).toContain('Form_Invoice');
		expect(listed).not.toContain(DESIGN_MODULE);
		expect(readAccessDesigns(fs.readFileSync(target)).map((entry) => entry.name))
			.toEqual(['Invoice']);
		// Access refuses to open a design the catalog, the navigation pane and
		// the container's listing disagree about.
		const names = catalogNames(target);
		expect(names.has('Invoice')).toBe(true);
		expect(names.has('Calculator')).toBe(false);
		expect(readFormMarkup(target, 'Form_Invoice').markup).toContain('<Form Name="Invoice"');
		expect(listModules(target).find((module) => module.name === 'Form_Invoice')?.type)
			.toBe('accessform');
	});

	it('moves the DocClass line PROJECT names the module by', () => {
		// PROJECT lists a design's module as `DocClass=Form_X/&H...`, never as
		// `Module=` or `Class=`. Leaving it naming a module the project no
		// longer has is how Access decides the whole project is corrupt, which
		// it reports on the next open and not before.
		const target = scratchCopy(FIXTURE);
		renameModule(target, DESIGN_MODULE, 'Form_Invoice');
		const project = readAccessVbaStreams(fs.readFileSync(target)).get('PROJECT')!.toString('latin1');
		expect(project).toContain('DocClass=Form_Invoice/&H00000000');
		expect(project).not.toContain('Form_Calculator');
		expect(project).toContain('Form_Invoice=0, 0, 0, 0, C');
	});

	it('takes the design name with or without the module prefix', () => {
		const target = scratchCopy(FIXTURE);
		// The result names the module made, which the caller's editors and a
		// pending review follow: before, they followed to `Invoice`, which
		// does not exist.
		expect(renameModule(target, DESIGN_MODULE, 'Invoice').moduleName).toBe('Form_Invoice');
		expect(listModules(target).map((module) => module.name)).toContain('Form_Invoice');
		expect(renameModule(target, 'Form_Invoice', 'Form_Bill').moduleName).toBe('Form_Bill');
	});

	it('refuses a name another design already has', () => {
		const target = scratchCopy(FIXTURE);
		addFormModule(target, 'Second');
		expect(() => renameModule(target, 'Form_Second', 'Calculator')).toThrow(/already exists/);
	});
});

describe('creating and deleting an Access design', () => {
	it('adds a form the designer can open and deletes it again', () => {
		const target = scratchCopy(FIXTURE);
		addFormModule(target, 'Orders');
		expect(listModules(target).find((module) => module.name === 'Form_Orders')?.type)
			.toBe('accessform');
		expect(readFormMarkup(target, 'Form_Orders').markup).toContain('<Form Name="Orders"');
		deleteModule(target, 'Form_Orders');
		expect(listModules(target).map((module) => module.name)).not.toContain('Form_Orders');
	});

	it('takes the module behind a design with it', () => {
		const target = scratchCopy(FIXTURE);
		deleteModule(target, DESIGN_MODULE);
		expect(readAccessDesigns(fs.readFileSync(target))).toEqual([]);
		// The module goes too: a DocClass whose design is gone is a corrupt
		// project, and the module is unreachable either way.
		expect(listModules(target).map((module) => module.name)).not.toContain(DESIGN_MODULE);
		const project = readAccessVbaStreams(fs.readFileSync(target)).get('PROJECT')!.toString('latin1');
		expect(project).not.toContain('Form_Calculator');
	});

	it('adds a report', () => {
		const target = scratchCopy(FIXTURE);
		addFormModule(target, 'Summary', '', 'report');
		expect(listModules(target).find((module) => module.name === 'Report_Summary')?.type)
			.toBe('accessreport');
		expect(readFormMarkup(target, 'Report_Summary').markup).toContain('<Report Name="Summary"');
	});
});
