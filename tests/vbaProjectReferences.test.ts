import { describe, expect, it } from 'vitest';
import * as fs from 'fs';
import * as path from 'path';
import { Cfb } from '../src/vba/cfb';
import { XlsxWorkbook } from '../src/vba/xlsx';
import { decompress } from '../src/vba/ovba';
import { readDirRecords, REC_PROJECTMODULES } from '../src/vba/vbaProject';
import {
	buildControlReference,
	buildMsFormsReference,
	guidBytes,
	hasMsFormsReference,
	insertReferenceRecords,
	libid,
	MSFORMS_TYPELIB_GUID,
	readProjectReferences,
} from '../src/vba/vbaProjectReferences';

// A project holding a UserForm must reference the Microsoft Forms library or
// the form cannot be instantiated, and Excel reports the workbook as broken
// while every reader here finds it healthy. The oracle for the record shapes
// is a workbook Excel itself wrote: the builder must reproduce its reference
// block byte for byte from that block's own fields.

const FIXTURE = path.join('tests', 'fixtures', 'binaries', 'FormFixtureVbide.xlsm');

function dirStream(file: string): Buffer {
	const cfb = Cfb.fromBytes(XlsxWorkbook.fromBuffer(fs.readFileSync(file)).readVbaProject());
	return decompress(cfb.getStreamInStorage('VBA', 'dir'), 'VBA/dir');
}

/** The span of the fixture's Microsoft Forms records, from its name to the modules. */
function msFormsSpan(dir: Buffer): { start: number; end: number } {
	const records = readDirRecords(dir);
	const start = records.find((r) => r.id === 0x0016
		&& dir.subarray(r.dataStart, r.dataEnd).toString('latin1') === 'MSForms');
	const modules = records.find((r) => r.id === REC_PROJECTMODULES);
	expect(start, 'fixture has a Microsoft Forms reference').toBeDefined();
	expect(modules, 'fixture has a module section').toBeDefined();
	return { start: start!.start, end: modules!.start };
}

describe('the references a project declares', () => {
	it('reads them from a workbook Excel wrote', () => {
		const references = readProjectReferences(dirStream(FIXTURE));
		// A control reference is one reference, whatever number of records carries it.
		expect(references.map((r) => r.name)).toEqual(['stdole', 'Office', 'MSForms']);
		expect(references[0]).toMatchObject({ kind: 'registered' });
		expect(references[0].libid).toContain('{00020430-0000-0000-C000-000000000046}');
		expect(references[0].libid).toContain('OLE Automation');
		// A form's library is a control reference, not a registered one.
		expect(references.find((r) => r.name === 'MSForms')?.kind).toBe('control');
		expect(references.find((r) => r.name === 'MSForms')?.libid).toContain(MSFORMS_TYPELIB_GUID);
	});

	it('sees when a project has no Microsoft Forms reference', () => {
		const dir = dirStream(FIXTURE);
		expect(hasMsFormsReference(dir)).toBe(true);
		const span = msFormsSpan(dir);
		const without = Buffer.concat([dir.subarray(0, span.start), dir.subarray(span.end)]);
		expect(hasMsFormsReference(without)).toBe(false);
		// Putting it back is what the engine does when it adds a form.
		expect(hasMsFormsReference(insertReferenceRecords(without, dir.subarray(span.start, span.end)))).toBe(true);
	});

	it('rebuilds the fixture s own reference block byte for byte', () => {
		const dir = dirStream(FIXTURE);
		const span = msFormsSpan(dir);
		const block = dir.subarray(span.start, span.end);
		// Read the block's own fields back out, then build from exactly those.
		const records = readDirRecords(dir).filter((r) => r.start >= span.start && r.start < span.end);
		const body = (id: number, nth = 0): Buffer => {
			const found = records.filter((r) => r.id === id)[nth];
			return dir.subarray(found.dataStart, found.dataEnd);
		};
		const sized = (b: Buffer): string => b.subarray(4, 4 + b.readUInt32LE(0)).toString('latin1');
		const extended = body(0x0030);
		const extendedLibid = sized(extended);
		const guidAt = 4 + Buffer.from(extendedLibid, 'latin1').length + 6;
		const rebuilt = buildControlReference({
			name: body(0x0016).toString('latin1'),
			libidOriginal: body(0x0033).toString('latin1'),
			libidTwiddled: sized(body(0x002f)),
			libidExtended: extendedLibid,
			typeLibGuid: MSFORMS_TYPELIB_GUID,
			cookie: extended.readUInt32LE(guidAt + 16),
		});
		expect(rebuilt.length).toBe(block.length);
		expect(rebuilt.equals(block)).toBe(true);
		// The GUID really is the one the fixture stores, in OLE's mixed-endian layout.
		expect(extended.subarray(guidAt, guidAt + 16).equals(guidBytes(MSFORMS_TYPELIB_GUID))).toBe(true);
	});

	it('builds a Microsoft Forms reference the reader recognises', () => {
		const dir = dirStream(FIXTURE);
		const span = msFormsSpan(dir);
		const without = Buffer.concat([dir.subarray(0, span.start), dir.subarray(span.end)]);
		const patched = insertReferenceRecords(without, buildMsFormsReference());
		expect(hasMsFormsReference(patched)).toBe(true);
		const added = readProjectReferences(patched).find((r) => r.name === 'MSForms');
		expect(added?.kind).toBe('control');
		expect(added?.libid).toContain(MSFORMS_TYPELIB_GUID);
		expect(added?.libid).toContain('Microsoft Forms 2.0 Object Library');
		// Everything else the project declared survives the insert.
		expect(readProjectReferences(patched).map((r) => r.name)).toEqual(['stdole', 'Office', 'MSForms']);
		// The module section still parses, which is what the format requires of the position.
		expect(readDirRecords(patched).some((r) => r.id === REC_PROJECTMODULES)).toBe(true);
	});

	it('spells a libid and a GUID the way the format does', () => {
		expect(libid('{0D452EE1-E08F-101A-852E-02608C4D0BB4}', '2.0', 'C:\\x\\FM20.DLL', 'Microsoft Forms 2.0 Object Library'))
			.toBe('*\\G{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0#C:\\x\\FM20.DLL#Microsoft Forms 2.0 Object Library');
		expect(guidBytes('{0D452EE1-E08F-101A-852E-02608C4D0BB4}').toString('hex'))
			.toBe('e12e450d8fe01a10852e02608c4d0bb4');
		expect(() => guidBytes('not-a-guid')).toThrow(/Not a GUID/);
	});
});
