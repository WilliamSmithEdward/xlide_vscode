// The ActiveX controls on a presentation's slides (issue #1555).
// PowerPointControlsFixture.pptm was saved with ActiveX enabled, so each
// p:control names its control and the slide module's header lists it;
// PowerPointControlsDisabledSaveFixture.pptm was saved after a VBA edit with
// ActiveX disabled, so only the fallback picture names it.

import * as fs from 'node:fs';
import * as path from 'node:path';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { slideControlNames, slideModuleParts, slidePartsById } from '../../src/vba/pptShapes';
import { listModules, readModulesFromBuffer } from '../../src/vba/projectService';
import { ZipArchive } from '../../src/vba/zip';

const FIXTURES = path.join(__dirname, '..', 'fixtures', 'binaries');

function zipOf(file: string): ZipArchive {
	return ZipArchive.read(fs.readFileSync(path.join(FIXTURES, file)));
}

/** The slide parts `run` reads, whole. */
function slideReads(run: () => void): string[] {
	const read = vi.spyOn(ZipArchive.prototype, 'read');
	run();
	return read.mock.calls.map(([name]) => String(name)).filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name)).sort();
}

afterEach(() => {
	vi.restoreAllMocks();
});

describe('the ActiveX controls on a presentation\'s slides', () => {
	it('finds each slide\'s part by its ID, however the slides were reordered', () => {
		expect(slidePartsById(zipOf('PowerPointControlsDisabledSaveFixture.pptm'))).toEqual(new Map([
			[258, 'ppt/slides/slide1.xml'], [260, 'ppt/slides/slide2.xml'], [256, 'ppt/slides/slide3.xml'], [259, 'ppt/slides/slide4.xml'],
		]));
	});

	it('reads a control named on its p:control once, though it is written twice', () => {
		const zip = zipOf('PowerPointControlsFixture.pptm');
		expect(slideControlNames(zip, slidePartsById(zip)!.get(256)!)).toEqual(['CommandButton1']);
	});

	it('reads a control named only on its fallback picture', () => {
		const zip = zipOf('PowerPointControlsDisabledSaveFixture.pptm');
		const slides = slidePartsById(zip)!;
		expect(new Map([...slides].map(([id, part]) => [id, slideControlNames(zip, part)]))).toEqual(new Map([
			[256, ['FirstBtn']], [258, ['ThirdBtn']], [259, ['LastBtn']], [260, ['NewBtn']],
		]));
	});

	it('does not read a slide that relates to no control\'s part', () => {
		const zip = zipOf('PowerPointControlsFixture.pptm');
		const part = slidePartsById(zip)!.get(256)!;
		const rels = part.replace(/([^/]+)$/, '_rels/$1.rels');
		zip.write(rels, Buffer.from(zip.read(rels).toString('utf8').replace(/<Relationship\b[^>]*relationships\/control"[^>]*\/>/g, ''), 'utf8'));
		expect(slideReads(() => expect(slideControlNames(zip, part)).toEqual([]))).toEqual([]);
	});

	it('says nothing for a slide with a control named in neither copy', () => {
		const zip = zipOf('PowerPointControlsDisabledSaveFixture.pptm');
		const part = slidePartsById(zip)!.get(258)!;
		zip.write(part, Buffer.from(zip.read(part).toString('utf8').replace(/(<p:cNvPr\b[^>]*?) name="ThirdBtn"/, '$1'), 'utf8'));
		expect(slideControlNames(zip, part)).toBeUndefined();
	});

	it('matches each slide module to the slide whose ID its name carries, or none of them', () => {
		const slides = new Map([[256, 'a.xml'], [258, 'b.xml']]);
		expect(slideModuleParts(['Slide1', 'Slide3'], slides)).toEqual(new Map([['slide1', 'a.xml'], ['slide3', 'b.xml']]));
		// A module renamed in the VBE, one PowerPoint would not name so, and one with no slide.
		expect(slideModuleParts(['Intro', 'Slide1'], slides)).toBeUndefined();
		expect(slideModuleParts(['Slide01'], slides)).toBeUndefined();
		expect(slideModuleParts(['Slide1', 'Slide2'], slides)).toBeUndefined();
	});

	it('says nothing for a package with no presentation', () => {
		expect(slidePartsById(zipOf('WordFixture.docm'))).toBeUndefined();
	});

	it('reads no slide when every slide module\'s header lists its controls', () => {
		const data = fs.readFileSync(path.join(FIXTURES, 'PowerPointControlsFixture.pptm'));
		expect(slideReads(() => readModulesFromBuffer(data, true))).toEqual([]);
	});

	it('gives the modules their controls when the project is listed from the file', () => {
		const modules = listModules(path.join(FIXTURES, 'PowerPointControlsDisabledSaveFixture.pptm'));
		const slide3 = modules.find((entry) => entry.name === 'Slide3');
		expect(slide3?.designerClass).toBe('PowerPoint.Slide');
		expect(slide3?.implicitMembers).toEqual([{ name: 'ThirdBtn', type: 'Object' }]);
	});

	it('reads each slide once when its module\'s header lists none', () => {
		const data = fs.readFileSync(path.join(FIXTURES, 'PowerPointControlsDisabledSaveFixture.pptm'));
		const reads = slideReads(() => readModulesFromBuffer(data, true));
		expect(reads).toEqual([...new Set(reads)]);
		expect(reads).toHaveLength(4);
	});
});
