import { describe, expect, it } from 'vitest';
import * as fs from 'fs';
import * as path from 'path';
import { parseAccessDesign } from '../src/vba/access/accessDesign';
import { CONTROL_TYPES } from '../src/vba/access/accessDesignTable';
import {
	accessDesignTemplate,
	availablePrototypes,
	rethemedAccessDesign,
	rethemedAccessPrototypes,
	withDesignGuid,
} from '../src/vba/access/accessDesignTemplates';
import {
	accessTintedColor,
	OFFICE_2023_THEME,
	readAccessDesignTheme,
	type AccessTheme,
} from '../src/vba/access/accessThemes';
import { AccessVbaWriter } from '../src/vba/access/accessVbaWriter';

// Issue #151, from pyOpenVBA 6.3.0: Access draws a new form, report or
// control in the theme its database keeps, and a database with none gets
// Office's 2023 theme with its first form. XLIDE's templates were captured on
// the 2023 theme, so a design for a database on it needs no change, and one
// for a database on another theme is worked out again from the theme.

const FIXTURES = path.join(__dirname, 'fixtures');
const OFFICE_2007: AccessTheme = {
	majorFont: 'Cambria',
	minorFont: 'Calibri',
	colors: ['000000', 'FFFFFF', '1F497D', 'EEECE1', '4F81BD', 'C0504D', '9BBB59', '8064A2', '4BACC6', 'F79646', '0000FF', '800080'],
};

describe('theme colours as Access works them out (issue #151)', () => {
	// Access's answer for every slot of each Office theme at every whole tint
	// and shade, as a BGR Long (pyOpenVBA's tests/live_access_test).
	it.each([
		['theme_tints_2007.csv', OFFICE_2007],
		['theme_tints_2023.csv', OFFICE_2023_THEME],
	])('reproduces every colour in %s', (file, theme) => {
		const wrong: string[] = [];
		let rows = 0;
		for (const line of fs.readFileSync(path.join(FIXTURES, 'access', file), 'utf8').split(/\r?\n/)) {
			if (!line.trim()) { continue; }
			const [slot, kind, amount, bgr] = line.split(',');
			const value = Number(bgr);
			const expected = [value & 0xff, (value >> 8) & 0xff, (value >> 16) & 0xff]
				.map((b) => b.toString(16).toUpperCase().padStart(2, '0')).join('');
			const color = theme.colors[Number(slot)];
			const got = kind === 'tint' ? accessTintedColor(color, Number(amount)) : accessTintedColor(color, 100, Number(amount));
			rows++;
			if (got !== expected) { wrong.push(`${slot} ${kind} ${amount}: ${got}, Access ${expected}`); }
		}
		expect(rows).toBe(2424);
		expect(wrong).toEqual([]);
	});
});

describe('the captured templates and the database\'s theme (issue #151)', () => {
	it('gives the templates back byte for byte on the 2023 theme they were captured in', () => {
		for (const kind of ['form', 'report'] as const) {
			const template = accessDesignTemplate(kind);
			expect(rethemedAccessDesign(template.blob, OFFICE_2023_THEME).equals(template.blob)).toBe(true);
			const again = rethemedAccessPrototypes(template.prototypes, OFFICE_2023_THEME);
			for (const [type, records] of template.prototypes) {
				expect(again.get(type), `${kind} ${CONTROL_TYPES.get(type)}`).toEqual(records);
			}
		}
	});

	it('draws a new form in the 2007 theme\'s faces and colours in a database on it', () => {
		const { blob } = withDesignGuid('form', Buffer.alloc(16, 7), OFFICE_2007);
		const own = parseAccessDesign(blob).objects[0].records;
		expect(own.find((r) => r.code === 160)?.value.toString('utf16le')).toBe('Calibri');
		// Calibri takes no family byte; Background 2 is lt2.
		expect(own.some((r) => r.code === 244)).toBe(false);
		expect(own.find((r) => r.code === 319)?.value.toString('hex')).toBe('eeece100');

		const prototypes = availablePrototypes('form', new Map(), OFFICE_2007);
		const faces = new Set<string>();
		for (const records of prototypes.values()) {
			const font = records.find((r) => r.code === 34);
			if (!font) { continue; }
			const face = font.value.toString('utf16le');
			faces.add(face);
			const family = records.find((r) => r.code === 243);
			// Cambria's family byte is 18; Calibri writes none.
			expect(family?.value[0], face).toBe(face === 'Cambria' ? 18 : undefined);
		}
		expect([...faces].sort()).toEqual(['Calibri', 'Cambria']);
	});

	it('reads a database\'s theme from its resources, and answers 2023 for one with none', () => {
		const themed = readAccessDesignTheme(fs.readFileSync(path.join(FIXTURES, 'binaries', 'AccessFormFixture.accdb')));
		// Read from the attachment, not the fallback: a new object, equal to it.
		expect(themed).not.toBe(OFFICE_2023_THEME);
		expect(themed).toEqual(OFFICE_2023_THEME);
		expect(readAccessDesignTheme(fs.readFileSync(path.join(FIXTURES, 'binaries', 'AccessFixture.accdb')))).toBe(OFFICE_2023_THEME);
	});

	it('leaves a new form in a database on the 2023 theme exactly as the template has it', () => {
		const writer = new AccessVbaWriter(fs.readFileSync(path.join(FIXTURES, 'binaries', 'AccessFormFixture.accdb')));
		const expected = withDesignGuid('form', Buffer.alloc(16, 3)).blob;
		expect(withDesignGuid('form', Buffer.alloc(16, 3), readAccessDesignTheme(writer.toBuffer())).blob).toEqual(expected);
	});
});
