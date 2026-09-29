/**
 * A database's Office theme, and the colours and fonts it gives a design
 * (issue #151, from pyOpenVBA 6.3.0).
 *
 * Access draws a form or report in the theme its database keeps: the `.thmx`
 * package in the MSysResources row that MSysDb's `Theme Resource Name`
 * property names. A design stores what the theme gives it, a font face beside
 * the theme font it came from and a colour beside the theme slot, tint and
 * shade it was worked out from, so the same new form differs between
 * databases whose themes differ. A database with no theme yet gets Office's
 * 2023 theme, which Access installs with its first form or report
 * (measured: Access 16.0 over COM, a blank database's first CreateForm).
 */

import { containerCodec } from '../containerCodec';
import { ZipArchive } from '../zip';
import { MSYS_OBJECTS_PAGE, readLongValue, readTableDefinition } from './accessFormat';
import { readAccessPropertyBlob } from './accessPropertyBlob';
import { readAccessCatalog, readTableRows } from './accessStorage';

/** What a design takes from its database's theme. */
export interface AccessTheme {
	majorFont: string;
	minorFont: string;
	/** `RRGGBB` for each slot, in `ACCESS_THEME_SLOTS` order. */
	colors: readonly string[];
}

/** Access's theme colour slots, in the order its ThemeColorIndex counts them. */
export const ACCESS_THEME_SLOTS = [
	'dk1', 'lt1', 'dk2', 'lt2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6',
	'hlink', 'folHlink',
] as const;

/**
 * Office's 2023 theme. It is what Access 16 installs in a database with its
 * first form or report, and the theme XLIDE's captured design templates were
 * made in (`assets/access/`), so a design for a database on it needs no
 * change.
 */
export const OFFICE_2023_THEME: AccessTheme = {
	majorFont: 'Aptos Display',
	minorFont: 'Aptos',
	colors: [
		'000000', 'FFFFFF', '0E2841', 'E8E8E8', '156082', 'E97132', '196B24', '0F9ED5', 'A02B93', '4EA72E',
		'467886', '96607D',
	],
};

export function sameAccessTheme(a: AccessTheme, b: AccessTheme): boolean {
	return a.majorFont === b.majorFont && a.minorFont === b.minorFont
		&& a.colors.length === b.colors.length && a.colors.every((color, at) => color === b.colors[at]);
}

/** The fonts and colours of a `.thmx` package; undefined when it has no usable scheme. */
export function parseAccessTheme(thmx: Buffer): AccessTheme | undefined {
	let xml: string;
	try {
		const zip = ZipArchive.read(thmx);
		const name = zip.names().find((entry) => /^theme\/theme\/theme\d+\.xml$/.test(entry));
		if (!name) {
			return undefined;
		}
		xml = zip.read(name).toString('utf8');
	} catch {
		return undefined;
	}
	const scheme = /<(?:\w+:)?clrScheme\b[\s\S]*?<\/(?:\w+:)?clrScheme>/.exec(xml)?.[0];
	const fonts = /<(?:\w+:)?fontScheme\b[\s\S]*?<\/(?:\w+:)?fontScheme>/.exec(xml)?.[0];
	if (!scheme || !fonts) {
		return undefined;
	}
	const colors: string[] = [];
	for (const slot of ACCESS_THEME_SLOTS) {
		const entry = new RegExp(`<(?:\\w+:)?${slot}>\\s*<(?:\\w+:)?(srgbClr|sysClr)\\b([^>]*)>`).exec(scheme);
		const attribute = entry?.[1] === 'sysClr' ? 'lastClr' : 'val';
		const value = entry ? new RegExp(`\\b${attribute}="([0-9A-Fa-f]{6})"`).exec(entry[2])?.[1] : undefined;
		if (!value) {
			return undefined;
		}
		colors.push(value.toUpperCase());
	}
	const face = (kind: string): string | undefined =>
		new RegExp(`<(?:\\w+:)?${kind}>\\s*<(?:\\w+:)?latin\\b[^>]*\\btypeface="([^"]+)"`).exec(fonts)?.[1];
	const majorFont = face('majorFont');
	const minorFont = face('minorFont');
	if (!majorFont || !minorFont) {
		return undefined;
	}
	return { majorFont, minorFont, colors };
}

/** A long-value or inline column value as bytes. */
function bytesOf(data: Buffer, value: unknown): Buffer | undefined {
	if (Buffer.isBuffer(value)) {
		return value;
	}
	if (value && typeof value === 'object' && 'kind' in value && (value as { kind: string }).kind === 'longValue') {
		return readLongValue(data, value as never);
	}
	return undefined;
}

/**
 * An attachment's file. `FileData` is <u32 1 when deflated> <u32 size> and the
 * payload; the payload opens with <u32 header size> and the extension, and
 * the file follows the header.
 */
function attachmentFile(stored: Buffer): Buffer | undefined {
	if (stored.length < 8) {
		return undefined;
	}
	const payload = stored.readUInt32LE(0) === 1 ? containerCodec().inflate(stored.subarray(8)) : stored.subarray(8);
	if (payload.length < 4) {
		return undefined;
	}
	const header = payload.readUInt32LE(0);
	return header > 0 && header < payload.length ? payload.subarray(header) : undefined;
}

/**
 * The theme a database's forms and reports are drawn in: the `.thmx` in
 * MSysResources that MSysDb's `Theme Resource Name` names. A database with
 * none yet, or one whose theme cannot be read, answers Office's 2023 theme,
 * which Access installs with its first form or report.
 */
export function readAccessDesignTheme(data: Buffer): AccessTheme {
	try {
		const objects = readTableRows(data, readTableDefinition(data, MSYS_OBJECTS_PAGE));
		const msysDb = objects.find((row) => row.values.get('Name') === 'MSysDb');
		const properties = msysDb ? bytesOf(data, msysDb.values.get('LvProp')) : undefined;
		const themeName = properties ? readAccessPropertyBlob(properties)?.get('Theme Resource Name') : undefined;
		if (typeof themeName !== 'string') {
			return OFFICE_2023_THEME;
		}
		const catalog = readAccessCatalog(data);
		const resources = catalog.find((table) => table.name === 'MSysResources');
		if (!resources) {
			return OFFICE_2023_THEME;
		}
		const row = readTableRows(data, readTableDefinition(data, resources.definitionPage)).find((entry) =>
			entry.values.get('Name') === themeName && String(entry.values.get('Type')).toLowerCase() === 'thmx');
		const key = row ? bytesOf(data, row.values.get('Data')) : undefined;
		if (!key || key.length < 4) {
			return OFFICE_2023_THEME;
		}
		const id = key.readInt32LE(0);
		// The attachment rows sit in a hidden table whose foreign-key column
		// is named for the resources table.
		for (const table of catalog) {
			if (!/^f_[0-9A-F]{32}_Data$/i.test(table.name)) {
				continue;
			}
			for (const file of readTableRows(data, readTableDefinition(data, table.definitionPage))) {
				if (file.values.get('MSysResources_Data') !== id) {
					continue;
				}
				const stored = bytesOf(data, file.values.get('FileData'));
				const thmx = stored ? attachmentFile(stored) : undefined;
				const theme = thmx ? parseAccessTheme(thmx) : undefined;
				if (theme) {
					return theme;
				}
			}
		}
	} catch {
		// A theme that cannot be read draws in the theme Access installs.
	}
	return OFFICE_2023_THEME;
}

// --- theme colours ---------------------------------------------------------------
// A tint lightens a slot's colour and a shade darkens it, both percentages, 100
// leaving it alone: in HSL, luminance goes to L * t + (1 - t) for a tint and to
// L * s for a shade, each channel rounding half up. That gives every colour
// Access answered for each slot of both Office themes at every whole tint and
// shade, 4,444 in all, but for the 148 where a channel lands on a half and
// Access rounds as its own arithmetic falls. Those are carried as measured,
// keyed `RRGGBB` + `t` or `s` + the percentage in three digits. They correct
// exactly the float path below, which is colorsys's, step for step as Python
// 3.11 and later take it; a port that reorders it moves other halves.

const MEASURED = [
	'000000t010E5E5E5 000000t0507F7F7F 0000FFs010000019 0000FFs05000007F 0000FFs0900000E5 0000FFt010E5E5FF',
	'0000FFt0507F7FFF 0000FFt0901919FF 0E2841s010010406 0E2841s025030A10 0E2841s050071420 0E2841s0700A1C2E',
	'0E2841t0543F8CD6 0F9ED5t0674FC5F3 156082s005010506 156082s025051820 156082s03507222E 156082s0550C3547',
	'156082s08512526F 156082s095145B7C 156082t095186B91 156082t099166285 196B24s002000201 196B24s014030F05',
	'196B24s026061C09 196B24s03809290E 196B24s0500D3512 196B24s0580E3E15 196B24s066104718 196B24s070114B19',
	'196B24s07813531C 196B24s08214581E 196B24s090166020 196B24s098186923 196B24t03987E394 1F497Ds006020408',
	'1F497Ds014040A11 1F497Ds02207101B 1F497Ds026081320 1F497Ds030091626 1F497Ds0340B192A 1F497Ds0420D1F35',
	'1F497Ds05010253F 1F497Ds054112743 1F497Ds058122A48 1F497Ds082193C67 1F497Ds0861B3F6C 1F497Ds0981E487B',
	'467886s005030607 467886s0150B1214 467886s025111E22 467886s035192A2F 467886s0451F363C 467886s05526424A',
	'467886s075355A64 467886s09542727F 467886t010EBF2F4 467886t030C3D9DF 467886t0509BC0CA 467886t07072A7B5',
	'4BACC6t050A5D6E2 4EA72Es025132A0C 4EA72Es050275317 4F81BDs013091119 800080t025FF9FFF 800080t075DF00DF',
	'8064A2t050BFB2D0 8064A2t070A692BE 8064A2t0908D74AB 96607Ds002030203 96607Ds006090608 96607Ds0100F0A0D',
	'96607Ds013130C10 96607Ds015160E13 96607Ds017191015 96607Ds0211F141A 96607Ds027281A22 96607Ds0292C1C24',
	'96607Ds0302D1D26 96607Ds0312E1E27 96607Ds033312029 96607Ds03737242E 96607Ds038392430 96607Ds0413E2733',
	'96607Ds0423F2835 96607Ds043402936 96607Ds045432B38 96607Ds046452C3A 96607Ds047462D3B 96607Ds0504B303F',
	'96607Ds057563747 96607Ds0615C3B4C 96607Ds0625D3C4E 96607Ds066633F53 96607Ds067644054 96607Ds069684256',
	'96607Ds0736D465B 96607Ds0746F475D 96607Ds078754B62 96607Ds0827B4F67 96607Ds0837C5068 96607Ds08681536C',
	'96607Ds08782546D 96607Ds090875671 96607Ds0938B5974 96607Ds0958F5B77 96607Ds097915D79 96607Ds098935E7B',
	'96607Ds099955F7C 9BBB59s01213180A 9BBB59t075B4CC82 A02B93s05050164A C0504Ds032401816 C0504Dt002FEFBFB',
	'C0504Dt014F6E6E6 C0504Dt018F4DFDF C0504Dt022F1D8D8 C0504Dt025EFD3D2 C0504Dt026EFD1D1 C0504Dt030ECCBCA',
	'C0504Dt038E7BDBB C0504Dt046E2AEAD C0504Dt050DFA8A6 C0504Dt062D89291 C0504Dt075D07C79 C0504Dt078CE7774',
	'C0504Dt086C96866 C0504Dt090C6615F C0504Dt094C45B58 E8E8E8t050F4F4F4 E97132s020331506 E97132t010FDF1EA',
	'E97132t025FADCCC E97132t075EF9465 EEECE1t025FBFAF7 EEECE1t050F6F5F0 EEECE1t055F6F5EF EEECE1t075F2F1E8',
	'EEECE1t095EFEDE2 F79646s013271302 F79646s0193A1B03 F79646s087F5801F F79646t010FEF4EC F79646t030FDDFC8',
	'F79646t090F8A059 FFFFFFs010191919 FFFFFFs0507F7F7F FFFFFFs090E5E5E5',
].join(' ');

const CORRECTIONS: ReadonlyMap<string, string> = new Map(
	MEASURED.split(' ').map((entry) => [entry.slice(0, 10), entry.slice(10)]),
);

/** `color` (`RRGGBB`) lightened by `tint` or darkened by `shade`, as Access works it out. */
export function accessTintedColor(color: string, tint = 100, shade = 100): string {
	const whole = shade === 100 ? { value: tint, letter: 't' } : tint === 100 ? { value: shade, letter: 's' } : undefined;
	if (whole && Number.isInteger(whole.value)) {
		const found = CORRECTIONS.get(`${color.toUpperCase()}${whole.letter}${String(whole.value).padStart(3, '0')}`);
		if (found) {
			return found;
		}
	}
	const [red, green, blue] = [0, 2, 4].map((at) => Number.parseInt(color.slice(at, at + 2), 16) / 255);
	const [hue, lightness, saturation] = toHls(red, green, blue);
	let luminance = lightness;
	// In this order: the table above corrects exactly this arithmetic.
	if (tint !== 100) {
		const fraction = tint / 100;
		luminance = luminance * fraction + (1 - fraction);
	}
	if (shade !== 100) {
		luminance = luminance * (shade / 100);
	}
	return fromHls(hue, luminance, saturation)
		.map((channel) => Math.floor(channel * 255 + 0.5).toString(16).toUpperCase().padStart(2, '0'))
		.join('');
}

/** A slot's colour as a design record holds it: red, green, blue, 0. */
export function accessThemeColor(theme: AccessTheme, index: number, tint = 100, shade = 100): Buffer {
	return Buffer.concat([Buffer.from(accessTintedColor(theme.colors[index], tint, shade), 'hex'), Buffer.alloc(1)]);
}

/** Python's float `x % 1.0`: `fmod`, moved into [0, 1) when negative. */
function modOne(value: number): number {
	const rest = value % 1;
	return rest < 0 ? rest + 1 : rest;
}

const ONE_THIRD = 1 / 3;
const ONE_SIXTH = 1 / 6;
const TWO_THIRDS = 2 / 3;

function toHls(red: number, green: number, blue: number): [number, number, number] {
	const high = Math.max(red, green, blue);
	const low = Math.min(red, green, blue);
	const total = high + low;
	const spread = high - low;
	const luminance = total / 2;
	if (low === high) {
		return [0, luminance, 0];
	}
	const saturation = luminance <= 0.5 ? spread / total : spread / (2 - high - low);
	const fromRed = (high - red) / spread;
	const fromGreen = (high - green) / spread;
	const fromBlue = (high - blue) / spread;
	let hue: number;
	if (red === high) {
		hue = fromBlue - fromGreen;
	} else if (green === high) {
		hue = 2 + fromRed - fromBlue;
	} else {
		hue = 4 + fromGreen - fromRed;
	}
	return [modOne(hue / 6), luminance, saturation];
}

function fromHls(hue: number, luminance: number, saturation: number): [number, number, number] {
	if (saturation === 0) {
		return [luminance, luminance, luminance];
	}
	const top = luminance <= 0.5 ? luminance * (1 + saturation) : luminance + saturation - (luminance * saturation);
	const bottom = 2 * luminance - top;
	return [channel(bottom, top, hue + ONE_THIRD), channel(bottom, top, hue), channel(bottom, top, hue - ONE_THIRD)];
}

function channel(bottom: number, top: number, hueIn: number): number {
	const hue = modOne(hueIn);
	if (hue < ONE_SIXTH) {
		return bottom + (top - bottom) * hue * 6;
	}
	if (hue < 0.5) {
		return top;
	}
	if (hue < TWO_THIRDS) {
		return bottom + (top - bottom) * (TWO_THIRDS - hue) * 6;
	}
	return bottom;
}
