// A fresh UserForm's designer storage and module header, composed natively.
//
// Everything here mirrors what live Excel wrote into FormFixture.xlsm, byte
// shape by byte shape: a minimal FormControl whose mask carries exactly
// NextAvailableID, DisplayedSize, LogicalSize and DrawBuffer; an empty object
// stream (the spec requires it to exist even empty); the textual VBFrame with
// its twips client box; the constant Forms 2.0 CompObj; and the exported
// header whose VB_Base carries two fresh GUIDs.

import { randomBytes } from 'crypto';
import { OformsWriter, pointsToHimetric } from './bytes';

/** The Forms 2.0 CompObj stream for a top-level form, as Excel writes it. */
export const FORMS20_COMPOBJ: Buffer = Buffer.from(
	'0100feff030a0000ffffffff00000000000000000000000000000000'
	+ '190000004d6963726f736f667420466f726d7320322e3020466f726d'
	+ '0010000000456d626564646564204f626a6563740000000000f439b271'
	+ '000000000000000000000000',
	'hex',
);

// A container's CompObj names the kind fm20 should bind the storage as - a
// Page's says Forms.Form.1 under the Form CLSID, a Frame's Forms.Frame.1 -
// and binding fails silently with the wrong one: a page authored with its
// parent MultiPage's CompObj loaded without erroring and simply did not
// appear in Pages. Bytes verbatim from the Excel-authored fixture.

/** CompObj for a MultiPage Page's storage. */
export const PAGE_COMPOBJ: Buffer = Buffer.from(
	'0100feff030a0000fffffffff0692ac6dc16ce119e9800aa00574a4f'
	+ '190000004d6963726f736f667420466f726d7320322e3020466f726d'
	+ '0010000000456d626564646564204f626a656374000d000000466f72'
	+ '6d732e466f726d2e3100f439b271000000000000000000000000',
	'hex',
);

/** CompObj for a Frame's storage. */
export const FRAME_COMPOBJ: Buffer = Buffer.from(
	'0100feff030a0000ffffffff2020186e60f4ce119bcd00aa00608e01'
	+ '1a0000004d6963726f736f667420466f726d7320322e30204672616d'
	+ '650010000000456d626564646564204f626a656374000e000000466f'
	+ '726d732e4672616d652e3100f439b271000000000000000000000000',
	'hex',
);

const USERFORM_CLSID = '{C62A69F0-16DC-11CE-9E98-00AA00574A4F}';

function newGuid(): string {
	const b = randomBytes(16);
	b[6] = (b[6] & 0x0f) | 0x40;
	b[8] = (b[8] & 0x3f) | 0x80;
	const hex = b.toString('hex').toUpperCase();
	return `{${hex.slice(0, 8)}-${hex.slice(8, 12)}-${hex.slice(12, 16)}-${hex.slice(16, 20)}-${hex.slice(20)}}`;
}

export interface NewFormOptions {
	name: string;
	caption?: string;
	/** Client area, points. Defaults match the VBE's new-form size. */
	widthPt?: number;
	heightPt?: number;
}

export interface NewFormStreams {
	f: Buffer;
	o: Buffer;
	vbFrame: string;
	compObj: Buffer;
	/** The module's attribute header, ready to sit above the code. */
	header: string;
}

export function composeNewForm(options: NewFormOptions): NewFormStreams {
	const widthPt = options.widthPt ?? 240;
	const heightPt = options.heightPt ?? 180;
	const caption = options.caption ?? options.name;

	// FormControl: header, mask, DataBlock (NextAvailableID, DrawBuffer),
	// ExtraDataBlock (DisplayedSize, LogicalSize).
	const w = new OformsWriter();
	const width = pointsToHimetric(widthPt);
	const height = pointsToHimetric(heightPt);
	const mask = (1 << 3) | (1 << 10) | (1 << 11) | (1 << 27);
	const body = new OformsWriter();
	// NextAvailableID records the LAST id handed out (Excel leaves it equal
	// to the highest live control ID), so an untouched form says zero.
	body.u32(0);
	body.u32(32000);   // DrawBuffer, the value Excel always writes
	body.i32(width); body.i32(height);   // DisplayedSize
	body.i32(width); body.i32(height);   // LogicalSize
	const bodyBytes = body.toBuffer();
	w.u8(0x00);
	w.u8(0x04);
	w.u16(bodyBytes.length + 4);
	w.u32(mask >>> 0);
	w.bytes(bodyBytes);
	// FormSiteData. The empty class-table COUNT WORD must be here: with
	// BooleanProperties defaulted, DONTSAVECLASSTABLE is 0 and fm20 reads a
	// count before CountOfSites - omit it and the low bytes of CountOfSites
	// are read AS the count. An empty form survives that by luck (both are
	// zero); the form's first control makes the misread count 1, fm20 parses
	// garbage as class info, and the whole form refuses to load. Real
	// Excel-authored root forms all carry the empty word.
	w.u16(0); // class table: zero SiteClassInfo entries
	w.u32(0); // CountOfSites
	w.u32(0); // CountOfBytes
	const f = w.toBuffer();

	const clientWidthTwips = Math.round(widthPt * 20);
	const clientHeightTwips = Math.round(heightPt * 20);
	const vbFrame = [
		'VERSION 5.00',
		`Begin ${USERFORM_CLSID} ${options.name} `,
		`   Caption         =   "${caption}"`,
		`   ClientHeight    =   ${clientHeightTwips}`,
		'   ClientLeft      =   120',
		'   ClientTop       =   465',
		`   ClientWidth     =   ${clientWidthTwips}`,
		"   StartUpPosition =   1  'CenterOwner",
		'End',
		'',
	].join('\r\n');

	const header = [
		`Attribute VB_Name = "${options.name}"`,
		`Attribute VB_Base = "0${newGuid()}${newGuid()}"`,
		'Attribute VB_GlobalNameSpace = False',
		'Attribute VB_Creatable = False',
		'Attribute VB_PredeclaredId = True',
		'Attribute VB_Exposed = False',
		'Attribute VB_TemplateDerived = False',
		'Attribute VB_Customizable = False',
		'',
	].join('\r\n');

	return { f, o: Buffer.alloc(0), vbFrame, compObj: FORMS20_COMPOBJ, header };
}
