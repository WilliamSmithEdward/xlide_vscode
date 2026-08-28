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

/** The Forms 2.0 CompObj stream, byte-for-byte as Excel writes it. */
export const FORMS20_COMPOBJ: Buffer = Buffer.from(
	'0100feff030a0000ffffffff00000000000000000000000000000000'
	+ '190000004d6963726f736f667420466f726d7320322e3020466f726d'
	+ '0010000000456d626564646564204f626a6563740000000000f439b271'
	+ '000000000000000000000000',
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
	body.u32(1);       // NextAvailableID
	body.u32(32000);   // DrawBuffer, the value Excel always writes
	body.i32(width); body.i32(height);   // DisplayedSize
	body.i32(width); body.i32(height);   // LogicalSize
	const bodyBytes = body.toBuffer();
	w.u8(0x00);
	w.u8(0x04);
	w.u16(bodyBytes.length + 4);
	w.u32(mask >>> 0);
	w.bytes(bodyBytes);
	// FormSiteData: no class table, no sites.
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
